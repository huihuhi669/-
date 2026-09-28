"""
故障报告自动生成系统 - 后端服务
基于 Flask 框架，整合智谱AI大模型进行电力故障分析报告生成
"""

import os
import re
import json
import time
import base64
import threading
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# 文档处理
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# AI 大模型
from zhipuai import ZhipuAI
import requests

# ============================================================
# 初始化 Flask 应用
# ============================================================
app = Flask(__name__)
CORS(app)

# ============================================================
# 配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads', 'template1')
REPORTS_DIR = os.path.join(BASE_DIR, 'generated_reports', 'template1')
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 上传限制
ALLOWED_EXTENSIONS = {'txt', 'docx', 'doc', 'pdf', 'png', 'jpg', 'jpeg', 'xlsx', 'xls', 'csv', 'json'}

# 智谱AI 配置
ZHIPU_API_KEY = "90ad75971b4f41caa79207fac0abeeea.oPTttiiw0joYpArx"
client = ZhipuAI(api_key=ZHIPU_API_KEY)

# 高德天气 API
AMAP_KEY = '4b98ff847848747533f76a6ca00da033'

# 任务状态存储（简单内存存储）
task_store = {}

# ============================================================
# 工具函数
# ============================================================

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def clean_text(text):
    """清理生成文本，去除不必要的符号和多余换行"""
    if not text:
        return ""
    cleaned = text.replace("\r\n", "\n")
    cleaned = cleaned.replace("\n\n", "\n")
    cleaned = cleaned.replace("-", "")
    cleaned = cleaned.replace("*", "")
    cleaned = cleaned.strip()
    return cleaned


def get_uploaded_files(category: str, uploads_root: Path) -> list:
    """扫描 uploads_root/<category> 目录下的所有文件"""
    dir_path = uploads_root / category
    if not dir_path.exists() or not dir_path.is_dir():
        raise FileNotFoundError(f"上传目录不存在：{dir_path}")
    files = [p for p in dir_path.iterdir() if p.is_file()]
    if not files:
        raise FileNotFoundError(f"目录 {category} 下没有找到任何上传文件")
    return sorted(files)


def fetch_current_weather(location, key):
    """获取实时天气信息"""
    api_url = f"https://restapi.amap.com/v3/weather/weatherInfo?city={location}&key={key}&extensions=base"
    try:
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()
        data = response.json()
        if data['status'] == '1':
            lives = data.get('lives', [])[0]
            temperature = lives['temperature']
            weather = lives['weather']
            report_time = lives['reporttime']
            formatted_time = (
                report_time.replace("-", "年", 1)
                .replace("-", "月", 1)
                .replace(" ", "日", 1)
                .replace(":", "时", 1)
                .replace(":", "分", 1) + "秒"
            )
            return f"气温：{temperature}℃，天气：{weather}", formatted_time
        else:
            return f"API请求失败: {data['info']}", None
    except Exception as e:
        return f"请求失败: {e}", None


def analyze_image_with_glm(image_base64, context_info=None):
    """调用GLM-4V-Plus模型分析录波图/时序图"""
    url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {ZHIPU_API_KEY}"
    }

    if context_info:
        text_prompt = (
            f"请结合前面所述的事件简述、基本情况、事件经过分析，"
            f"对这张录波图进行电气专业的分析和解释，解读这个录波图的特点，"
            f"包括发生故障的相别电压，故障相别的电流，有零序电流和零序电压;"
            f"非故障相的电流、电压;零序电流与故障相电流；"
            f"零序电流与故障相电流的方向，超前或滞后零序电压，故障相电流超前或滞后故障相电压多少。"
            f"其他的输出内容需包括：是否为区内故障，故障发展过程，各个过程保护动作行为，"
            f"故障电流大小，故障切除时间，测距情况，非电量保护动作情况等，"
            f"并分析是否发生不正确动作。以下是上下文信息：{context_info}"
            f"注意：输出的内容不要有特殊符号和多余的回车和空格，包括[-][*]，"
            f"若是分点输出的话，使用括号分点，如（1）（2）（3）。"
        )
    else:
        text_prompt = (
            "请对二次系统动作时序图进行电气专业的分析，包括具体描述和各个时间点发生的情况，"
            "图中有时间点的话要在文字中体现"
            "注意：输出的内容不要有特殊符号和多余的回车和空格，包括[-][*]。"
            "若是分点输出的话，使用括号分点，如（1）（2）（3）。"
        )

    data = {
        "model": "glm-4v-plus",
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": image_base64}},
                {"type": "text", "text": text_prompt}
            ]
        }]
    }

    response = requests.post(url, json=data, headers=headers, timeout=120)
    raw_text = response.json()['choices'][0]['message']['content']
    return clean_text(raw_text)


def read_document_text(file_path):
    """读取单个故障文档的全文（支持 txt / docx / csv / json / xlsx）"""
    ext = Path(file_path).suffix.lower()
    if ext == '.txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    if ext == '.docx':
        doc = Document(file_path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    if ext == '.csv':
        import csv
        lines = []
        with open(file_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
            for row in csv.reader(f):
                lines.append(" | ".join(row))
        return "\n".join(lines)
    if ext == '.json':
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    if ext == '.xlsx':
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"【工作表：{ws.title}】")
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) if c is not None else '' for c in row]
                if any(cells):
                    lines.append(" | ".join(cells))
        wb.close()
        return "\n".join(lines)
    if ext == '.xls':
        raise ValueError(".xls 为旧版格式，请另存为 .xlsx 后上传")
    if ext == '.pdf':
        raise ValueError("PDF 解析需要额外安装 pypdf 库，请改用其他格式")
    raise ValueError(f"不支持的文件类型：{ext}")


def extract_event_fields(text):
    """调用智谱AI，从故障文档全文抽取事件简述字段，返回 dict"""
    prompt = (
        "你是电力系统故障分析专家。请从下面的故障文档中提取关键信息，"
        "返回一个严格的 JSON 对象，键名固定为以下字段（文档里没有的字段值填空字符串）："
        "事件时间、厂站、电压等级、设备、跳闸情况、重合闸动作情况、故障线路、故障相、"
        "故障、保护动作情况、保护故障测距、一次设备情况、二次设备情况、"
        "事件地区气象情况、事件相关单位概况。"
        "只输出 JSON 本身，不要任何解释、不要 markdown 代码块标记。\n\n"
        f"故障文档内容：\n{text}"
    )
    response = client.chat.completions.create(
        model="glm-4-plus",
        messages=[
            {"role": "system", "content": "你是专业的电力故障分析助手，只输出 JSON。"},
            {"role": "user", "content": prompt},
        ],
    )
    raw = response.choices[0].message.content.strip()
    # 容错：去掉可能的 ```json ... ``` 包裹
    if raw.startswith("```"):
        raw = raw.strip("`").strip()
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
    return json.loads(raw)


# ============================================================
# 报告生成相关函数
# ============================================================

def read_data_from_event_summary(file_path):
    """解析事件简述文件"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    data = {}
    for line in lines:
        if '：' in line:
            key, value = line.split('：', 1)
            data[key.strip()] = value.strip()
    return (
        data.get('事件时间'), data.get('厂站'), data.get('电压等级'), data.get('设备'),
        data.get('跳闸情况'), data.get('重合闸动作情况'), data.get('故障线路'), data.get('故障相'),
        data.get('故障'), data.get('保护动作情况'), data.get('保护故障测距'),
        data.get('一次设备情况'), data.get('二次设备情况'),
        data.get('事件地区气象情况'), data.get('事件相关单位概况'),
    )


def read_data_from_event_phase(file_path: str) -> dict:
    """读取事件前/后运行方式 TXT 文件"""
    data = {}
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            if '：' in line:
                key, val = line.split('：', 1)
                data[key.strip()] = val.strip()
    return data


def read_data_from_equipment_basic_info(file_path):
    """解析一次设备基本信息"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    data = {}
    for line in lines:
        if '：' in line:
            key, val = line.split('：', 1)
            data[key.strip()] = val.strip()
    return (
        data.get('型号'), data.get('设备厂家'),
        data.get('生产年份'), data.get('投产年份'),
    )


def read_data_from_equipment_maintenance_info(file_path):
    """解析一次设备维护信息"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    data = {}
    for line in lines:
        if '：' in line:
            key, val = line.split('：', 1)
            data[key.strip()] = val.strip()
    return (
        data.get('上次预试时间'), data.get('一次设备维护情况'),
        data.get('上次预试时间及内容'), data.get('上次维护时间及内容'),
    )


def read_data_from_secondary_equipment_info(file_path):
    """解析二次设备信息"""
    equipment_info = []
    with open(file_path, 'r', encoding='utf-8') as f:
        raw = f.read()
    content = raw.replace('\r\n', '\n').strip()
    blocks = [blk.strip() for blk in content.split('\n\n') if blk.strip()]
    for blk in blocks:
        data = {}
        for line in blk.splitlines():
            if '：' in line:
                key, val = line.split('：', 1)
                data[key.strip()] = val.strip()
        model = data.get('型号')
        interval = data.get('间隔名称')
        if model and interval:
            equipment_info.append({
                "间隔名称": interval,
                "保护型号": model,
                "设备厂家": data.get('设备厂家'),
                "生产年份": data.get('生产年份'),
                "投产年份": data.get('投产年份')
            })
    return equipment_info


def read_secondary_device_maintenance_info(file_path):
    """解析二次设备维护信息"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    data = {}
    for line in lines:
        if '：' in line:
            key, val = line.split('：', 1)
            data[key.strip()] = val.strip()
    return (
        data.get('二次设备维护情况'),
        data.get('上次预试时间及内容'),
        data.get('上次维护时间及内容')
    )


def _set_run_font(run, name, size, bold=False):
    """设置 run 的中英文字体、字号、加粗"""
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size)
    run.bold = bold


def _add_heading(doc, text, level=1):
    """添加章节标题"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    if level == 1:
        _set_run_font(run, '黑体', 16, bold=True)
    elif level == 2:
        _set_run_font(run, '黑体', 15, bold=True)
    else:
        _set_run_font(run, '仿宋', 14, bold=True)
    p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    return p


def _add_body(doc, text):
    """添加正文段落（自动按换行拆分）"""
    text = (text or "").strip()
    if not text:
        text = "（未提供）"
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        p = doc.add_paragraph()
        run = p.add_run(line)
        _set_run_font(run, '仿宋', 14)
        p.paragraph_format.first_line_indent = Pt(24)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)


def _add_kv_table(doc, rows):
    """添加两列键值表（项目 | 详情），带边框"""
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    for key, value in rows:
        cells = table.add_row().cells
        kp = cells[0].paragraphs[0]
        krun = kp.add_run(key)
        _set_run_font(krun, '黑体', 12, bold=True)
        vp = cells[1].paragraphs[0]
        vrun = vp.add_run(value)
        _set_run_font(vrun, '仿宋', 12)
        cells[0].width = Inches(1.6)
        cells[1].width = Inches(4.6)
    return table


def generate_report_from_fields_sync(task_id, fields, reports_dir):
    """根据 AI 抽取的字段，用固定模板渲染专业报告（无额外 AI 调用，稳定、快速）"""
    def fmt(v):
        if v is None:
            return "（未提供）"
        s = str(v).strip()
        return s if s else "（未提供）"

    try:
        task_store[task_id]['status'] = 'processing'
        task_store[task_id]['progress'] = 10
        task_store[task_id]['message'] = "正在套用报告模板..."

        # 字段归一化（去掉键名可能存在的空格），并解包
        nf = {str(k).strip(): v for k, v in (fields or {}).items()}
        event_time = fmt(nf.get('事件时间'))
        station = fmt(nf.get('厂站'))
        voltage_level = fmt(nf.get('电压等级'))
        equipment = fmt(nf.get('设备'))
        trip_condition = fmt(nf.get('跳闸情况'))
        reclose_action = fmt(nf.get('重合闸动作情况'))
        fault_line = fmt(nf.get('故障线路'))
        fault_phase = fmt(nf.get('故障相'))
        fault = fmt(nf.get('故障'))
        protection_action = fmt(nf.get('保护动作情况'))
        protection_fault_distance = fmt(nf.get('保护故障测距'))
        primary_device_info = fmt(nf.get('一次设备情况'))
        secondary_device_info = fmt(nf.get('二次设备情况'))
        weather_info = fmt(nf.get('事件地区气象情况'))
        related_units_info = fmt(nf.get('事件相关单位概况'))

        doc = Document()

        # 标题
        title_p = doc.add_paragraph()
        title_run = title_p.add_run("电力系统故障分析报告")
        _set_run_font(title_run, '方正小标宋', 22, bold=False)
        title_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        title_p.paragraph_format.space_after = Pt(12)

        # 一、故障概述
        _add_heading(doc, "一、故障概述", 1)
        _add_kv_table(doc, [
            ("事件时间", event_time),
            ("厂站", station),
            ("电压等级", voltage_level),
            ("故障设备", equipment),
            ("故障线路", fault_line),
            ("故障相", fault_phase),
            ("故障类型", fault),
            ("跳闸情况", trip_condition),
            ("重合闸情况", reclose_action),
        ])

        # 二、故障详细信息
        _add_heading(doc, "二、故障详细信息", 1)
        _add_heading(doc, "（一）故障经过", 2)
        _add_body(doc, f"{station} {fault_line} 于 {event_time} 发生 {fault}，故障相为 {fault_phase}。"
                       f"跳闸情况：{trip_condition}。重合闸动作情况：{reclose_action}。")
        _add_heading(doc, "（二）保护动作情况", 2)
        _add_body(doc, protection_action)
        _add_heading(doc, "（三）保护故障测距", 2)
        _add_body(doc, protection_fault_distance)

        # 三、故障设备情况
        _add_heading(doc, "三、故障设备情况", 1)
        _add_heading(doc, "（一）一次设备情况", 2)
        _add_body(doc, primary_device_info)
        _add_heading(doc, "（二）二次设备情况", 2)
        _add_body(doc, secondary_device_info)

        # 四、事件地区气象情况
        _add_heading(doc, "四、事件地区气象情况", 1)
        _add_body(doc, weather_info)

        # 五、事件相关单位概况
        _add_heading(doc, "五、事件相关单位概况", 1)
        _add_body(doc, related_units_info)

        # 六、结论与建议
        _add_heading(doc, "六、结论与建议", 1)
        conclusion = (
            f"本次 {station} {fault_line} 于 {event_time} 发生 {fault}，故障相为 {fault_phase}，"
            f"保护动作情况：{protection_action}，{protection_fault_distance}。"
            f"经分析，故障原因与设备运行状态、外部环境等因素相关。"
            f"建议加强设备巡视与在线监测，落实防范措施，防止同类故障再次发生。"
        )
        _add_body(doc, conclusion)

        # 保存文档
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_id = f"REP-{timestamp}"
        output_doc = os.path.join(reports_dir, f'{report_id}.docx')
        doc.save(output_doc)

        task_store[task_id]['status'] = 'completed'
        task_store[task_id]['progress'] = 100
        task_store[task_id]['message'] = "报告生成完成！"
        task_store[task_id]['report_file'] = f'{report_id}.docx'
        task_store[task_id]['report_id'] = report_id

    except Exception as e:
        import traceback
        task_store[task_id]['status'] = 'failed'
        task_store[task_id]['error'] = str(e)
        task_store[task_id]['traceback'] = traceback.format_exc()
        print(f"[ERROR] 报告生成失败: {traceback.format_exc()}")


def generate_report_sync(task_id, uploads_dir, reports_dir):
    """同步生成报告（在后台线程中运行）"""
    try:
        task_store[task_id]['status'] = 'processing'
        task_store[task_id]['progress'] = 0
        uploads_root = Path(uploads_dir)
        total_steps = 12

        def update_progress(step):
            task_store[task_id]['progress'] = int(step / total_steps * 100)
            task_store[task_id]['message'] = f"正在生成报告... ({step}/{total_steps})"

        # === 1. 解析事件简述 ===
        update_progress(0)
        try:
            summary_files = get_uploaded_files('事件简述', uploads_root)
            event_summary_path = str(summary_files[0])
            event_inputs = read_data_from_event_summary(event_summary_path)
        except Exception as e:
            task_store[task_id]['status'] = 'failed'
            task_store[task_id]['error'] = f"解析事件简述失败: {str(e)}"
            return

        (event_time, station, voltage_level, equipment, trip_condition,
         reclose_action, fault_line, fault_phase, fault, protection_action,
         protection_fault_distance, primary_device_info, secondary_device_info,
         weather_info, related_units_info) = event_inputs

        # === 2. 解析一次设备信息 ===
        update_progress(1)
        manufacturer = production_year = commissioning_year = model = None
        equipment_maintenance_info = last_yushi_content = last_maintenance_content = None
        last_yushi_time = None
        try:
            files1 = get_uploaded_files('故障一次设备信息及维护情况', uploads_root)
            basic_f = maint_f = None
            for f in files1:
                first_line = f.open(encoding='utf-8').readline().strip()
                if first_line.startswith('型号：'):
                    basic_f = f
                else:
                    maint_f = f
            if basic_f:
                manufacturer, _, production_year, commissioning_year = read_data_from_equipment_basic_info(str(basic_f))
            if maint_f:
                last_yushi_time, equipment_maintenance_info, last_yushi_content, last_maintenance_content = \
                    read_data_from_equipment_maintenance_info(str(maint_f))
        except Exception as e:
            print(f"[WARN] 一次设备信息解析: {e}")

        # === 3. 解析二次设备信息 ===
        update_progress(2)
        secondary_equipment_info = []
        secondary_maintenance_info = secondary_last_yushi_content = secondary_last_maintenance_content = None
        try:
            files2 = get_uploaded_files('故障二次设备信息及维护情况', uploads_root)
            basic2 = maint2 = None
            for f in files2:
                first_line = f.open(encoding='utf-8').readline().strip()
                if first_line.startswith('型号：'):
                    basic2 = f
                else:
                    maint2 = f
            if basic2:
                secondary_equipment_info = read_data_from_secondary_equipment_info(str(basic2))
            if maint2:
                secondary_maintenance_info, secondary_last_yushi_content, secondary_last_maintenance_content = \
                    read_secondary_device_maintenance_info(str(maint2))
        except Exception as e:
            print(f"[WARN] 二次设备信息解析: {e}")

        # === 4. 获取天气 ===
        update_progress(3)
        station = station or ''
        real_time_weather = formatted_report_time = None
        if station:
            location = station.rstrip('站')
            real_time_weather, formatted_report_time = fetch_current_weather(location, AMAP_KEY)
        if formatted_report_time:
            event_time = formatted_report_time

        # === 5. 创建Word文档 ===
        update_progress(4)
        document = Document()

        # 标题
        title = document.add_paragraph()
        run = title.add_run("故障事件分析报告")
        run.font.name = '方正小标宋'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋')
        run.font.size = Pt(22)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # 辅助函数：添加章节
        def add_section(title_text, content=None, level=1):
            if title_text.strip():
                paragraph = document.add_paragraph()
                run = paragraph.add_run(title_text)
                if level == 1:
                    run.bold = True
                    run.font.name = '黑体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                    run.font.size = Pt(16)
                elif level == 2:
                    run.bold = True
                    run.font.name = '仿宋'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                    run.font.size = Pt(15)
                elif level == 3:
                    run.bold = True
                    run.font.name = '仿宋'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                    run.font.size = Pt(14)
                paragraph.paragraph_format.first_line_indent = Pt(24)
                paragraph.paragraph_format.line_spacing = 1.5
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)

            if content and content.strip():
                cleaned = content.strip().replace("\n\n", "\n")
                for line in cleaned.split("\n"):
                    if line.strip():
                        p = document.add_paragraph(line)
                        run = p.runs[0]
                        run.font.name = '仿宋'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                        run.font.size = Pt(14)
                        p.paragraph_format.first_line_indent = Pt(24)
                        p.paragraph_format.line_spacing = 1.5
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)

        # === 6. 一、事件简述 ===
        update_progress(5)
        if station and event_time:
            prompt = (
                f"请根据以下信息生成一段详细的、连贯的保护动作情况，不需要分析或总结，不要分点撰写："
                f"{station}变电站：{event_time} {voltage_level} {fault_line} {fault_phase}发生{fault}，"
                f"保护动作情况：{protection_action}，重合{reclose_action}，"
                f"保护故障测距{protection_fault_distance}。"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是一个专业的故障分析助手。"},
                    {"role": "user", "content": prompt}
                ],
            )
            summary = response.choices[0].message.content.strip()
        else:
            summary = "缺少[事件简述]文件，无法生成本节内容。"
        add_section("一、事件简述", summary, level=1)

        # === 7. 二、基本情况 ===
        update_progress(6)
        add_section("二、基本情况", "", level=1)

        # 事件前运行方式
        before_dir = uploads_root / "事件前运行方式"
        event_before_txt = None
        if before_dir.exists():
            files = list(before_dir.glob("*"))
            if files:
                event_before_txt = files[0]

        if event_before_txt:
            pre_data = read_data_from_event_phase(str(event_before_txt))
            pre_prompt = (
                "请根据以下信息，生成一段专业的[事件前运行方式]描述，不要输出字段名，只要连贯的叙述："
                f"厂站：{pre_data.get('厂站', '')}; 电压等级：{pre_data.get('电压等级', '')}; "
                f"设备：{pre_data.get('设备', '')}; 跳闸情况：{pre_data.get('跳闸情况', '')}; "
                f"重合闸动作情况：{pre_data.get('重合闸动作情况', '')}; "
                f"故障线路：{pre_data.get('故障线路', '')}; 故障相：{pre_data.get('故障相', '')}; "
                f"保护动作情况：{pre_data.get('保护动作情况', '')}; "
                f"保护故障测距：{pre_data.get('保护故障测距', '')};"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是专业的电力系统分析专家。"},
                    {"role": "user", "content": pre_prompt}
                ],
            )
            add_section("（一）事件前运行方式", response.choices[0].message.content.strip(), level=2)
        else:
            add_section("（一）事件前运行方式", "缺少[事件前运行方式]文件，无法生成本节内容。", level=2)

        # 事件后运行方式
        after_dir = uploads_root / "事件后运行方式"
        event_after_txt = None
        if after_dir.exists():
            files = list(after_dir.glob("*"))
            if files:
                event_after_txt = files[0]

        if event_after_txt:
            post_data = read_data_from_event_phase(str(event_after_txt))
            post_prompt = (
                "请根据以下信息，生成一段专业的[事件后运行方式]描述："
                f"厂站：{post_data.get('厂站', '')}; 设备：{post_data.get('设备', '')}; "
                f"保护装置动作：{post_data.get('保护动作情况', '')}; "
                f"应急处理措施：{post_data.get('重合闸动作情况', '')}; "
                f"系统状态：{post_data.get('系统状态', '已恢复正常')}; "
                f"故障设备状态：{post_data.get('故障设备状态', '已隔离')};"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是专业的电力系统分析专家。"},
                    {"role": "user", "content": post_prompt}
                ],
            )
            add_section("（二）事件后运行方式", response.choices[0].message.content.strip(), level=2)
        else:
            add_section("（二）事件后运行方式", "缺少[事件后运行方式]文件，无法生成本节内容。", level=2)

        # 一次设备信息
        if manufacturer:
            add_section("（三）故障一次设备信息及维护情况",
                       f"{station}发电站{equipment}，型号：{manufacturer}，生产厂家为{production_year}，"
                       f"设备{commissioning_year}年出厂产品，于{model}投运。", level=2)
            if equipment_maintenance_info:
                add_section("", f"一次设备维护情况：{equipment_maintenance_info}", level=3)
            if last_yushi_content:
                add_section("", f"上次预试时间及内容：{last_yushi_content}", level=3)
            if last_maintenance_content:
                add_section("", f"上次维护时间及内容：{last_maintenance_content}", level=3)
        else:
            add_section("（三）故障一次设备信息及维护情况", "（无一次设备信息）", level=2)

        # 二次设备信息
        add_section("（四）故障二次设备信息及维护情况", "", level=2)
        if secondary_equipment_info:
            for row in secondary_equipment_info:
                sentence = (
                    f"{station}{voltage_level}{row['间隔名称']}，型号：{row['保护型号']}，"
                    f"生产厂家为{row['设备厂家']}，设备{row['生产年份']}年出厂产品，于{row['投产年份']}投运。"
                )
                add_section("", sentence, level=3)
        else:
            add_section("", "（无二次设备信息）", level=3)

        if secondary_maintenance_info:
            add_section("", str(secondary_maintenance_info), level=3)

        # 气象情况
        if real_time_weather and formatted_report_time:
            add_section("（五）事件地区气象情况",
                       f"{formatted_report_time}，{station}当地的气象情况如下：", level=2)
            add_section("", f"{real_time_weather}。", level=3)
        else:
            add_section("（五）事件地区气象情况", "缺少气象信息，无法生成本节内容。", level=2)

        # 相关单位概况
        if station and event_time:
            location = station.rstrip('站')
            prompt = (
                f"请总结[{location}]供电局的基本情况，包括设施基础、处理故障方面的能力情况或策略，"
                f"不要分点撰写，以段落格式输出。"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是一个专业的助手，负责生成有关单位的信息总结。"},
                    {"role": "user", "content": prompt}
                ],
            )
            add_section("（六）事件相关单位概况", response.choices[0].message.content.strip(), level=2)
        else:
            add_section("（六）事件相关单位概况", "缺少相关文件，无法生成本节内容。", level=2)

        # === 8. 三、事件经过分析 ===
        update_progress(7)
        add_section("三、事件经过分析", "", level=1)
        add_section("（一）一次设备分析", "", level=2)

        if station and event_time:
            add_section("1.一次设备故障情况",
                       f"{station}变电站于{event_time} {voltage_level} {fault_line} {fault_phase}发生{fault}。", level=3)

        # 故障原因分析
        fault_analysis = ""
        if station and event_time and manufacturer:
            add_section("2.故障原因分析", "", level=3)
            prompt = (
                f"请对以下设备故障情况进行电气专业的原因分析，只需写原因分析，不要分点："
                f"{station}变电站于{event_time} {voltage_level} {fault_line} {fault_phase}发生{fault}。"
                f"型号：{manufacturer}，生产厂家为{manufacturer}，出厂于{production_year}年，投运于{commissioning_year}。"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是一个电气设备故障分析专家。"},
                    {"role": "user", "content": prompt}
                ],
            )
            fault_analysis = response.choices[0].message.content.strip()
            add_section("", fault_analysis, level=3)
        else:
            add_section("2.故障原因分析", "缺少文件，无法生成本节内容。", level=3)

        # === 9. 二次设备分析 ===
        update_progress(8)
        add_section("（二）二次设备分析", "", level=2)
        add_section("1.相关保护动作原理", "", level=3)
        protection_action_principle = ""
        try:
            manu_files = get_uploaded_files('二次设备分析', uploads_root)
            manual_doc = next((f for f in manu_files if f.suffix.lower() in {'.doc', '.docx'}), manu_files[0])
            src_doc = Document(str(manual_doc))
            protection_action_principle = "\n".join([p.text.strip() for p in src_doc.paragraphs if p.text.strip()])
            # 插入文档内容
            for para in src_doc.paragraphs:
                if para.text.strip():
                    p = document.add_paragraph(para.text)
                    for run in p.runs:
                        run.font.name = '仿宋'
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                        run.font.size = Pt(14)
                    p.paragraph_format.first_line_indent = Pt(24)
        except Exception as e:
            print(f"[WARN] 二次设备分析文档: {e}")

        add_section("2.二次系统动作时序图", "（需上传时序图图片进行分析）", level=3)
        action_timing_diagram = ""

        add_section("3.二次系统录波图分析", "（需上传录波图图片进行分析）", level=3)
        action_behavior_analysis = ""

        # 小结
        if protection_action_principle.strip() or action_timing_diagram.strip():
            prompt = (
                f"请根据以下内容生成小结：保护动作原理：{protection_action_principle}；"
                f"动作时序图分析：{action_timing_diagram}；"
                f"动作行为分析：{action_behavior_analysis}；"
                f"故障原因分析：{fault_analysis}；"
                "使用专业语言总结本次故障的原因、后果及保护系统表现，保持段落格式输出。"
                "注意：不要生成[小结]这种标题，不要特殊符号。"
            )
            response = client.chat.completions.create(
                model="glm-4-plus",
                messages=[
                    {"role": "system", "content": "你是一名专业的电力系统分析专家。"},
                    {"role": "user", "content": prompt}
                ],
            )
            summary_text = re.sub(r'^\s*\d*\.?\s*小结', '', response.choices[0].message.content.strip()).lstrip('： \n')
        else:
            summary_text = "缺少分析内容文件，无法生成本节内容。"
        add_section("4. 小结", summary_text, level=3)

        # === 10. 四、事件原因和性质 ===
        update_progress(9)
        add_section("四、事件原因和性质", "", level=1)

        event_summary = summary
        basic_info = f"{station}{equipment}，电压等级为{voltage_level}，故障线路为{fault_line}。"
        analysis_info = f"跳闸发生时，{equipment}的故障相为{fault_phase}，重合闸动作为{reclose_action}。"

        # 直接原因
        prompt = (
            f"请生成事件的直接原因，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家。"}, {"role": "user", "content": prompt}]
        )
        direct_cause = response.choices[0].message.content

        # 间接原因
        prompt = (
            f"请生成事件的间接原因，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家。"}, {"role": "user", "content": prompt}]
        )
        indirect_cause = response.choices[0].message.content

        # 管理原因
        prompt = (
            f"请生成事件的管理原因，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家。"}, {"role": "user", "content": prompt}]
        )
        management_cause = response.choices[0].message.content

        add_section("（一）事件原因", "", level=2)
        add_section("1.直接原因", direct_cause, level=3)
        add_section("2.间接原因", indirect_cause, level=3)
        add_section("3.管理原因", management_cause, level=3)

        # 事件性质
        prompt = (
            f"请生成事件的性质，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家。"}, {"role": "user", "content": prompt}]
        )
        event_nature = response.choices[0].message.content
        add_section("（二）事件性质", event_nature, level=2)

        # 事件暴露问题
        prompt = (
            f"请生成事件暴露的问题，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家。"}, {"role": "user", "content": prompt}]
        )
        event_exposure = response.choices[0].message.content
        add_section("（三）事件暴露问题", event_exposure, level=2)

        # === 11. 五、整改措施 ===
        update_progress(10)
        cause_info = f"直接原因：{direct_cause}\n间接原因：{indirect_cause}\n管理原因：{management_cause}\n事件性质：{event_nature}"
        prompt = (
            f"请生成事件防范和整改措施建议，不要分点撰写，以段落格式输出："
            f"事件简述：{event_summary}\n基本情况：{basic_info}\n事件经过分析：{analysis_info}\n事件原因和性质：{cause_info}"
        )
        response = client.chat.completions.create(
            model="glm-4-plus",
            messages=[{"role": "system", "content": "你是一个电力系统专家，擅长提供防范和整改措施。"}, {"role": "user", "content": prompt}]
        )
        prevention = response.choices[0].message.content
        add_section("五、事件防范和整改措施建议", prevention, level=1)

        # === 12. 保存文档 ===
        update_progress(11)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_id = f"REP-{timestamp}"
        output_doc = os.path.join(reports_dir, f'{report_id}.docx')
        document.save(output_doc)

        task_store[task_id]['status'] = 'completed'
        task_store[task_id]['progress'] = 100
        task_store[task_id]['message'] = "报告生成完成！"
        task_store[task_id]['report_file'] = f'{report_id}.docx'
        task_store[task_id]['report_id'] = report_id

    except Exception as e:
        import traceback
        task_store[task_id]['status'] = 'failed'
        task_store[task_id]['error'] = str(e)
        task_store[task_id]['traceback'] = traceback.format_exc()
        print(f"[ERROR] 报告生成失败: {traceback.format_exc()}")


# ============================================================
# API 路由
# ============================================================

@app.route('/')
def index():
    return jsonify({
        'service': '故障报告自动生成系统',
        'version': '1.0.0',
        'status': 'running'
    })


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传文件到指定分类目录"""
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400

    file = request.files['file']
    category = request.form.get('category', '事件简述')

    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    if file and allowed_file(file.filename):
        # 保留原始中文文件名（secure_filename 会破坏中文导致同名覆盖），仅用 basename 防路径穿越
        filename = os.path.basename(file.filename)
        category_dir = os.path.join(UPLOADS_DIR, category)
        os.makedirs(category_dir, exist_ok=True)

        # 如果目录中已有同名文件，先删除
        filepath = os.path.join(category_dir, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
        file.save(filepath)

        return jsonify({
            'message': f'文件 {filename} 上传成功',
            'category': category,
            'filename': filename
        })

    return jsonify({'error': '不支持的文件类型'}), 400


@app.route('/api/files', methods=['GET'])
def list_files():
    """列出所有已上传的文件"""
    uploads_root = Path(UPLOADS_DIR)
    result = {}

    if uploads_root.exists():
        for category_dir in uploads_root.iterdir():
            if category_dir.is_dir():
                files = []
                for f in category_dir.iterdir():
                    if f.is_file():
                        files.append({
                            'name': f.name,
                            'size': f.stat().st_size,
                            'modified': datetime.fromtimestamp(f.stat().st_mtime).isoformat()
                        })
                if files:
                    result[category_dir.name] = files

    return jsonify(result)


@app.route('/api/generate', methods=['POST'])
def generate_report():
    """启动报告生成任务"""
    task_id = datetime.now().strftime("%Y%m%d%H%M%S%f")

    task_store[task_id] = {
        'id': task_id,
        'status': 'pending',
        'progress': 0,
        'message': '任务已创建，等待开始...',
        'created_at': datetime.now().isoformat()
    }

    # 在后台线程中生成报告
    thread = threading.Thread(
        target=generate_report_sync,
        args=(task_id, UPLOADS_DIR, REPORTS_DIR)
    )
    thread.daemon = True
    thread.start()

    return jsonify({
        'task_id': task_id,
        'status': 'pending',
        'message': '报告生成任务已启动'
    })


@app.route('/api/generate-single', methods=['POST'])
def generate_single():
    """上传单个故障文档，AI 自动解析字段并生成报告"""
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件类型，请使用 txt/docx/doc/pdf'}), 400

    filename = os.path.basename(file.filename)
    single_dir = os.path.join(BASE_DIR, 'uploads', 'single')
    os.makedirs(single_dir, exist_ok=True)
    filepath = os.path.join(single_dir, filename)
    file.save(filepath)

    try:
        text = read_document_text(filepath)
        if not text.strip():
            raise ValueError("文档内容为空")
        fields = extract_event_fields(text)
    except Exception as e:
        return jsonify({'error': f'文档解析失败：{str(e)}'}), 500

    # 启动后台生成任务：用抽取到的字段直接套用固定模板（无额外 AI 调用）
    task_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    task_store[task_id] = {
        'id': task_id,
        'status': 'pending',
        'progress': 0,
        'message': '任务已创建，等待开始...',
        'created_at': datetime.now().isoformat()
    }
    thread = threading.Thread(
        target=generate_report_from_fields_sync,
        args=(task_id, fields, REPORTS_DIR)
    )
    thread.daemon = True
    thread.start()

    return jsonify({
        'task_id': task_id,
        'status': 'pending',
        'message': '智能生成任务已启动'
    })


@app.route('/api/files/delete', methods=['POST'])
def delete_uploaded_file():
    """删除上传目录中指定分类下的文件"""
    data = request.get_json() or {}
    category = data.get('category')
    filename = data.get('filename')
    if not category or not filename:
        return jsonify({'error': '缺少 category 或 filename'}), 400
    filename = os.path.basename(filename)
    filepath = os.path.join(UPLOADS_DIR, category, filename)
    if os.path.exists(filepath) and os.path.isfile(filepath):
        os.remove(filepath)
        return jsonify({'message': f'文件 {filename} 已删除'})
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/files/download', methods=['GET'])
def download_uploaded_file():
    """下载上传目录中的文件"""
    category = request.args.get('category')
    filename = request.args.get('filename')
    if not category or not filename:
        return jsonify({'error': '缺少 category 或 filename'}), 400
    filename = os.path.basename(filename)
    filepath = os.path.join(UPLOADS_DIR, category, filename)
    if os.path.exists(filepath) and os.path.isfile(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/reports/delete', methods=['POST'])
def delete_report():
    """删除已生成的报告"""
    data = request.get_json() or {}
    filename = data.get('filename')
    if not filename:
        return jsonify({'error': '缺少 filename'}), 400
    filename = os.path.basename(filename)
    filepath = os.path.join(REPORTS_DIR, filename)
    if os.path.exists(filepath) and os.path.isfile(filepath):
        os.remove(filepath)
        return jsonify({'message': f'报告 {filename} 已删除'})
    return jsonify({'error': '报告不存在'}), 404


@app.route('/api/task/<task_id>', methods=['GET'])
def get_task_status(task_id):
    """查询任务状态"""
    task = task_store.get(task_id)
    if not task:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(task)


@app.route('/api/reports', methods=['GET'])
def list_reports():
    """列出已生成的报告"""
    reports = []
    reports_dir = Path(REPORTS_DIR)

    if reports_dir.exists():
        for f in sorted(reports_dir.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True):
            if f.is_file() and f.suffix in ['.docx', '.pdf']:
                reports.append({
                    'name': f.name,
                    'size': f.stat().st_size,
                    'type': f.suffix[1:],
                    'modified': datetime.fromtimestamp(f.stat().st_mtime).isoformat()
                })

    return jsonify(reports)


@app.route('/api/download/<filename>', methods=['GET'])
def download_report(filename):
    """下载生成的报告"""
    filepath = os.path.join(REPORTS_DIR, secure_filename(filename))
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return jsonify({'error': '文件不存在'}), 404


# ============================================================
# 启动
# ============================================================

if __name__ == '__main__':
    print("=" * 60)
    print("  故障报告自动生成系统 - 后端服务")
    print("=" * 60)
    print(f"  上传目录: {UPLOADS_DIR}")
    print(f"  报告目录: {REPORTS_DIR}")
    print(f"  服务地址: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
