#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测评反馈报告生成器 - Web版
=========================
启动: python3 app.py
访问: http://localhost:5000
"""

import base64
import gc
import hashlib
import io
import json
import os
import re
import shutil
import threading
import uuid
import zipfile
from datetime import datetime, timedelta

import openpyxl
from flask import Flask, render_template, request, jsonify, send_file
from PIL import Image
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max upload
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 预览页版本号：升级分析逻辑后修改，旧预览页提交时会被拒绝，避免旧内容覆盖新结果
APP_VERSION = "2026-10-28-1"

# 报告生成任务（run_id -> 进度信息），用于前端轮询
JOBS = {}
JOBS_LOCK = threading.Lock()
# 同一时间只跑一个生成任务，避免多人同时生成把 512MB 内存撑爆
GENERATION_LOCK = threading.Lock()

# 默认板块建议（Excel 中出现 config 里没有的板块时用）
DEFAULT_ADVICE = {
    "topic": "该板块知识",
    "advice": "针对薄弱知识点做专项练习，每天坚持练几道题找找感觉。",
}

# 上传文件临时目录
UPLOAD_DIR = os.path.join(BASE_DIR, "_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 报告输出目录
OUTPUT_DIR = os.path.join(BASE_DIR, "反馈报告")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def new_run_dir():
    """为每次生成创建独立目录，避免不同老师之间的报告混在一起"""
    # 顺手清理 24 小时前的旧批次，避免云服务器磁盘被占满
    now = datetime.now()
    for entry in os.listdir(OUTPUT_DIR):
        path = os.path.join(OUTPUT_DIR, entry)
        if os.path.isdir(path) and (now - datetime.fromtimestamp(os.path.getmtime(path))) > timedelta(hours=24):
            shutil.rmtree(path, ignore_errors=True)
    run_id = now.strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    run_dir = os.path.join(OUTPUT_DIR, run_id)
    os.makedirs(run_dir, exist_ok=True)
    return run_id, run_dir


# ============================================================
#  图片处理
# ============================================================
def remove_white_bg(path):
    """去除图片白色背景，返回透明 PNG 的 base64"""
    img = Image.open(path).convert("RGBA")
    data = img.getdata()
    new_data = []
    for r, g, b, a in data:
        if r > 235 and g > 235 and b > 235:
            new_data.append((r, g, b, 0))
        else:
            new_data.append((r, g, b, a))
    img.putdata(new_data)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


def img_to_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


# ============================================================
#  Excel 解析（自动从表头提取满分）
# ============================================================
def parse_headers(headers):
    """宽松解析表头：任何非姓名/总分列都视为板块列，满分从表头提取或后续从数据推断。
    支持的满分格式（按优先级）：
      - 圆括号: 单词(30) / 单词（30）/ 单词（满分30）/ 单词（30分）
      - 方括号: 单词[30] / 单词【30】
      - 斜杠: 单词/30
      - 后缀分: 单词 30分
      - 满分前缀: 单词 满分30 / 单词 满分：30
      - 纯数字: 单词 30
    若表头无数字，max_score 返回 None，由 load_excel 从数据推断。
    """
    sections = []
    score_indices = []
    name_idx = None
    total_idx = None
    for i, h in enumerate(headers):
        h = str(h) if h else ""
        h_strip = h.strip()
        if not h_strip:
            continue
        # 姓名列（宽松匹配：表头含「姓名/名字/学生/name」即识别）
        if any(k in h_strip for k in ("姓名", "名字", "学生")) or \
           h_strip.lower() in ("name", "student"):
            name_idx = i
            continue
        # 总分列（宽松匹配：表头含「总分/合计/total」即识别）
        h_clean = re.sub(r"\s*[（(].*?[)）]", "", h).strip()
        if any(k in h_strip for k in ("总分", "合计")) or \
           h_clean.lower() in ("total",) or \
           re.sub(r"\s*\d+(?:\.\d+)?\s*分?\s*$", "", h_clean).strip() in ("总分", "合计"):
            total_idx = i
            continue
        # 各种满分格式
        max_score = None
        name = h
        # 1) 圆括号 (30) / （30）/ （满分30）/ （30分）/ （满分30分）
        m = re.search(r"[（(](?:满分)?[^)）\d]*(\d+(?:\.\d+)?)[^)）]*[)）]", h)
        if m:
            max_score = float(m.group(1))
            name = re.sub(r"\s*[（(][^)）]*[)）]", "", h).strip()
        # 2) 方括号 [30] / 【30】 / 【30分】
        if max_score is None:
            m = re.search(r"[\[【](?:满分)?[^\]】\d]*(\d+(?:\.\d+)?)[^\]】]*[\]】]", h)
            if m:
                max_score = float(m.group(1))
                name = re.sub(r"\s*[\[【][^\]】]*[\]】]", "", h).strip()
        # 3) 斜杠 30（如 单词/30）
        if max_score is None:
            m = re.search(r"[/／]\s*(\d+(?:\.\d+)?)\s*$", h)
            if m:
                max_score = float(m.group(1))
                name = re.sub(r"\s*[/／]\s*\d+(?:\.\d+)?\s*$", "", h).strip()
        # 4) 后缀「30分」（如 单词 30分 / 单词30分）
        if max_score is None:
            m = re.search(r"(\d+(?:\.\d+)?)\s*分\s*$", h)
            if m:
                max_score = float(m.group(1))
                name = re.sub(r"\s*\d+(?:\.\d+)?\s*分\s*$", "", h).strip()
        # 5) 「满分30」或「满分：30」或「满分 30」
        if max_score is None:
            m = re.search(r"满分[:：]?\s*(\d+(?:\.\d+)?)", h)
            if m:
                max_score = float(m.group(1))
                name = re.sub(r"\s*满分[:：]?\s*\d+(?:\.\d+)?\s*", "", h).strip()
        # 6) 单独数字列（如「单词 30」或「30 单词」）
        if max_score is None:
            m = re.search(r"(\d+(?:\.\d+)?)", h)
            if m:
                max_score = float(m.group(1))
                name = re.sub(r"\s*\d+(?:\.\d+)?\s*", "", h).strip()
        # 7) 表头无数字（如纯「单词」「短语」「阅读」），max_score=None，后续从数据推断
        if max_score is None:
            name = h_strip
        # 板块名清洗：去掉常见后缀（「得分」「成绩」「分数」）
        name = re.sub(r"\s*[得分成绩分数]+\s*$", "", name).strip()
        if name:
            sections.append((name, max_score))
            score_indices.append(i)
    return sections, score_indices, name_idx, total_idx


def infer_max_from_column(values):
    """从一列得分数据推断满分：取最大值；若最大值为0或无效，返回 100 兜底"""
    valid = [v for v in values if v is not None and v > 0]
    if not valid:
        return 100.0
    max_val = max(valid)
    # 满分通常是整数且是 5/10 的倍数，但这里直接用最大值更通用
    return float(max_val)


def ocr_image_to_rows(path):
    """用 macOS Vision API 识别图片中的表格，返回 rows = [[cell, cell, ...], ...]
    Vision API 按行识别文字，然后按 Y 坐标聚类成行，按 X 坐标排序成列。
    """
    try:
        from Vision import VNRecognizeTextRequest, VNImageRequestHandler
        from Quartz import CIImage, NSURL
    except ImportError:
        print("Vision API 不可用，无法识别图片")
        return None

    url = NSURL.fileURLWithPath_(path)
    ci_image = CIImage.imageWithContentsOfURL_(url)
    if ci_image is None:
        print("无法加载图片")
        return None

    request = VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(1)  # 1 = accurate
    request.setRecognitionLanguages_(["zh-Hans", "zh-Hant", "en"])
    request.setUsesLanguageCorrection_(True)

    handler = VNImageRequestHandler.alloc().initWithCIImage_options_(ci_image, None)
    success, error = handler.performRequests_error_([request], None)
    if not success:
        print(f"OCR 识别失败: {error}")
        return None

    results = request.results()
    if not results:
        return None

    # 收集每个识别结果的文字和位置
    # boundingBox: (x, y, w, h) in normalized coordinates (0-1), origin at bottom-left
    items = []
    for obs in results:
        try:
            candidate = obs.topCandidates_(1)[0]
            text = str(candidate.string())
        except Exception:
            continue
        bbox = obs.boundingBox()
        origin_x = float(bbox.origin.x)
        origin_y = float(bbox.origin.y)
        width = float(bbox.size.width)
        height = float(bbox.size.height)
        # 转换为 top-left origin
        y_top = 1.0 - (origin_y + height)
        x_center = origin_x + width / 2
        items.append({"text": text, "x": origin_x, "xc": x_center, "y": y_top, "h": height, "w": width})

    if not items:
        return None

    # 按 Y 坐标聚类成行
    items.sort(key=lambda t: (t["y"], t["x"]))
    rows_raw = []
    current_row = []
    current_y = None
    y_tolerance = 0.025  # 2.5% 的 Y 容差

    for item in items:
        if current_y is None or abs(item["y"] - current_y) <= y_tolerance:
            current_row.append(item)
            if current_y is None:
                current_y = item["y"]
        else:
            if current_row:
                rows_raw.append(current_row)
            current_row = [item]
            current_y = item["y"]
    if current_row:
        rows_raw.append(current_row)

    # 每行内按 X 坐标排序，输出单元格
    rows = []
    for row_items in rows_raw:
        row_items.sort(key=lambda t: t["x"])
        # 合并相邻且距离很近的文字为一个单元格
        merged = []
        for item in row_items:
            if merged and abs(item["x"] - (merged[-1]["x"] + merged[-1]["w"])) < 0.015:
                # 合并
                merged[-1]["text"] += item["text"]
                merged[-1]["w"] = (item["x"] + item["w"]) - merged[-1]["x"]
            else:
                merged.append(dict(item))
        cells = [c["text"].strip() for c in merged if c["text"].strip()]
        if cells:
            rows.append(cells)

    return rows


def load_excel(path):
    """返回 (students, sections)
    students = [(姓名, {板块: 得分}, 总分), ...]
    sections  = [(板块名, 满分), ...]
    宽松策略：表头没有满分时，从该列数据取最大值作为满分。
    支持 .xlsx / .xls / .csv（含 GBK/UTF-8 自动检测）/ 图片（.png/.jpg/.jpeg 用 OCR）
    """
    import csv as csv_mod
    lower = path.lower()
    # 图片格式：用 macOS Vision API OCR 识别
    if lower.endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
        rows = ocr_image_to_rows(path)
        if not rows:
            raise RuntimeError("图片识别失败，请确保图片清晰，或改用 Excel/CSV 格式")
    elif lower.endswith(".csv"):
        # 自动检测编码
        text_rows = None
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030", "big5"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv_mod.reader(f)
                    text_rows = [row for row in reader]
                break
            except (UnicodeDecodeError, Exception):
                continue
        if not text_rows:
            return [], []
        rows = text_rows
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []

    headers = [str(h) if h else "" for h in rows[0]]
    sections, score_indices, name_idx, total_idx = parse_headers(headers)

    if name_idx is None:
        # 没有明确姓名列时，默认第 0 列
        name_idx = 0

    students = []
    for row in rows[1:]:
        # 跳过完全空行
        if not any(c not in (None, "", " ") for c in row):
            continue
        name = row[name_idx] if name_idx < len(row) else None
        if not name or str(name).strip() == "":
            continue
        scores = {}
        for sec_idx, col_idx in enumerate(score_indices):
            if col_idx < len(row):
                sec_name = sections[sec_idx][0]
                val = row[col_idx]
                try:
                    scores[sec_name] = float(val) if val not in (None, "", " ") else 0
                except (ValueError, TypeError):
                    scores[sec_name] = 0

        if total_idx is not None and total_idx < len(row) and row[total_idx] is not None:
            try:
                total = float(row[total_idx])
            except (ValueError, TypeError):
                total = sum(scores.values())
        else:
            total = sum(scores.values())

        students.append((str(name).strip(), scores, total))

    # 对表头没有满分的板块，从数据推断
    for sec_idx, (name, max_score) in enumerate(sections):
        if max_score is None:
            col_values = [s[name] for _, s, _ in students if name in s]
            inferred = infer_max_from_column(col_values)
            sections[sec_idx] = (name, inferred)

    return students, sections


def load_default_advice():
    """从 config.json 加载默认板块建议"""
    config_path = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            return cfg.get("section_advice", {})
    return {}


# ============================================================
#  PDF 试卷分析
# ============================================================
def extract_pdf_text(path):
    """提取 PDF 全文，自动识别并正确拼接双栏排版（期中/中考试卷常见）。
    当页面宽度 > 1000pt 且明显宽于高度时，按左栏→右栏顺序拼接，
    避免 pdfplumber 默认抽取把左右栏文字交错粘连。
    """
    text = ""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                w, h = page.width, page.height
                if w > 1000 and w > h * 1.2:
                    left = page.crop((0, 0, w * 0.52, h)).extract_text() or ''
                    right = page.crop((w * 0.48, 0, w, h)).extract_text() or ''
                    page_text = left + "\n" + right
                else:
                    page_text = page.extract_text() or ''
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"PDF 提取失败: {e}")
    return text


# ============================================================
#  人教版初中英语语法体系 - 考点检测引擎
#  覆盖：8大时态 / 被动语态 / 非谓语 / 从句 / 特殊句式 / 词法
# ============================================================
POS_ZH = {'n': '名词', 'v': '动词', 'adj': '形容词', 'adv': '副词',
         'prep': '介词', 'conj': '连词', 'num': '数词',
         'art': '冠词', 'pron': '代词', 'int': '感叹词'}

# 时态检测规则（人教版8大时态 + 主将从现）
TENSE_RULES = [
    ("现在完成时", [
        r'\b(?:have|has)\s+(?:\w+ed|been|done|gone|written|taken|given|known|seen|shown|driven|grown|flown|fallen|broken|chosen|spoken|stolen|woken|ridden|risen|beaten|eaten|forgotten|hidden|forbidden|built|bought|brought|caught|taught|thought|fought|sought)\b',
        r'\b(?:already|yet|ever|never|just|recently|so far|in the past)\b',
        r'\bfor\s+\w+\b.*\b(?:have|has)\b',
        r'\bsince\s+\w+\b.*\b(?:have|has)\b',
    ]),
    ("过去进行时", [r'\b(?:was|were)\s+\w+ing\b']),
    ("过去完成时", [
        r'\bhad\s+(?:\w+ed|been|done|gone|written|taken|given|known|seen|shown|flown|fallen|broken|chosen|spoken|stolen|woken|ridden|risen|beaten|eaten|forgotten|hidden|forbidden|built|bought|brought|caught|taught|thought|fought|sought)\b',
        r'\bby the time\b',
        r'\bby the end of\b',
    ]),
    ("一般将来时", [
        r'\bwill\s+\w+\b',
        r'\bshall\s+\w+\b',
        r'\bbe going to\s+\w+\b',
        r'\bnext (?:week|month|year|Monday)\b',
        r'\b(?:tomorrow|soon|in a few minutes)\b',
    ]),
    ("现在进行时", [r'\b(?:am|is|are)\s+\w+ing\b']),
    ("一般过去时", [
        r'\b(?:went|came|took|gave|saw|made|said|told|found|got|bought|brought|caught|taught|thought|knew|grew|flew|drove|rode|spoke|broke|chose|wrote|fell|beat|ate|forgot|hid|forbade|rose|woke|spent|sent|lent|built|meant|dealt|paid|laid|stood|understood|held|sold|told|won|shot|lost|cost|cut|hit|hurt|let|put|set|read|beat)\b',
        r'\b\w+ed\b',
        r'\blast (?:year|week|month|night)\b',
        r'\b(?:yesterday|in \d{4}|just now)\b',
    ]),
    ("一般现在时", [
        r'\b(?:do|does|don\'t|doesn\'t)\s+\w+\b',
        r'\b(?:always|usually|often|sometimes|never|every (?:day|week|year))\b',
    ]),
    ("主将从现（if条件句）", [
        r'\bif\s+\w+\s+\w+\b.*\bwill\b',
        r'\bwon\'t\s+\w+.*\bif\b',
    ]),
]

# 不定式常见动词（人教版7-9年级）
TO_DO_VERBS = ["decide", "want", "hope", "refuse", "agree", "ask", "tell", "advise",
               "encourage", "invite", "expect", "promise", "plan", "teach", "choose",
               "afford", "offer", "manage", "pretend", "fail", "happen", "seem",
               "appear", "wish", "would like", "learn", "remember", "forget", "stop",
               "try", "need", "used", "allow", "permit", "require", "cause"]

# 动名词常见动词
DOING_VERBS = ["enjoy", "finish", "practice", "practise", "mind", "avoid", "consider",
               "suggest", "can't help", "can not help", "cannot help", "look forward to",
               "be busy", "feel like", "give up", "keep", "imagine", "dislike", "miss",
               "escape", "risk", "postpone", "delay", "report", "admit", "deny",
               "appreciate", "tolerate", "resist", "spend"]


SECTION_KEYWORDS = [
    # 中文
    "单项选择", "单项填空", "选择填空", "完形填空", "语法填空",
    "适当形式填空", "任务型阅读", "六选五", "七选五", "阅读理解",
    "阅读", "完成句子", "书面表达", "作文", "单词", "短语", "翻译",
    "听力", "选词填空", "还原句子",
    # 英文（不分大小写，canonical_section_name 会统一处理）
    "multiple choice", "cloze test", "grammar fill", "grammar sort",
    "reading", "writing", "composition", "translation", "listening",
]


def split_pdf_by_sections(text):
    """按序号或板块标题切分试卷文本，返回 [(板块标题, 板块内容), ...]。
    兼容：一、单项选择 / Ⅰ. 单项选择 / 第一部分 听力 / 无序号直接写板块名，
    也兼容标题嵌在上一句结尾的情况（如 "...sports, from...people. V. 语法填空..."）。
    """
    # 1) 先按"完整标题+罗马数字/中文数字"抓，允许前面是空格或句子结尾
    numbered = re.compile(
        r'(?:^|[.．!！?？；;\s])\s*'
        r'((?:I{1,3}|IV|V|VI{0,3}|IX|X|Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)'
        r'\s*[、.．]\s*[^\n]{1,40}|'
        r'[一二三四五六七八九十]+\s*[、.．]\s*[^\n]{1,40}|'
        r'第\s*[一二三四五六七八九十]+\s*[部分节]\s*[^\n]{0,40})'
    )
    found = []
    for m in numbered.finditer(text):
        title = m.group(1).strip()
        # 只接受真正的板块标题，避免把选项 D. / 句子里的数字当成板块
        if canonical_section_name(title) in KNOWN_SECTIONS:
            found.append((m.end(), title))  # 记录"标题结束"的位置，正文从这里开始

    # 2) 没抓到或抓得不够时，兜底用板块关键词所在行（支持中英文、不区分大小写）
    if len(found) < 2:
        kw = re.compile(r'(?:^|\n|\s)\s*((?:' + '|'.join(re.escape(k) for k in SECTION_KEYWORDS) + r')[^\n]{0,25})',
                        re.IGNORECASE)
        seen = set()
        for m in kw.finditer(text):
            title = m.group(1).strip()
            key = canonical_section_name(title)
            if key not in KNOWN_SECTIONS or key in seen:
                continue
            if title.startswith(("阅读下面", "阅读短文", "阅读下列", "根据")):
                continue
            seen.add(key)
            found.append((m.end(), title))

    if not found:
        return [("全文", text)]

    found.sort(key=lambda x: x[0])
    # 同一位置去重
    picked = []
    for pos, title in found:
        if picked and abs(pos - picked[-1][0]) <= 2:
            continue
        picked.append((pos, title))

    if len(picked) < 2:
        return [("全文", text)]

    sections = []
    for i, (end_pos, title) in enumerate(picked):
        start = end_pos
        end = picked[i + 1][0] if i + 1 < len(picked) else len(text)
        sections.append((title, text[start:end].strip()))
    return sections


KNOWN_SECTIONS = {'单选', '语法填空', '完形', '填空', '翻译', '六选五',
                  '任务型阅读', '阅读', '作文', '短语', '单词', '听力'}


def canonical_section_name(name):
    """把 Excel 表头 / 试卷标题统一成标准板块名。
    支持中文 + 英文标签，例如 Multiple Choice/单项选择 -> 单选。
    """
    raw = str(name or '')
    t = re.sub(r'[\(（\[【].*?[\)）\]】]', '', raw)
    t = re.sub(r'\s+', '', t).strip()
    if not t:
        return ''

    # ===== 英文别名（大小写不敏感） =====
    # 先去掉开头的罗马/中文数字序号，避免 "III. Reading" 匹配不上
    body = re.sub(r'^[IVXLCDMivxlcdmⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ一二三四五六七八九十]+\s*[、.．]\s*', '', t)
    t_clean = re.sub(r'[\s\-]+', '', body.lower())
    # 单选：Multiple Choice / Multiple-choice / Part I Multiple Choice
    if any(k in t_clean for k in ["multiplechoice", "partimultiplechoice"]):
        return "单选"
    if "grammarsort" in t_clean or "grammarfill" in t_clean or "gramfill" in t_clean:
        return "语法填空"
    if "clozetest" in t_clean or t_clean == "cloze":
        return "完形"
    if t_clean in {"reading", "read"} or "partivreading" in t_clean or "part4reading" in t_clean:
        return "阅读"
    if t_clean in {"translate", "translation"} or "partvi" in t_clean or "part6translate" in t_clean:
        return "翻译"
    if t_clean in {"writing", "composition", "essay"} or "partvii" in t_clean or "part7writing" in t_clean:
        return "作文"
    if t_clean == "listening" or "partilistening" in t_clean:
        return "听力"

    # ===== 中文别名（顺序很重要：先特殊板块，后大类） =====
    if any(k in t for k in ["单项选择", "选择填空", "选择题"]):
        return "单选"
    if "语法填空" in t:
        return "语法填空"
    if "完形" in t:
        return "完形"
    if "任务型" in t:
        return "任务型阅读"
    if any(k in t for k in ["六选五", "七选五", "五选五", "5选5", "6选5", "7选5", "还原句子", "选词填空"]):
        return "六选五"
    if any(k in t for k in ["适当形式", "词形变化", "词形转换", "变形填空"]):
        return "填空"
    if any(k in t for k in ["完成句子", "翻译", "汉译英", "英译汉", "互译"]):
        return "翻译"
    if "阅读" in t:
        return "阅读"
    if any(k in t for k in ["作文", "书面表达", "写作"]):
        return "作文"
    if any(k in t for k in ["短语", "词组"]):
        return "短语"
    if any(k in t for k in ["单词", "词汇"]):
        return "单词"
    if "听力" in t:
        return "听力"
    if "填空" in t:
        return "填空"
    return t


def map_to_section_key(title, content=""):
    """把试卷板块标题映射到 Excel 中的板块名（返回标准名，未识别返回 None）"""
    key = canonical_section_name(title)
    return key if key in KNOWN_SECTIONS else None


def detect_tenses(text):
    """检测文本中出现的时态，返回 [时点名]"""
    text_l = text.lower()
    found = []
    for tense, patterns in TENSE_RULES:
        for p in patterns:
            if re.search(p, text_l, re.IGNORECASE):
                if tense not in found:
                    found.append(tense)
                break
    return found


def detect_passive_voice(text):
    """检测被动语态及其时态，返回 [时态被动]"""
    text_l = text.lower()
    detail = []
    if re.search(r'\b(?:am|is|are)\s+\w+(?:ed|en)\b', text_l) or \
       re.search(r'\b(?:am|is|are)\s+being\s+\w+(?:ed|en)\b', text_l):
        detail.append("一般现在时被动")
    if re.search(r'\b(?:was|were)\s+\w+(?:ed|en)\b', text_l) or \
       re.search(r'\b(?:was|were)\s+being\s+\w+(?:ed|en)\b', text_l):
        detail.append("一般过去时被动")
    if re.search(r'\bwill be\s+\w+(?:ed|en)\b', text_l) or \
       re.search(r'\b(?:am|is|are) going to be\s+\w+(?:ed|en)\b', text_l):
        detail.append("一般将来时被动")
    if re.search(r'\b(?:have|has) been\s+\w+(?:ed|en)\b', text_l):
        detail.append("现在完成时被动")
    if re.search(r'\bhad been\s+\w+(?:ed|en)\b', text_l):
        detail.append("过去完成时被动")
    if re.search(r'\b(?:must|can|should|may|need|could|might)\s+be\s+\w+(?:ed|en)\b', text_l):
        detail.append("含情态动词被动")
    return detail


def detect_nonfinite(text):
    """检测非谓语动词，返回 [非谓语类型]"""
    text_l = text.lower()
    detail = []
    # 不定式 to do
    to_do_patterns = [
        r'\bto\s+(?:do|be|improve|close|waste|complete|learn|study|finish|get|go|come|know|see|realize|reach|help|stop|keep|remember|forget|decide|want|hope|refuse|agree|ask|tell|advise|encourage|invite|expect|promise|plan|teach|choose|afford|offer|manage|pretend|fail|happen|seem|appear|wish|not\s+\w+)\b',
    ]
    for p in to_do_patterns:
        if re.search(p, text_l):
            detail.append("不定式 to do")
            break
    # 动名词 doing
    for verb in DOING_VERBS:
        if re.search(r'\b' + re.escape(verb) + r'\b\s+\w+ing\b', text_l):
            detail.append("动名词 doing")
            break
    # would rather do
    if re.search(r'\bwould rather\s+\w+', text_l):
        detail.append("would rather do")
    # 过去分词 done（作定语/表语/宾补）
    if re.search(r'\b(?:is|are|was|were|be|been|seem|seemed|look|looked|become|became|remain|remained|feel|felt)\s+\w+(?:ed|en)\b', text_l):
        detail.append("过去分词 done")
    return detail


def detect_all_patterns(text):
    """综合检测：从句 + 特殊句式 + 固定搭配 + 试卷自带结构标签。
    返回 [句型名...]。先捞试卷写死的标签（最准），再用规则扫英文结构。
    """
    text_l = text.lower()
    found = []

    # ====== A. 试卷自带结构标签（PDF 里括号/题后常写，最准） ======
    paper_tags = [
        # 从句类
        ("同位语", "同位语从句"), ("同位语从句", "同位语从句"),
        ("宾语从句", "宾语从句"), ("宾从", "宾语从句"),
        ("定语从句", "定语从句"), ("定从", "定语从句"),
        ("时间状从", "时间状语从句"), ("条件状从", "条件状语从句"),
        ("原因状从", "原因状语从句"),
        ("让步状从", "让步状语从句"), ("结果状从", "结果状语从句"),
        ("目的状从", "目的状语从句"),
        ("主语从句", "主语从句"), ("表语从句", "表语从句"),
        # 特殊句式
        ("强调句", "强调句"), ("倒装句", "倒装句"),
        ("感叹句", "感叹句"), ("反意疑问句", "反意疑问句"),
        # 固定搭配 / 结构
        ("伴随", "with 伴随结构"), ("伴随状语", "with 伴随结构"),
        ("with", "with 伴随结构"),  # 注意不要误抓介词
        ("oneof", "one of + 复数名词"), ("one of", "one of + 复数名词"),
        ("notonly", "not only...but also..."), ("not only", "not only...but also..."),
        ("neither", "neither...nor..."), ("either", "either...or..."),
        ("tooto", "too...to..."), ("too...to", "too...to..."),
        ("notuntil", "not...until..."), ("not...until", "not...until..."),
        ("wouldrather", "would rather...than..."), ("would rather", "would rather...than..."),
        ("sothat", "so...that..."), ("so...that", "so...that..."),
        ("suchthat", "such...that..."), ("such...that", "such...that..."),
        ("soasto", "so as to / in order to"), ("inorderto", "so as to / in order to"),
        ("it takes", "It takes sb. ... to do"),
        ("hadbetter", "had better do"), ("had better", "had better do"),
        ("usedto", "used to do"), ("used to", "used to do"),
        ("beusedto", "be used to doing"), ("be used to", "be used to doing"),
    ]
    # 中文固定结构（必须用正则，避免“太棒了”“不仅”等误判）
    cn_patterns = [
        (r'不仅.*?而且', "not only...but also..."),
        (r'既不.*?也', "neither...nor..."),
        (r'要么.*?要么', "either...or..."),
        (r'太.*?以致', "too...to... 句型"),
        (r'直到.*?才', "not...until... 句型"),
        (r'如此.*?以致', "so...that... 句型"),
        (r'花费.*?时间', "It takes sb. ... to do"),
        (r'最好', "had better do"),
        (r'过去常', "used to do"),
        (r'习惯于', "be used to doing"),
    ]
    for pat, name in cn_patterns:
        if re.search(pat, text) and name not in found:
            found.append(name)
    for tag, name in paper_tags:
        if tag in text or tag in text_l:
            if name not in found:
                found.append(name)

    # ====== B. 英文从句结构检测（仅当文本里有英文句子时才扫，避免中文题干误判） ======
    has_english = bool(re.search(r'[a-zA-Z]{3,}', text))
    if has_english:
        # 宾语从句
        if re.search(r'\b(?:i|we|he|she|they|you)\s+(?:think|believe|hope|wonder|know|said|asked|told|heard|found|saw|realized|didn\'t know|don\'t know|want to know|understand)\b', text_l) or \
           re.search(r'\bdo you (?:think|know|believe|realize)\b', text_l) or \
           re.search(r'\b(?:that|if|whether|what|who|whom|whose|where|when|why|how)\b.*\b(?:think|know|said|told|asked|wonder)\b', text_l):
            if "宾语从句" not in found:
                found.append("宾语从句")
        # 定语从句
        if re.search(r'\b(?:which|who|whom|whose|that)\s+(?:is|are|was|were|has|have|will|can|do|does|did|likes|like|goes|come|said|told)\b', text_l) or \
           re.search(r'\bthe\s+\w+\s+(?:which|who|whom|that|whose)\b', text_l) or \
           re.search(r'\bby which\b|\bin which\b|\bfor which\b|\bwith whom\b', text_l):
            if "定语从句" not in found:
                found.append("定语从句")
        # 状语从句（细分类型）
        if re.search(r'\bbecause\b|\bsince\b.*\bbecause\b', text_l):
            if "原因状语从句" not in found: found.append("原因状语从句")
        if re.search(r'\bif\b|\bunless\b|\bas long as\b|\bon condition that\b', text_l):
            if "条件状语从句" not in found: found.append("条件状语从句")
        if re.search(r'\bthough\b|\balthough\b|\beven if\b|\beven though\b|\bwhatever\b|\bwherever\b|\bhowever\b|\bno matter', text_l):
            if "让步状语从句" not in found: found.append("让步状语从句")
        if re.search(r'\bso that\b|\bin order that\b|\bto the end that\b', text_l):
            if "目的状语从句" not in found: found.append("目的状语从句")
        if re.search(r'\bso\s+\w+\s+that\b|\bsuch\s+\w+\s+that\b', text_l):
            if "结果状语从句" not in found: found.append("结果状语从句")
        if re.search(r'\bwhen\b|\bwhile\b|\bbefore\b|\bafter\b|\bas\b|\buntil\b|\btill\b|\bwhenever\b|\bsoon after\b|\bhardly.*when', text_l):
            if "时间状语从句" not in found: found.append("时间状语从句")
        if re.search(r'\bwhere\b|\bwherever\b', text_l) and 'the' in text_l:
            if "地点状语从句" not in found: found.append("地点状语从句")
        # 同位语从句（that 引导，前面常是 news/fact/idea/belief/hope）
        if re.search(r'\b(?:news|fact|idea|belief|hope|thought|suggestion|advice|message|sign|report|truth)\s+that\b', text_l):
            if "同位语从句" not in found:
                found.append("同位语从句")
        # 主语从句
        if re.search(r'\b(?:that|what|who|whom|whose|where|when|why|how|whether)\b.*\b(?:is|are|was|were|will be|becomes|seems|looks)\b', text_l) or \
           re.search(r'\bit\s+(?:is|was)\s+\w+\s+(?:that|what|whether)\b', text_l):
            if "主语从句" not in found: found.append("主语从句")
        # 表语从句
        if re.search(r'\b(?:is|are|was|were|become|become|remain|seem|look|feel|sound)\s+(?:that|what|who|whom|whose|where|when|why|how|whether)\b', text_l):
            if "表语从句" not in found: found.append("表语从句")

    # ====== C. 特殊句式（同样仅对英文句子生效） ======
    if has_english:
        # 强调句 It is/was ... that/who ...
        if re.search(r'\bit (?:is|was|has been)\s+(?:not)?\s+\w+\s+(?:that|who|whom)\b', text_l):
            if "强调句" not in found: found.append("强调句")
        # 倒装句（否定词/only/not only 提前）
        if re.search(r'\b(?:never|hardly|seldom|rarely|not only|only by|only when|only after|only then|scarcely)\b', text_l) and \
           re.search(r'\b(?:have|has|had|did|do|does|is|are|was|were|will|can|could|should|would)\s+\w+\b', text_l):
            if "倒装句" not in found: found.append("倒装句")
        # 感叹句
        if (re.search(r'\bwhat\s+(?:a|an|the)?\s+\w+', text_l) and '!' in text) or \
           re.search(r'\bhow\s+\w+\s+!', text_l):
            if "感叹句" not in found: found.append("感叹句")
        # not only...but also...
        if re.search(r'\bnot only\b', text_l) and re.search(r'\bbut also\b|\bbut\s*\w+\s*also\b', text_l):
            if "not only...but also..." not in found: found.append("not only...but also...")
        # neither...nor... / either...or...
        if re.search(r'\bneither\b.*\bnor\b', text_l):
            if "neither...nor..." not in found: found.append("neither...nor...")
        if re.search(r'\beither\b.*\bor\b', text_l):
            if "either...or..." not in found: found.append("either...or...")
        # too...to...
        if re.search(r'\btoo\s+\w+\s+to\s+\w+', text_l):
            if "too...to..." not in found: found.append("too...to... 句型")
        # not...until...
        if re.search(r'\b(?:not|never|hardly)\b.*\buntil\b', text_l) or re.search(r"\bwon\'t\b.*\buntil\b", text_l):
            if "not...until..." not in found: found.append("not...until... 句型")
        # so...that... / such...that...
        if re.search(r'\bso\s+\w+\s+that\b', text_l):
            if "so...that..." not in found: found.append("so...that... 句型")
        if re.search(r'\bsuch\s+\w+\s+that\b', text_l):
            if "such...that..." not in found: found.append("such...that... 句型")
        # would rather do A than do B
        if re.search(r'\bwould rather\s+\w+.*\bthan\b', text_l):
            if "would rather...than..." not in found: found.append("would rather...than...")
        # It takes sb. ... to do
        if re.search(r'\bit takes?\s+\w+\s+(?:\w+\s+)?\w+\s+to\s+\w+', text_l):
            if "It takes sb. ... to do" not in found: found.append("It takes sb. ... to do")
        # 主将从现（if 条件句 + will）
        if re.search(r'\bif\b.*\b(?:will|won\'t|shall)\b', text_l):
            if "主将从现（if 条件句）" not in found: found.append("主将从现（if 条件句）")
        # had better do
        if re.search(r'\bhad better\b', text_l):
            if "had better do" not in found: found.append("had better do")
        # used to do / be used to doing / be used to do
        if re.search(r'\bused to\b', text_l):
            if "used to do / be used to" not in found: found.append("used to do / be used to")

        # ====== D. 固定搭配结构 ======
        # one of + 复数名词
        if re.search(r'\bone of\s+\w+s\b|\bone of\s+the\s+\w+s\b', text_l):
            if "one of + 复数名词（最高级语境）" not in found:
                found.append("one of + 复数名词（最高级语境）")
        # with + n. + doing/done（伴随结构）
        if re.search(r'\bwith\s+\w+\s+\w+(?:ing|ed|en)\b', text_l):
            if "with 复合结构（伴随状语）" not in found:
                found.append("with 复合结构（伴随状语）")
        # enough to do
        if re.search(r'\benough\s+(?:\w+\s+)?to\s+\w+', text_l):
            if "enough...to..." not in found: found.append("enough...to... 句型")
        # so as to / in order to
        if re.search(r'\bso as to\b|\bin order to\b', text_l):
            if "so as to / in order to 目的状语" not in found:
                found.append("so as to / in order to 目的状语")

    return found


def detect_clauses(text):
    """保留给旧调用，内部用 detect_all_patterns 过滤出从句类"""
    all_found = detect_all_patterns(text)
    return [x for x in all_found
            if "从句" in x or "状语" in x or "同位语" in x.split("（")[0]]


def detect_special_sentences(text):
    """保留给旧调用，内部用 detect_all_patterns 过滤出特殊句式"""
    all_found = detect_all_patterns(text)
    return [x for x in all_found
            if any(k in x for k in ["强调", "倒装", "感叹", "too...to", "not...until",
                                    "would rather", "It takes", "主将从现", "not only",
                                    "neither", "either", "so...that", "such...that",
                                    "had better", "used to", "enough...to", "so as to",
                                    "with 复合", "one of"])]


def detect_modal_verbs(text):
    """检测情态动词"""
    text_l = text.lower()
    found = []
    modal_map = {
        "can/could": [r'\bcan\s+\w+\b', r'\bcould\s+\w+\b'],
        "may/might": [r'\bmay\s+\w+\b', r'\bmight\s+\w+\b'],
        "must": [r'\bmust\s+\w+\b'],
        "should": [r'\bshould\s+\w+\b'],
        "need": [r'\bneed\s+\w+\b'],
        "have to": [r'\bhave to\b', r'\bhas to\b'],
    }
    for modal, patterns in modal_map.items():
        for p in patterns:
            if re.search(p, text_l):
                found.append(modal)
                break
    return found


def detect_comparisons(text):
    """检测比较级/最高级/同级比较"""
    text_l = text.lower()
    found = []
    if re.search(r'\b\w+er\s+than\b', text_l) or re.search(r'\bmore\s+\w+\s+than\b', text_l):
        found.append("比较级")
    if re.search(r'\bthe\s+\w+est\b', text_l) or re.search(r'\bthe most\s+\w+\b', text_l):
        found.append("最高级")
    if re.search(r'\bas\s+\w+\s+as\b', text_l):
        found.append("as...as 同级比较")
    return found


# ============================================================
#  人教版 7-9 年级核心考点规则库
#  每条： (考点名, [正则], 做题提示)
#  顺序：先具体后笼统，避免"一般过去时"吞掉所有考点
# ============================================================
PAST_PARTICIPLES = (r'been|done|gone|written|taken|given|known|seen|shown|driven|grown|'
                    r'flown|fallen|broken|chosen|spoken|stolen|woken|ridden|risen|beaten|'
                    r'eaten|forgotten|hidden|forbidden|built|bought|brought|caught|taught|'
                    r'thought|fought|sought|made|paid|sent|spent|lost|won|found|left|kept|'
                    r'heard|felt|meant|met|read|run|come|become|begun|drunk|swum')

GRAMMAR_POINT_RULES = [
    ("现在完成时", [
        rf'\b(?:have|has)\s+(?:{PAST_PARTICIPLES})\b',
        r'\b(?:already|yet|ever|never|just)\b',
        r'\bsince\s+(?:\w+|\d{4})\b',
        r'\bfor\s+(?:a|an|two|three|several|\d+)\s+\w+',
        r'\bso far\b|\bin the past\b|\brecently\b',
    ], "抓 have/has + 过去分词；圈 already/yet/since/for 等标志词"),
    ("过去完成时", [
        rf'\bhad\s+(?:{PAST_PARTICIPLES})\b',
        r'\bby the time\b|\bby the end of\b',
    ], "抓 had + 过去分词；by the time / by the end of 是强提示"),
    ("过去进行时", [
        r'\b(?:was|were)\s+\w+ing\b',
        r'\b(?:at that time|at this time yesterday|while)\b',
    ], "抓 was/were + doing；while/at that time 提示背景动作"),
    ("现在进行时", [
        r'\b(?:am|is|are)\s+\w+ing\b',
        r'\b(?:now|look|listen|at the moment)\b',
    ], "抓 am/is/are + doing；now/Look/Listen 是标志词"),
    ("一般将来时", [
        r'\b(?:will|shall)\s+\w+',
        r'\bbe going to\s+\w+',
        r'\b(?:tomorrow|next\s+(?:week|month|year|Monday)|soon|in the future)\b',
    ], "抓 will / be going to；注意 tomorrow/next week 等将来标志"),
    ("一般过去时", [
        r'\b(?:yesterday|last\s+(?:week|month|year|night|weekend)|just now|ago|in \d{4})\b',
        r'\b(?:went|came|took|gave|saw|made|said|told|found|got|bought|thought|knew|'
        r'grew|flew|drove|rode|spoke|broke|chose|wrote|fell|ate|forgot|spent|sent|'
        r'lent|built|meant|paid|stood|understood|held|sold|won|shot|lost|kept|left)\b',
    ], "抓 yesterday/last.../ago 等过去时间；动词用过去式"),
    ("一般现在时", [
        r'\b(?:always|usually|often|sometimes|never|every\s+(?:day|week|year|morning))\b',
        r'\b(?:do|does)\s+\w+',
    ], "抓 always/usually/every day；主语三单时动词加 s/es"),
    ("被动语态", [
        r'\b(?:am|is|are|was|were|be|been|being)\s+(?:not\s+)?\w+(?:ed|en)\b',
        r'\b(?:must|can|should|may|might|could)\s+be\s+\w+(?:ed|en)\b',
        r'\b(?:have|has|had)\s+been\s+\w+(?:ed|en)\b',
        r'\bwill\s+be\s+\w+(?:ed|en)\b',
    ], "先判断主语是动作承受者，再用 be + 过去分词，注意时态一致"),
    ("非谓语动词", [
        r'\bto\s+(?:do|be|learn|study|finish|get|go|come|know|see|help|keep|'
        r'remember|forget|decide|want|hope|refuse|agree|ask|tell|advise|encourage|'
        r'invite|expect|promise|plan|choose|afford|offer|manage|pretend|fail|'
        r'happen|seem|appear|wish|improve|close|complete|realize|reach)\b',
        r'\b(?:enjoy|finish|practice|mind|avoid|consider|suggest|keep|imagine|'
        r'miss|give up|look forward to|be busy|feel like|can\'t help)\s+\w+ing\b',
        r'\b(?:is|are|was|were|be|been|feel|felt|look|looked|seem|seemed|become|became)\s+\w+(?:ed|en)\b',
    ], "看空格前后结构：目的/将来用 to do，介词后/习惯搭配用 doing，被动/完成用 done"),
    ("宾语从句", [
        r'\b(?:think|believe|hope|wonder|know|say|said|ask|asked|tell|told|'
        r'find|found|guess|remember|forget)\s+(?:that|if|whether|what|where|when|why|how|who)\b',
        r'\bdo you (?:think|know|believe)\b',
        r'\bcould you tell me\b',
    ], "注意引导词（that/if/whether/疑问词）和陈述语序，不能再用疑问语序"),
    ("定语从句", [
        r'\b(?:which|who|whom|whose|that)\s+(?:is|are|was|were|has|have|will|can|do|does|did|likes|like)\b',
        r'\bthe\s+(?:\w+\s+){0,2}(?:which|who|whom|that|whose)\b',
    ], "先找先行词，再看关系词在从句中作主语/宾语/定语"),
    ("状语从句", [
        r'\b(?:because|although|though|since|unless|as long as|so that|in order that|'
        r'even if|as soon as|while|when|before|after)\b',
        r'\bso\s+\w+\s+that\b|\bsuch\s+\w+\s+that\b',
    ], "判断是时间/原因/条件/让步/结果状语从句，再选引导词"),
    ("情态动词", [
        r'\b(?:can|could|may|might|must|should|need|have to|has to)\s+\w+',
        r'\bmustn\'t\b|\bdon\'t have to\b|\bhad better\b',
    ], "区分 must（必须/肯定推测）与 can't（否定推测）；注意后接动词原形"),
    ("比较级与最高级", [
        r'\b\w+er\s+than\b|\bmore\s+\w+\s+than\b',
        r'\bthe\s+\w+est\b|\bthe most\s+\w+\b',
        r'\bas\s+\w+\s+as\b|\bless\s+\w+\s+than\b',
    ], "than 前用比较级，the + 范围用最高级，as...as 中间用原级"),
    ("名词单复数", [
        r'\b(?:two|three|four|five|six|seven|eight|nine|ten|many|several|few|'
        r'a few|some|lots of|a lot of|plenty of)\s+\w+s\b',
        r'\b(?:hundreds of|thousands of|millions of)\b',
    ], "数字/many/some 后接可数名词复数；注意不规则复数"),
    ("代词", [
        r'\b(?:someone|somebody|anyone|anybody|everyone|everybody|nobody|'
        r'something|anything|everything|nothing)\b',
        r'\b(?:both|either|neither|none|other|another|others|the other)\b',
        r'\b(?:myself|yourself|himself|herself|itself|ourselves|yourselves|themselves)\b',
    ], "不定代词看范围（both/either/neither/none）；反身代词看主语一致"),
    ("冠词", [
        r'\b(?:a|an)\s+\w+',
        r'\bthe\s+\w+',
        r'\b(?:play|playing)\s+the\s+\w+',
    ], "元音音素前用 an；特指/独一无二/序数词前用 the；泛指用 a/an"),
    ("介词与介词短语", [
        r'\b(?:in|on|at|by|with|for|from|about|of|to|into|across|through|during|'
        r'without|after|before|between|among)\s+(?:the|a|an|my|his|her|their|our|your|\w+ing)\b',
        r'\b(?:be interested in|be good at|be afraid of|be proud of|be famous for|'
        r'be strict with|be full of|be made of|be made from)\b',
    ], "记固定搭配：in the morning / on Monday / at night；形容词+介词要成组记"),
    ("连词", [
        r'\b(?:both|either|neither|not only)\s+\w+\s+(?:and|or|nor|but)\b',
        r'\b(?:and|but|or|so|because|although|while)\b',
    ], "看前后逻辑：并列/转折/选择/因果，再选连词"),
    ("主谓一致", [
        r'\b(?:one of|each of|every one of|the number of)\b',
        r'\b(?:each|every|either|neither|nobody|everybody|somebody)\b',
    ], "one of + 复数名词作主语，谓语用单数；the number of 用单数，a number of 用复数"),
    ("特殊句式", [
        r'\bwhat\s+(?:a|an)\s+\w+|\bhow\s+\w+\s+!',
        r'\bit (?:is|was)\s+\w+\s+(?:that|who)\b',
        r'\b(?:never|hardly|seldom|not only|only when|only after)\b',
        r'\btoo\s+\w+\s+to\b|\bnot\s+\w+\s+until\b|\bit takes?\b',
    ], "感叹句看 what/how；强调句看 It is/was...that；倒装看否定词提前"),
    ("动词短语与固定搭配", [
        r'\b(?:look|take|give|come|get|put|turn|keep|go|set|make|bring|call|'
        r'pick|show|work|hand|break|find|deal|depend|belong|stick|lead|refer)\s+'
        r'(?:up|down|on|off|out|in|away|back|over|through|to|for|with|about|after)\b',
        r'\b(?:be used to|used to|look forward to|pay attention to|make a difference)\b',
    ], "动词短语按核心动词归类记忆，注意 up/off/out 等小品词改变词义"),
    ("情景交际", [
        r'\b(?:would you like|why not|how about|what about|shall we|'
        r'i\'d love to|sounds good|no problem|never mind|with pleasure)\b',
    ], "先判断语境（邀请/建议/道歉/感谢），再选符合英语习惯的答语"),
]


def detect_grammar_points(text, limit=None):
    """按规则库统计考点，返回 [{'name', 'count', 'tip'}]，按命中次数排序"""
    text_l = (text or '').lower()
    scored = []
    for name, patterns, tip in GRAMMAR_POINT_RULES:
        hits = 0
        for p in patterns:
            hits += len(re.findall(p, text_l, re.IGNORECASE))
        if hits:
            scored.append({"name": name, "count": hits, "tip": tip})
    scored.sort(key=lambda x: -x["count"])
    return scored[:limit] if limit else scored


NOISY_POINTS = {"冠词", "介词与介词短语", "连词", "代词", "名词单复数", "一般现在时"}


def first_grammar_point(text, fallback_name="词汇/短语辨析"):
    """逐题判断主考点：按规则库顺序取第一个明确考点，跳过过于宽泛的词法项"""
    text_l = (text or '').lower()
    fallback = None
    for name, patterns, tip in GRAMMAR_POINT_RULES:
        if any(re.search(p, text_l, re.IGNORECASE) for p in patterns):
            if fallback is None:
                fallback = (name, tip)
            if name not in NOISY_POINTS:
                return name, tip
    if fallback_name:
        return fallback_name, "先看选项的词性和搭配差异，再用代入法排除"
    if fallback:
        return fallback
    return None, ""


def extract_question_items(text):
    """按题号拆成 (题号, 题干)，只保留连续递增的题号序列。
    适合翻译/语法填空等逐题分析，并自动截断答案横线和串进来的无关段落。
    """
    flat = re.sub(r'\s+', ' ', text or '')
    marks = list(re.finditer(r'(?<![\d.])(\d{1,2})\s*[.．、)）]\s*', flat))
    raw = []
    for i, m in enumerate(marks):
        num = int(m.group(1))
        end = marks[i + 1].start() if i + 1 < len(marks) else len(flat)
        seg = flat[m.end():end]
        # 去掉题号后紧跟的填空横线（如“4.______Tom...”），再截断末尾答案横线
        seg = seg.lstrip(' _')
        seg = re.split(r'_{4,}', seg)[0]
        seg = seg.strip()[:220]
        if len(seg) >= 5:
            raw.append((num, seg))

    # 找出最长的连续题号序列（51,52,53...），避免把年份/数字当题号
    best, cur = [], []
    for item in raw:
        if cur and item[0] == cur[-1][0] + 1:
            cur.append(item)
        else:
            if len(cur) > len(best):
                best = cur
            cur = [item]
    if len(cur) > len(best):
        best = cur
    return best if len(best) >= 2 else raw


def extract_numbered_items(block_text):
    """把板块内容按题号拆成单题，便于逐题判断考点"""
    items = extract_question_items(block_text)
    if items:
        return [t for _, t in items]
    flat = re.sub(r'\s+', ' ', block_text or '')
    parts = re.split(
        r'(?:^|\s)(?:[\(（]\s*[\)）]\s*)?(?:[\(（]\d{1,2}[\)）]|\d{1,2}\s*[\.．、])\s*',
        flat
    )
    return [p.strip() for p in parts if len(p.strip()) >= 6]


def summarize_exam_points(block_text, limit=5, fallback_name="词汇/短语辨析"):
    """逐题统计主考点，返回 ([(考点, 题数)], 题目数)"""
    items = extract_numbered_items(block_text)
    counter = {}
    if len(items) >= 3:
        for item in items:
            name, _ = first_grammar_point(item, fallback_name=fallback_name)
            if name:
                counter[name] = counter.get(name, 0) + 1
        if not counter:
            # 题目里没有明显语法特征时，退回整段检测
            for p in detect_grammar_points(block_text, limit=limit):
                counter[p["name"]] = p["count"]
        ranked = sorted(counter.items(), key=lambda x: -x[1])[:limit]
        return ranked, len(items)
    ranked = [(p["name"], p["count"]) for p in detect_grammar_points(block_text, limit=limit)]
    return ranked, len(items)


def build_points_advice(points, item_count, base_tip=""):
    """把考点分布拼成一句老师能直接用的分析"""
    if not points:
        return ""
    dist = "、".join(f"{name}{count}题" for name, count in points)
    tips = []
    for name, _ in points[:3]:
        for rule_name, _, tip in GRAMMAR_POINT_RULES:
            if rule_name == name and tip not in tips:
                tips.append(tip)
                break
    prefix = f"本次共{item_count}题，考点分布：{dist}。" if item_count >= 3 else f"本次主要涉及：{dist}。"
    return prefix + (base_tip or "") + ("重点方法：" + "；".join(tips) + "。" if tips else "")


def analyze_vocab_section(block_text):
    """单词板块分析"""
    pos_tags = re.findall(r'\b(?:n|v|adj|adv|prep|conj|num|art|pron|int)\.', block_text)
    pos_counts = {}
    for tag in pos_tags:
        pos_counts[tag] = pos_counts.get(tag, 0) + 1

    nums = re.findall(r'\b\d+\.\s', block_text)
    vocab_count = len(nums) or max(len(pos_tags) // 2, 1)

    if pos_counts:
        top = sorted(pos_counts.items(), key=lambda x: -x[1])[:3]
        main_pos = [POS_ZH.get(t[0], t[0]) for t in top]
        topic = "+".join(main_pos) + "核心词汇"
        advice = (f"本次约{vocab_count}个核心词，主要考" + "+".join(main_pos) +
                  "。每天朗读+默写15-20分钟，当天学的词当天过一遍，"
                  "用「三次记忆法」巩固（早中晚各过一遍），记得牢其他板块才上得去。")
    else:
        topic = "核心词汇"
        advice = (f"本次约{vocab_count}个核心词。每天朗读+默写15-20分钟，"
                  "用「三次记忆法」巩固（早中晚各过一遍），单词记得牢其他板块才能上去。")
    return {"topic": topic, "advice": advice}


def analyze_phrases_section(block_text):
    """短语板块分析"""
    nums = re.findall(r'\b\d+\.\s', block_text)
    phrase_count = len(nums) or 1

    cat = []
    if re.search(r'\b(?:look|take|give|come|get|put|keep|bring|turn|set|make|go|run|carry|call|cut|pick|show|work|hand|break|find)\b', block_text.lower()):
        cat.append("动词短语")
    if re.search(r'\b(?:as|in|on|at|by|for|of|to|with|from|above|after|before|into|out of)\b\s+\w+', block_text.lower()):
        cat.append("介词短语")
    if re.search(r'\b(?:above all|after all|in short|in a word|in fact|as a result|first of all|in one\'s case|by the way|in addition)\b', block_text.lower()):
        cat.append("固定搭配")

    topic = "+".join(cat) + "及固定搭配" if cat else "短语搭配"
    cat_str = "，涉及" + "、".join(cat) if cat else ""
    advice = (f"本次约{phrase_count}个短语{cat_str}。"
              "把短语按功能归类整理（带来/发生/受益/推迟等），搭配着记比单独背更牢固，"
              "每周末抽10分钟回顾一次，巩固记忆。")
    return {"topic": topic, "advice": advice}


def analyze_choice_section(block_text):
    """单选板块 - 逐题识别考点并统计分布"""
    points, item_count = summarize_exam_points(block_text, limit=5)
    if points:
        topic = "、".join(name for name, _ in points[:3]) + ("等" if len(points) > 3 else "")
        dist = "、".join(f"{name}{count}题" for name, count in points)
        tips = []
        for name, _ in points[:3]:
            for rule_name, _, tip in GRAMMAR_POINT_RULES:
                if rule_name == name:
                    tips.append(tip)
                    break
        advice = ((f"本次共{item_count}题，" if item_count >= 3 else "本次主要涉及")
                  + f"考点分布：{dist}。"
                  "做题四步：①圈题干里的时间、搭配、逻辑关键词；"
                  "②判断考什么（时态/语态/从句/词法）；③排除明显错误选项；"
                  "④把错题按考点归类。")
        if tips:
            advice += "重点方法：" + "；".join(tips) + "。"
        advice += "建议每天练5道，做完核对考点并归纳错题。"
        return {"topic": topic, "advice": advice}

    topic = "语法综合"
    advice = ("做题时先判断考点（时态/语态/从句/介词），再选最佳答案。"
              "建议每天练5道单选，做完后核对考点并归纳错题。")
    return {"topic": topic, "advice": advice}


def analyze_fill_blank_section(block_text):
    """填空板块 - 逐题识别词形变化考点"""
    points, item_count = summarize_exam_points(block_text, limit=5, fallback_name="词形变化")
    blank_count = len(re.findall(r'_{2,}', block_text))
    if blank_count >= 3:
        item_count = blank_count
    nonfinites = detect_nonfinite(block_text)
    passives = detect_passive_voice(block_text)

    labels = [name for name, _ in points]
    if not labels:
        labels = (nonfinites or []) + (["被动语态"] if passives else [])
    topic = "、".join(labels[:3]) + ("等" if len(labels) > 3 else "") if labels else "词形变化"

    dist = "、".join(f"{name}{count}题" for name, count in points) if points else "、".join(labels)
    advice = ("做题两步走：①定词性——看空格前后缺名词/动词/形容词/副词；"
              "②变词形——名词看单复数，动词看时态/语态/非谓语，形容词副词看比较级。")
    if dist:
        advice += (f"本次共{item_count}题，" if item_count >= 3 else "本次主要涉及：") + f"考点分布：{dist}。"
    advice += "建议每天练3道填空专项，每空旁标注词性和变形理由。"
    return {"topic": topic, "advice": advice}


def analyze_translation_section(block_text):
    """翻译板块 - 逐题识别每一句的句型结构，给出逐题清单+对应做法。"""
    # 只保留题干含中文的题（英文阅读段落混进来时自动剔除）
    items = [(n, t) for n, t in extract_question_items(block_text)
             if len(re.findall(r'[\u4e00-\u9fff]', t)) >= 3]
    if not items:
        items = [(i + 1, t) for i, t in enumerate(extract_numbered_items(block_text))]

    # 每个句型对应的“落笔方法”——讲具体怎么写，而不只是报一个语法名称
    PATTERN_METHODS = {
        "宾语从句": "主句 + 引导词(that/if/whether/what…) + 陈述语序从句，从句不能再用疑问语序",
        "定语从句": "先行词 + 关系词(who/which/that/whose) + 缺主/宾的从句；人用 who，物用 which",
        "时间状语从句": "when/while/before/after/as soon as + 从句，注意主将从现",
        "条件状语从句": "if/unless + 一般现在时从句，主句用 will",
        "原因状语从句": "because/since/as 引导原因，because 不能和 so 同时出现",
        "让步状语从句": "though/although/even though 引导让步，不能再加 but",
        "结果状语从句": "so + 形容词/副词 + that，或 such + 名词短语 + that",
        "目的状语从句": "so that / in order that + 从句，从句常带 can/could",
        "地点状语从句": "where 引导，前面不要加介词",
        "同位语从句": "先写抽象名词(news/fact/idea/belief)，再 that + 完整句子解释其内容，从句用陈述语序",
        "主语从句": "that 引导时不能省；长主语可用 it 作形式主语，真正主语后置",
        "表语从句": "系动词后接 that/why/how 等引导的从句，that 不能省",
        "强调句": "It is/was + 被强调部分 + that/who + 其余部分；去掉框架句子仍完整",
        "倒装句": "否定词/only/not only 提前时，把助动词(do/does/did/have/is)提到主语前",
        "感叹句": "What + a/an + 形容词 + 单数名词；How + 形容词/副词",
        "反意疑问句": "前肯后否、前否后肯，助动词和主语与主句保持一致",
        "not only...but also...": "not only 放句首要部分倒装；连接两个并列成分时形式要一致",
        "neither...nor...": "neither + 助动词 + 主语 + nor + 另一主语；谓语就近一致",
        "either...or...": "either...or... 连接并列成分，谓语就近一致",
        "too...to... 句型": "too + 形容词/副词 + to do，to do 含否定意义，译作“太……而不能”",
        "not...until... 句型": "not...until... 译作“直到……才”；Not until 放句首要倒装",
        "so...that... 句型": "so + 形容词/副词 + that + 从句，结果状语",
        "such...that... 句型": "such + a/an + 形容词 + 单数名词 + that + 从句",
        "would rather...than...": "would rather do A than do B，前后都用动词原形",
        "It takes sb. ... to do": "It takes sb. + 时间/精力 + to do sth.，it 是形式主语",
        "主将从现（if 条件句）": "if 从句用一般现在时，主句用 will/won't",
        "one of + 复数名词（最高级语境）": "one of the + 最高级 + 复数名词，谓语动词用单数",
        "with 复合结构（伴随状语）": "with + 宾语 + doing(主动)/done(被动)/形容词/介词短语",
        "with 伴随结构": "with + 宾语 + doing(主动)/done(被动)/形容词/介词短语，表示伴随状态",
        "one of + 复数名词": "one of the + 最高级 + 复数名词，谓语动词用单数",
        "enough...to... 句型": "形容词/副词 + enough + to do；enough 修饰形容词时后置",
        "so as to / in order to 目的状语": "so as to / in order to + 动词原形，表目的",
        "used to do / be used to": "used to do 过去常做；be used to doing 习惯于；be used to do 被用来做",
        "had better do": "had better + 动词原形，否定用 had better not do",
    }

    def _method_of(name):
        if name in PATTERN_METHODS:
            return PATTERN_METHODS[name]
        for key, val in PATTERN_METHODS.items():
            if name.startswith(key) or key.startswith(name):
                return val
        return ""

    def _short_prompt(txt):
        """截一段题干做提示，让每条建议对应到具体句子（清掉标注、英文提示和结构标签）"""
        t = re.sub(r'[（(][^）)]{0,40}[）)]', ' ', txt)
        t = re.sub(r'^\s*\d+\s*[.、．]?\s*', '', t)
        t = re.sub(r'_{3,}', ' ', t)
        t = re.sub(r'同位语从句|同位语|宾语从句|定语从句|状语从句|主语从句|表语从句|'
                   r'宾从|定从|状从|引导的时间状从|时间状从|强调句|倒装句|感叹句|结构', ' ', t)
        t = re.sub(r'[A-Za-z]+', ' ', t)   # 英文提示（capital/with/oneof/whenever）不是题干
        # 丢掉孤立的注释小词（如“省会”“毛发”），只保留成句的中文
        segs = [s for s in re.split(r'\s+', t)
                if len(re.findall(r'[\u4e00-\u9fff]', s)) >= 3]
        t = ''.join(segs)
        t = t.strip('。，、；：,.;:·-—')
        if len(re.findall(r'[\u4e00-\u9fff]', t)) < 6:
            return ""
        return t[:28] + ("…" if len(t) > 28 else "")

    per_q, unique = [], []
    for num, txt in items[:10]:
        pats = detect_all_patterns(txt)
        # 每题最多展示 2 个结构，避免堆砌
        pats = pats[:2] if pats else ["基础句型（主谓宾/时态）"]
        prompt = _short_prompt(txt)
        head = f"第{num}题"
        if prompt:
            head += f"（{prompt}）"
        head += f"：{'、'.join(pats)}"
        methods = []
        for p in pats:
            m = _method_of(p)
            if m and m not in methods:
                methods.append(m)
            if p not in unique:
                unique.append(p)
        if methods:
            head += "——" + "；".join(methods[:2])
        per_q.append(head)

    topic = "、".join(unique[:3]) + ("等" if len(unique) > 3 else "") if unique else "句子翻译（基础句型）"

    advice_parts = ["逐题考点与写法：" + "；".join(per_q) + "。"]
    advice_parts.append("通用步骤：①圈出主干（谁+做什么）先译；②再挂修饰成分（从句/非谓语/介词短语）；"
                        "③按英文语序落笔；④最后核对时态、主谓一致、单复数和冠词。")
    if len(unique) >= 3:
        advice_parts.append(f"本次共涉及 {len(unique)} 类结构，说明不能只按一种句型套，必须逐句判断。")

    return {"topic": topic, "advice": "".join(advice_parts)}


def analyze_reading_section(block_text):
    """阅读理解板块 - 圈-找-选三步法
    核心做题逻辑（人教版）：
    ①圈——圈题干关键词（人名/时间/数字/否定词）
    ②找——回原文定位信息句，划出对应句
    ③选——对比选项与原文，排除明显错项
    """
    text_l = block_text.lower()
    # 检测题型
    question_types = []
    # 细节题
    if re.search(r'\b(?:according to (?:the )?(?:passage|text|article)|who\s+\?|where\s+\?|when\s+\?|what\s+\?|how many|how much|how long|how old)\b', text_l) or \
       re.search(r'(根据|哪一|什么|几个|多少|多长|多大|何时|何地|是谁)', block_text):
        question_types.append("细节题")
    # 推断题
    if re.search(r'\b(?:infer|imply|suggest|learn from|conclude|conclusion|attitude|tone|purpose of writing)\b', text_l) or \
       re.search(r'(推断|暗示|表明|可以得出|作者态度|写作目的)', block_text):
        question_types.append("推断题")
    # 主旨题
    if re.search(r'\b(?:main idea|main topic|best title|title is|mainly about|mainly discusses|the passage is mainly)\b', text_l) or \
       re.search(r'(主旨|大意|标题|最佳标题|主要讲|主要讨论)', block_text):
        question_types.append("主旨题")
    # 词义题
    if re.search(r'\b(?:the underlined (?:word|phrase|sentence)|the (?:word|phrase) "[^"]+" (?:means|refers to)|closest in meaning|replace[^,]*with)\b', text_l) or \
       re.search(r'(画线|下划线|词义|指代|替换|含义)', block_text):
        question_types.append("词义题")

    if not question_types:
        question_types = ["细节题", "推断题"]

    topic = "阅读理解（圈-找-选）"
    advice = (f"做题牢记三步法：①圈——圈题干关键词（人名/时间/数字/否定词），"
              f"明确题目问什么；②找——回原文定位信息句，划出对应句；"
              f"③选——对比选项与原文，排除绝对化、张冠李戴、无关项。"
              f"本次重点题型：{'、'.join(question_types)}。"
              "细节题找准定位句即可作答；推断题必须基于原文事实，不能凭主观感觉；"
              "主旨题看首尾段和各段首句；词义题代入选项看上下文是否通顺。"
              "建议每天精练1-2篇，限时5-7分钟/篇，做完后分析错因。")
    return {"topic": topic, "advice": advice}


def analyze_six_choice_section(block_text):
    """六选五/七选五板块 - 识别前后文逻辑关系
    核心做题逻辑（人教版）：抓逻辑衔接点
    ①代词指代（he/it/this/that 指代上文人物或事物）
    ②连词（however/so/but/and 判断前后是转折/因果/并列）
    ③顺序词（first/then/next 梳理段落顺序）
    """
    text_l = block_text.lower()
    logic_features = []
    # 代词指代
    if re.search(r'\b(?:he|she|it|they|we|you|this|that|these|those|such|both|either|neither|one|ones)\b', text_l):
        logic_features.append("代词指代")
    # 转折关系
    if re.search(r'\b(?:however|but|although|though|while|whereas|on the contrary|in contrast|nevertheless)\b', text_l):
        logic_features.append("转折关系")
    # 因果关系
    if re.search(r'\b(?:therefore|so|thus|because|since|as a result|consequently|due to|thanks to|owing to)\b', text_l):
        logic_features.append("因果关系")
    # 顺序关系
    if re.search(r'\b(?:first(?:ly)?|second(?:ly)?|then|next|after that|before|after|finally|last(?:ly)?|in the end)\b', text_l):
        logic_features.append("顺序关系")
    # 举例关系
    if re.search(r'\b(?:for example|for instance|such as|namely|like|take[^.]*for example)\b', text_l):
        logic_features.append("举例关系")
    # 解释/总结关系
    if re.search(r'\b(?:in addition|in other words|in short|in a word|in conclusion|in fact|above all|after all|all in all)\b', text_l):
        logic_features.append("解释/总结关系")

    if not logic_features:
        logic_features = ["代词指代", "逻辑衔接词"]

    topic = "六选五/七选五（逻辑衔接）"
    advice = (f"做题先通读全文抓大意，再重点找「逻辑衔接点」："
              f"①代词（he/it/this/that）——指代上文人物或事物，看空格前后是否复现；"
              f"②连词（however/so/but/and）——判断前后是转折/因果/并列；"
              f"③顺序词（first/then/next）——梳理段落顺序。"
              f"本次涉及：{'、'.join(logic_features)}。"
              "选不出时回看空格前后句的主语和复现词（同根词/同义替换），"
              "排除与上下文逻辑矛盾的选项。建议每天练1-2篇，做完梳理逻辑链条。")
    return {"topic": topic, "advice": advice}


def analyze_cloze_section(block_text):
    """完形填空板块 - 固定搭配
    核心做题逻辑（人教版）：
    ①先通读全文抓主旨和大意（不要急着看选项）
    ②逐空找线索——前后句逻辑+固定搭配+语境提示
    ③选完再读一遍验证通顺
    """
    text_l = block_text.lower()
    collocations = []
    # 动词短语（人教版7-9年级常见搭配）
    verb_phrases = [
        "look at", "look for", "look after", "look up", "look forward to",
        "take care of", "take part in", "take off", "take place", "take away",
        "give up", "give back", "give away", "give out",
        "come true", "come up with", "come across", "come back",
        "get on with", "get ready for", "get up", "get rid of", "get lost",
        "put on", "put up", "put off", "put away", "put down",
        "turn on", "turn off", "turn down", "turn up",
        "depend on", "belong to", "stick to", "lead to", "refer to",
        "be interested in", "be good at", "be afraid of", "be proud of",
        "be strict with", "be strict in", "be famous for", "be late for",
        "be made of", "be made from", "be made in", "be made up of",
        "be filled with", "be covered with", "be used to", "be used for",
        "keep on", "keep up with", "keep away from",
        "go on", "go through", "go over", "go over", "go ahead",
        "think about", "think of", "think over",
        "hear from", "hear of",
        "worry about", "wait for", "laugh at", "listen to",
        "deal with", "catch up with", "drop by", "drop in",
    ]
    for phrase in verb_phrases:
        if re.search(r'\b' + re.escape(phrase) + r'\b', text_l):
            collocations.append(phrase)
    # 介词搭配（形容词+介词）
    if re.search(r'\b(?:good|bad|interested|strict|famous|afraid|proud|full|tired|fond|sick|kind|polite|pleased|angry|surprised|excited|bored)\s+(?:of|in|at|with|for|about|from|to)\b', text_l):
        collocations.append("形容词+介词")

    topic = "完形填空（固定搭配）"
    if collocations:
        collocations_str = "、".join(collocations[:5])
        advice = (f"做题三步走：①先通读全文抓主旨和大意（不要急着看选项）；"
                  f"②逐空找线索——前后句逻辑+固定搭配+语境提示；"
                  f"③选完再读一遍验证通顺。本次重点搭配：{collocations_str}。"
                  "固定搭配题直接选搭配词；语境题先判断作者态度（褒贬），"
                  "再排除不符合作者情感的选项；遇到动词题注意上下文时态一致。"
                  "建议每天精练1篇（10-15分钟），做完积累本子记固定搭配。")
    else:
        advice = ("做题三步走：①先通读全文抓主旨和大意（不要急着看选项）；"
                  "②逐空找线索——前后句逻辑+固定搭配+语境提示；"
                  "③选完再读一遍验证通顺。"
                  "固定搭配题直接选搭配词；语境题先判断作者态度（褒贬），"
                  "再排除不符合作者情感的选项；遇到动词题注意上下文时态一致。"
                  "建议每天精练1篇（10-15分钟），做完积累本子记固定搭配。")
    return {"topic": topic, "advice": advice}


def analyze_grammar_blanks(block_text):
    """逐空分析语法填空：识别每空原词、词性变化和预期答案。
    返回 [(空号, 原词, 考点, 示例答案)]。
    """
    flat = re.sub(r'\s+', ' ', block_text or '')
    blanks = []
    pat = re.compile(r'(\d{1,2})\s*[.．]?\s*_{3,}\s*(?:[\(（]\s*([A-Za-z]+)\s*[\)）])?')
    for m in pat.finditer(flat):
        num = int(m.group(1))
        base = (m.group(2) or '').strip()
        ctx = flat[max(0, m.start() - 45):m.start()]
        ctx_l = ctx.lower().replace(' ', '')
        b = base.lower()
        kind, example = "词形变化", ""

        if not base:
            if re.search(r'half$|a$|an$|the$', ctx_l):
                kind, example = "冠词", "an"
            elif re.search(r'(?:help\w{0,3}|with|in|on|at|for|to|of|from|about|by)$', ctx_l):
                kind, example = "介词", "with"
            elif re.search(r'[,，;；]$', ctx) or re.search(r'(?:and|but|or|so|because|if|when|while|though)$', ctx_l):
                kind, example = "连词", "but"
            else:
                kind, example = "冠词/介词/连词", ""
        else:
            if b in {'he', 'she', 'they', 'we', 'i'}:
                kind = "代词宾格"
                example = {'he': 'him', 'she': 'her', 'they': 'them', 'we': 'us', 'i': 'me'}.get(b, '')
            elif b in {'my', 'your', 'his', 'her', 'their', 'our'}:
                kind, example = "名词性物主代词", {'my': 'mine', 'your': 'yours', 'her': 'hers', 'our': 'ours', 'their': 'theirs'}.get(b, '')
            elif re.search(r'(?:oneof|many|several|alotof)\w*$', ctx_l):
                kind = "名词复数"
                example = b + ('es' if b.endswith(('s', 'x', 'ch', 'sh')) else 's')
            elif re.search(r'than$', ctx_l):
                kind, example = "比较级", b + 'er'
            elif re.search(r'(?:enjoy|finish|mind|practice|keep|avoid|suggest)', ctx_l):
                kind, example = "非谓语 doing", b[:-1] + 'ing' if b.endswith('e') else b + 'ing'
            elif re.search(r'(?:is|are|was|were|be|been)$', ctx_l):
                kind = "被动/进行时"
            elif re.search(r'to$', ctx_l):
                kind, example = "不定式 to do", "to " + b
            elif re.search(r'(?:something|anything|nothing|everything)$', ctx_l):
                kind, example = "形容词后置", {'interest': 'interesting', 'excite': 'exciting',
                                                'bore': 'boring', 'surprise': 'surprising'}.get(b, b + 'ing')
            elif re.search(r'(?:has|have|had)$', ctx_l):
                kind = "完成时/过去分词"
            elif re.search(r'(?:often|usually|always|sometimes|never)$', ctx_l):
                kind, example = "动词三单", b + ('es' if b.endswith(('s', 'x', 'ch', 'sh', 'o')) else 's')
            elif b in {'loud', 'careful', 'quick', 'slow', 'happy', 'sudden', 'final', 'quiet',
                       'clear', 'safe', 'easy', 'real', 'bad', 'good', 'beautiful'}:
                kind = "副词"
                example = {'good': 'well', 'happy': 'happily', 'easy': 'easily'}.get(b, b + 'ly')
            elif re.search(r'(?:lastmonth|lastyear|yesterday|ago|lastweek)', flat[max(0, m.start()-80):m.start()].lower()):
                kind = "一般过去时"
                example = {'win': 'won', 'go': 'went', 'see': 'saw', 'buy': 'bought',
                           'teach': 'taught', 'take': 'took'}.get(b, b + 'ed')
            else:
                kind = "词形变化"

        blanks.append((num, base, kind, example))

    # 只保留连续递增的空号（避免把年份等数字当空号）
    blanks.sort(key=lambda x: x[0])
    best, cur = [], []
    for item in blanks:
        if cur and item[0] == cur[-1][0] + 1:
            cur.append(item)
        else:
            if len(cur) > len(best):
                best = cur
            cur = [item]
    if len(cur) > len(best):
        best = cur
    return best if len(best) >= 3 else blanks


def analyze_grammar_fill_section(block_text):
    """语法填空板块 - 定词性 → 变词形
    核心做题逻辑（人教版）：
    ①定词性——空格前后缺什么词性？主语位缺n/谓语位缺v/修饰位缺adj或adv
    ②变词形——根据语境变形（时态/被动/非谓语/单复数/比较级）
    """
    forms = []
    blanks = analyze_grammar_blanks(block_text)
    points, item_count = summarize_exam_points(block_text, limit=5, fallback_name="词形变化")
    blank_count = len(re.findall(r'_{2,}', block_text))
    if blank_count >= 3:
        item_count = blank_count
    tenses = detect_tenses(block_text)
    if tenses:
        forms.append(f"时态（{'、'.join(tenses[:3])}）")
    passives = detect_passive_voice(block_text)
    if passives:
        forms.append("被动语态")
    nonfinites = detect_nonfinite(block_text)
    if nonfinites:
        forms.append("非谓语")
    text_l = block_text.lower()
    # 比较级/最高级
    if re.search(r'\b\w+er\s+than\b|\bmore\s+\w+\s+than\b|\bthe\s+\w+est\b|\bthe most\s+\w+\b', text_l):
        forms.append("比较级/最高级")
    # 名词单复数
    if re.search(r'\b(?:two|three|four|five|six|seven|eight|nine|ten|many|several|some|a few|few|a lot of|lots of|plenty of)\s+\w+s\b', text_l):
        forms.append("名词单复数")

    # 优先用逐空识别结果（最具体）
    if len(blanks) >= 3:
        counter = {}
        per_blank = []
        for num, base, kind, example in blanks[:12]:
            counter[kind] = counter.get(kind, 0) + 1
            desc = f"第{num}空"
            if base:
                desc += f"：{base}→{kind}" + (f"（如 {example}）" if example else "")
            else:
                desc += f"：{kind}" + (f"（如 {example}）" if example else "")
            per_blank.append(desc)
        ranked = sorted(counter.items(), key=lambda x: -x[1])
        topic = "、".join(n for n, _ in ranked[:3]) + ("等" if len(ranked) > 3 else "")
        advice = ("逐空考点：" + "；".join(per_blank) + "。"
                  "做题两步走：①定词性——看空格前后缺名词/动词/形容词/副词；"
                  "②变词形——名词看单复数，动词看时态/语态/非谓语，形容词副词看比较级。"
                  "建议每空旁标注词性和变形理由，做完对照答案核对。")
        return {"topic": topic, "advice": advice}

    # 用逐题统计的考点作为更细的分布
    detailed = "、".join(f"{name}{count}题" for name, count in points)
    topic_labels = [name for name, _ in points] or forms
    topic = "、".join(topic_labels[:3]) + ("等" if len(topic_labels) > 3 else "") if topic_labels else "语法填空"
    forms_str = detailed or ("、".join(forms) if forms else "时态/语态/非谓语")
    advice = (f"做题两步走：①定词性——空格前后缺什么？主语位缺名词n，谓语位缺动词v，"
              f"修饰位缺形容词adj或副词adv；②变词形——根据句子语境变形。"
              + (f"本次共{item_count}空，" if item_count >= 3 else "本次重点考点：")
              + f"{forms_str}。"
              "动词变形最常考：看时间标志词定时态（already/yet 配完成时，"
              "last year 配过去时，next week 配将来时）、看主被动关系定语态、"
              "看前后结构定非谓语（to do/doing/done）。"
              "名词看单复数（a/an/数字/复数修饰词），"
              "形容词看比较级/最高级（than/the+最高级）。"
              "建议每天练1篇（约8空），每空旁标注词性和变形理由，加深理解。")
    return {"topic": topic, "advice": advice}


def analyze_task_reading_section(block_text):
    """任务型阅读板块 - 定词性 + 圈-找-选
    核心做题逻辑（人教版）：
    ①定词性——分析空格所需词性（看空格前后结构）
    ②圈——圈题干/表格提示词关键词
    ③找——回原文定位信息句
    ④选——按词性变形后填入
    """
    text_l = block_text.lower()
    # 检测空格周围可能需要的词性
    pos_required = []
    # 名词位（冠词/物主代词/形容词后）
    if re.search(r'\b(?:the|a|an|my|his|her|their|our|your|this|that|these|those|some|any|many|much|no|each|every|all|another|other)\s+____?_', text_l) or \
       re.search(r'\b(?:the|a|an|my|his|her|their|our|your|this|that|some|any|many|much|no|each|every|all|another|other)\s+__\b', text_l):
        pos_required.append("名词")
    # 动词位（主语后）
    if re.search(r'\b(?:he|she|it|they|we|you|i|tom|lucy|jack|the (?:boy|girl|students?|teacher|children))\s+____?_', text_l) or \
       re.search(r'\b____?_\s+(?:a|an|the|my|his|her|their)\b', text_l):
        pos_required.append("动词")
    # 形容词位（be动词/系动词后）
    if re.search(r'\b(?:is|are|was|were|be|been|look|feel|seem|sound|taste|smell|get|become|turn|grow|keep|stay)\s+____?_', text_l):
        pos_required.append("形容词")
    # 副词位（very/quite/too/so 后）
    if re.search(r'\b(?:very|quite|too|so|rather|pretty|really)\s+____?_', text_l):
        pos_required.append("副词")
    # 数字位（how many/how much/数字位）
    if re.search(r'\bhow (?:many|much|long|old|far|often)\b[^.?]*____?_', text_l) or \
       re.search(r'\b____?_\s+(?:o\'?clock|years?|months?|weeks?|days?|hours?|minutes?|yuan|dollars?|kilos?|meters?|kilometers?)\b', text_l):
        pos_required.append("数字/数量")

    # 检测变形类型
    forms = []
    tenses = detect_tenses(block_text)
    if tenses:
        forms.append(f"时态（{'、'.join(tenses[:3])}）")
    nonfinites = detect_nonfinite(block_text)
    if nonfinites:
        forms.append("非谓语")
    passives = detect_passive_voice(block_text)
    if passives:
        forms.append("被动语态")
    if re.search(r'\b\w+er\s+than\b|\bmore\s+\w+\s+than\b|\bthe\s+\w+est\b|\bthe most\s+\w+\b', text_l):
        forms.append("比较级/最高级")
    if re.search(r'\b(?:two|three|four|five|six|seven|eight|nine|ten|many|several|some|a few|few|a lot of|lots of)\s+\w+s\b', text_l):
        forms.append("名词单复数")

    if not pos_required:
        pos_required = ["名词", "动词", "形容词"]

    topic = "任务型阅读（定词性+圈找选）"
    forms_str = "，涉及" + "、".join(forms) if forms else ""
    advice = (f"做题四步走：①定词性——分析空格前后结构判断空格需要什么词性"
              f"（{'、'.join(pos_required)}）；②圈——圈题干/表格提示词关键词；"
              f"③找——回原文定位信息句，划出对应句；"
              f"④选——从原文提取信息后按词性变形填入{forms_str}。"
              "动词变形最常考：看时间标志词定时态（already/yet 配完成时，"
              "last year 配过去时，next week 配将来时）、看主被动关系定语态、"
              "看前后结构定非谓语（to do/doing/done）。"
              "名词看单复数（a/an/数字/复数修饰词），"
              "形容词看比较级/最高级（than/the+最高级）。"
              "注意：每空一般不超过3词，原文长句要概括精简。"
              "建议每天练1篇，限时5-7分钟，做完对照原文核对词性和变形。")
    return {"topic": topic, "advice": advice}


def analyze_writing_section(block_text):
    """作文板块 - 审题-列提纲-写句子-检查
    核心做题逻辑（人教版）：标准写作四步流程
    ①审题——文体/时态/人称
    ②列提纲——按要点顺序列出（开头/主体/结尾）
    ③写句子——每要点扩展成1-2句，用连接词衔接
    ④检查——时态一致/主谓一致/单词拼写/字数
    """
    # 检测文体
    genres = []
    if re.search(r'(邮件|信件| Dear |Dear\s+Sir|Dear\s+Mr|email|e-mail|letter|邀请函|感谢信)', block_text, re.IGNORECASE):
        genres.append("邮件/书信")
    if re.search(r'(记叙文|记叙|narrate|yesterday|last week|once|story)', block_text, re.IGNORECASE):
        genres.append("记叙文")
    if re.search(r'(议论文|议论|argument|opinion|in my opinion|my view|think that|believe that)', block_text, re.IGNORECASE):
        genres.append("议论文")
    if re.search(r'(说明文|说明|introduce|explain|how to|such as|for example)', block_text, re.IGNORECASE):
        genres.append("说明文")
    if re.search(r'(日记|diary|today|i went|i had|sunday|saturday|weather|日期)', block_text, re.IGNORECASE):
        genres.append("日记")
    if not genres:
        genres = ["应用文"]

    # 检测字数要求
    word_count = ""
    m = re.search(r'(\d+)\s*(?:词|字|words|word)', block_text.lower())
    if m:
        word_count = f"，字数要求约{m.group(1)}词"

    # 检测可用的句式提示（用作高级句式建议）
    clauses = detect_clauses(block_text)
    nonfinites = detect_nonfinite(block_text)
    specials = detect_special_sentences(block_text)

    sentence_tips = []
    if clauses:
        sentence_tips.append("从句")
    if nonfinites:
        sentence_tips.append("非谓语")
    if specials:
        sentence_tips.append("特殊句式")
    if not sentence_tips:
        sentence_tips = ["定语从句", "宾语从句"]

    topic = "作文（审题-列提纲-写-检查）"
    advice = (f"写作四步走：①审题——文体（{'、'.join(genres)}）+ 时态"
              f"（记叙多用过去时，议论/说明用一般现在时）+ 人称（第一/第三）；"
              f"②列提纲——按要点顺序列出（开头/主体/结尾），每要点用关键词记下；"
              f"③写句子——每要点扩展成1-2句，用连接词衔接（and/but/so/what's more/besides/in addition）；"
              f"④检查——时态一致、主谓一致、单词拼写{word_count}。"
              f"得分技巧：用1-2个高级句型（{'/'.join(sentence_tips)}）"
              f"+ 过渡词（however/therefore/besides）让行文更连贯，"
              f"开头用引子句（There be/It is...）+ 结尾用总结句（All in all/In a word）。"
              f"建议每周精写2-3篇，写完对照范文修改并背诵亮点句。")
    return {"topic": topic, "advice": advice}


def analyze_generic_section(block_text, sec_name=""):
    """通用分析器"""
    points, item_count = summarize_exam_points(block_text, limit=5)
    if points:
        topic = "、".join(name for name, _ in points[:3]) + ("等" if len(points) > 3 else "")
        dist = "、".join(f"{name}{count}题" for name, count in points)
        advice = (f"本板块" + (f"共{item_count}题，" if item_count >= 3 else "")
                  + f"考点分布：{dist}。建议按考点归类错题，"
                  "每天针对最薄弱的一类练5题，连续一周就能看出提升。")
        return {"topic": topic, "advice": advice}

    topic = f"{sec_name}板块" if sec_name else "本板块知识"
    advice = "针对薄弱知识点做专项练习，每天坚持练几道题找找感觉。"
    return {"topic": topic, "advice": advice}


# 板块分析器映射
SECTION_ANALYZERS = {
    "单词": analyze_vocab_section,
    "词汇": analyze_vocab_section,
    "短语": analyze_phrases_section,
    "词组": analyze_phrases_section,
    "单选": analyze_choice_section,
    "单项选择": analyze_choice_section,
    "填空": analyze_fill_blank_section,
    "适当形式填空": analyze_fill_blank_section,
    "翻译": analyze_translation_section,
    "句子翻译": analyze_translation_section,
    # 阅读类
    "阅读": analyze_reading_section,
    "阅读理解": analyze_reading_section,
    # 任务型阅读
    "任务型阅读": analyze_task_reading_section,
    # 六选五/七选五
    "六选五": analyze_six_choice_section,
    "七选五": analyze_six_choice_section,
    "五选五": analyze_six_choice_section,
    # 完形填空
    "完形": analyze_cloze_section,
    "完形填空": analyze_cloze_section,
    # 语法填空
    "语法填空": analyze_grammar_fill_section,
    # 作文
    "作文": analyze_writing_section,
    "书面表达": analyze_writing_section,
    "写作": analyze_writing_section,
}


def analyze_pdf_content(text, section_names):
    """基于人教版初中英语语法体系，按板块分割试卷并分析考点。
    覆盖：8大时态、被动语态、非谓语、从句、特殊句式、词法。
    返回 {section_name: {topic, advice}}
    """
    result = {}

    # 1. 按罗马数字切分试卷文本
    section_blocks = split_pdf_by_sections(text)

    # 2. 把试卷板块映射到标准板块名（同一个板块的多段内容合并）
    sec_text_map = {}
    for title, content in section_blocks:
        key = map_to_section_key(title, content)
        if key:
            sec_text_map[key] = (sec_text_map.get(key, "") + "\n" + content).strip()

    # 3. 对每个板块做针对性分析
    # 兼容板块名既可以是字符串，也可以是 (名称, 满分) 元组
    whole_text_fallback = len(section_blocks) == 1 and section_blocks[0][0] == "全文"
    for item in section_names:
        sec_name = item[0] if isinstance(item, tuple) else item
        sec_label = sec_name
        key = canonical_section_name(sec_name)
        block_text = sec_text_map.get(key, "")
        if not block_text:
            if whole_text_fallback:
                block_text = text
            else:
                result[sec_label] = {
                    "topic": "该板块未在这份试卷中出现",
                    "advice": "这份PDF里没有识别到该板块的题目，"
                              "可能是这次考试没考，或标题格式特殊没被识别到。"
                              "可以先手动填写考点，或提供对应的试卷PDF再试。",
                }
                continue
        analyzer = SECTION_ANALYZERS.get(key) or SECTION_ANALYZERS.get(sec_name)
        if analyzer:
            result[sec_label] = analyzer(block_text)
        else:
            result[sec_label] = analyze_generic_section(block_text, sec_name)

    return result


# ============================================================
#  个性化内容生成
# ============================================================
# ============================================================
#  个性化评语库
#  每条评语含占位符：{name} {score} {total} {section} {lose} {rate}
#  用学生名字 hash 选变体，保证同一学生稳定、不同学生不同
# ============================================================


def _pick_variant(name, variants):
    """用名字的稳定哈希挑一条，保证同一学生每次结果一致、不同学生各不相同"""
    digest = hashlib.md5(str(name).encode('utf-8')).hexdigest()
    idx = int(digest[:8], 16) % len(variants)
    return variants[idx]


def _stable_num(text):
    """稳定的排序扰动值，让同分板块在不同学生报告里轮换展示"""
    return int(hashlib.md5(str(text).encode('utf-8')).hexdigest()[:8], 16)


# —— 不同总分档的开头评语（每种 4 条备选）——
OPENING_EXCELLENT = [  # 总分率 >= 0.85
    "{name}同学太棒啦！本次{test_name}拿了 <b>{score:g}</b> 分，稳居班级前列👏",
    "{name}这次发挥相当出色，{test_name} <b>{score:g}</b> 分，看出来平时没少下功夫！",
    "{name}今天必须重点表扬！{test_name}一举拿下 <b>{score:g}</b> 分，妥妥的班级领跑者✨",
    "{name}同学好样的！本次{test_name} <b>{score:g}</b> 分，稳、准、狠，继续保持！",
]
OPENING_GOOD = [  # 0.70-0.85
    "{name}同学本次{test_name}取得 <b>{score:g}</b> 分，表现不错，有扎实的英语基础！",
    "{name}这次考得挺好，{test_name} <b>{score:g}</b> 分，看得出来状态在线。",
    "{name}的英语底子挺扎实，{test_name} <b>{score:g}</b> 分，整体发挥良好。",
    "{name}同学加油！本次{test_name} <b>{score:g}</b> 分，稳中有进步的感觉。",
]
OPENING_MEDIUM = [  # 0.55-0.70
    "{name}同学本次{test_name}拿到 <b>{score:g}</b> 分，处于班级中等，还有明显上升空间。",
    "{name}这次{test_name} <b>{score:g}</b> 分，不突出但也不差，找准薄弱点就能再上一个台阶。",
    "{name}的英语水平目前在班级中等位置，{test_name} <b>{score:g}</b> 分，冲一冲就能挤进前列。",
    "{name}同学，{test_name} <b>{score:g}</b> 分，其实你有不少可圈可点的地方，薄弱环节补上来分数会好看很多。",
]
OPENING_WEAK = [  # < 0.55
    "{name}同学本次{test_name}考了 <b>{score:g}</b> 分，看得出基础还有些不牢，咱们一起查漏补缺。",
    "{name}这次{test_name} <b>{score:g}</b> 分，别灰心，分数低反而暴露了问题的根源，对症下药就能翻盘。",
    "{name}同学，{test_name} <b>{score:g}</b> 分，目前的主要任务是先把基础夯实，基础上来了一切都会好起来。",
    "{name}的{test_name} <b>{score:g}</b> 分确实不太理想，但这不是终点，是帮咱们找到起点的信号，一起加油！",
]


# —— 强项表扬（每种 3 条）——
STRENGTH_PRAISE_FULL = [  # 板块满分
    "{section}拿到了满分，太赞了！",
    "{section} <b>{score:g}/{total:g}</b> 拿下满分，这部分你是全班榜样！",
    "{section}模块 <b>{score:g}/{total:g}</b> 满分通过，知识点掌握得相当扎实。",
]
STRENGTH_PRAISE_HIGH = [  # 得分率 >= 0.8
    "{section}表现出色，只丢了 <b>{lose:g}</b> 分，值得继续保持！",
    "{section} <b>{score:g}/{total:g}</b> 做得相当不错，这是你的优势板块！",
    "{section}模块得分率 {rate:.0%}，看得出来这块你花了心思。",
]
STRENGTH_PRAISE_ANY = [  # 只要是 top 就说
    "特别是{section}，这里是你的拿分利器，继续保持这个势头！",
    "{section}是你最出彩的地方，分数很稳定，值得肯定。",
    "{section}让老师眼前一亮，平时的积累在这里体现出来了。",
]


# —— 弱项提醒（每种 3 条）——
WEAKNESS_REMIND_GAP = [  # 得分率 < 0.55
    "但{section}是明显短板，这次丢了 <b>{lose:g}</b> 分，需要赶紧补起来。",
    "{section}失分较多，{score:g}/{total:g}的得分率只有 {rate:.0%}，这是当务之急要补的。",
    "{section}这部分基础比较薄弱，建议优先找出来多加练习。",
]
WEAKNESS_REMIND_NORMAL = [  # 0.55-0.70
    "{section}还有提升空间，这次丢了 <b>{lose:g}</b> 分，注意分析错因。",
    "但{section}需要多花些时间，得分率 {rate:.0%}，再往上冲一冲。",
    "{section}失分主要出在细节上，平时多留意就能避免。",
]


# —— 收尾鼓励（优/良/中/差 各 3 条）——
ENDING_EXCELLENT = [
    "保持这个节奏，下次继续稳拿高分！",
    "坚持这么好的学习习惯，英语成绩只会越来越好。",
    "你已经在正确的路上大步前进啦，继续保持！",
]
ENDING_GOOD = [
    "再加一把劲，把薄弱的{section}也拉上来，冲刺满分不是梦！",
    "整体状态不错，把{section}的漏洞补上，排名还能往前冲。",
    "英语学习只要保持这个势头，进步是必然的。",
]
ENDING_MEDIUM = [
    "把{section}专项练起来，相信下次考试会有看得见的进步。",
    "每天坚持一点点，把{section}的坑填上，分数就能上一个台阶。",
    "找准节奏，把薄弱环节逐个攻破，下次一定能看到进步！",
]
ENDING_WEAK = [
    "先从词汇和基础句型开始，每天坚持背5个单词、1个句型，慢慢就能跟上。",
    "别着急，咱们从小目标开始（每天5个单词+1段小短文），积累起来就是质变。",
    "老师相信你有追上的能力，只要每天多花15分钟在英语上就够了。",
]


def gen_overall_eval(name, total, total_max, scores, sections, test_name):
    """基于学生个人板块强弱分布，生成完全个性化的评语。
    每个学生因板块得分不同 + 名字 hash 选变体，评语不会重复。
    """
    rate = total / total_max if total_max > 0 else 0
    # 板块得分率，同时带上原始分数和满分
    sec_details = []
    for sec_name, max_score in sections:
        got = scores.get(sec_name, 0)
        r = got / max_score if max_score > 0 else 0
        sec_details.append({
            "name": sec_name, "score": got, "max": max_score,
            "rate": r, "lose": max_score - got,
        })
    if not sec_details:
        return f"{name}同学本次{test_name}取得 {total:g} 分。"

    # 排序（同分板块按学生名扰动，让不同学生的优势板块展示有差异）
    sec_by_rate = sorted(sec_details, key=lambda x: (-x["rate"], _stable_num(name + x["name"])))
    strongest = sec_by_rate[0]
    second_strongest = sec_by_rate[1] if len(sec_by_rate) >= 2 else None
    weakest = sec_by_rate[-1]

    # 1) 开头
    if rate >= 0.85:
        opening = _pick_variant(name, OPENING_EXCELLENT).format(
            name=name, score=total, total=total_max, test_name=test_name)
    elif rate >= 0.70:
        opening = _pick_variant(name, OPENING_GOOD).format(
            name=name, score=total, total=total_max, test_name=test_name)
    elif rate >= 0.55:
        opening = _pick_variant(name, OPENING_MEDIUM).format(
            name=name, score=total, total=total_max, test_name=test_name)
    else:
        opening = _pick_variant(name, OPENING_WEAK).format(
            name=name, score=total, total=total_max, test_name=test_name)

    # 2) 强项表扬（区分满分 / 高分 / 普通 top）
    praise_parts = []
    for sec in [strongest, second_strongest]:
        if not sec:
            continue
        if sec["rate"] >= 1.0:
            tpl = _pick_variant(name + sec["name"], STRENGTH_PRAISE_FULL)
        elif sec["rate"] >= 0.8:
            tpl = _pick_variant(name + sec["name"], STRENGTH_PRAISE_HIGH)
        else:
            tpl = _pick_variant(name + sec["name"], STRENGTH_PRAISE_ANY)
        praise_parts.append(tpl.format(
            section=sec["name"], score=sec["score"], total=sec["max"],
            lose=sec["lose"], rate=sec["rate"]))
    praise = " ".join(praise_parts)

    # 3) 弱项提醒
    warn_parts = []
    # 只提醒得分率明显低于其他的那个最弱项（避免碎碎念）
    if weakest["rate"] < strongest["rate"] - 0.20:  # 弱项和强项差 20% 以上才说
        if weakest["rate"] < 0.55:
            tpl = _pick_variant(name + "w1", WEAKNESS_REMIND_GAP)
        else:
            tpl = _pick_variant(name + "w1", WEAKNESS_REMIND_NORMAL)
        warn_parts.append(tpl.format(
            section=weakest["name"], score=weakest["score"], total=weakest["max"],
            lose=weakest["lose"], rate=weakest["rate"]))
    warn = " ".join(warn_parts)

    # 4) 收尾
    # 收尾里可以提一下最弱板块作为下一步目标
    if rate >= 0.85:
        ending_pool = ENDING_EXCELLENT
    elif rate >= 0.70:
        ending_pool = ENDING_GOOD
    elif rate >= 0.55:
        ending_pool = ENDING_MEDIUM
    else:
        ending_pool = ENDING_WEAK
    ending_sec = weakest["name"] if warn_parts else strongest["name"]
    ending = _pick_variant(name + "e", ending_pool).format(section=ending_sec)

    # 组装
    parts = [opening]
    if praise:
        parts.append(praise)
    if warn:
        parts.append(warn)
    parts.append(ending)
    return " ".join(parts) + "💙"


# —— 个性化班主任寄语（更口语化，每条 4 种备选）——
TEACHER_MSG_GOOD = [  # >= 0.75
    "{name}同学，你的努力大家都看得到！继续保持这个势头，英语成绩稳中有升！",
    "{name}这次的表现很亮眼，老师为你骄傲！只要保持住，前途无量🌟",
    "{name}真厉害！{test_name}考得这么好，离不开平时的积累，继续加油！",
    "{name}同学好样的，你在英语学习上展现出的潜力让人印象深刻，继续冲！",
]
TEACHER_MSG_MEDIUM = [  # 0.55-0.75
    "{name}同学，底子不错，把薄弱的{section}补上就能再上一个台阶，加油！",
    "{name}这次发挥稳中有进步，{section}再好好练一练就能大幅提分，老师看好你！",
    "{name}的表现一直在稳步提升，但{section}这块还要多下点功夫，咱们一起加油！",
    "{name}同学其实很有潜力，就是{section}这块拖了点后腿，补上就厉害啦！",
]
TEACHER_MSG_WEAK = [  # < 0.55
    "{name}同学，先把基础打牢，每天坚持一点点积累，你会看到自己的变化的！",
    "{name}别着急，咱们从最基础的单词和句型开始补，一步一个脚印，肯定能跟上！",
    "{name}同学你并不笨，只是之前有些东西没跟上，咱们一起把基础补回来！",
    "{name}老师相信你能追上，咱们先定个小目标（每天5个单词+1个句型），坚持就行！",
]

ENGLISH_QUOTES = [
    '"Practice makes perfect." 熟能生巧，你的坚持终将换来质的飞跃。',
    '"Every expert was once a beginner." 每位高手都曾是初学者，坚持你就一定能行。',
    '"Slow and steady wins the race." 稳扎稳打，终会成功。',
    '"A journey of a thousand miles begins with a single step." 千里之行始于足下。',
    '"Fall seven times, stand up eight." 跌倒了就爬起来，坚持就是胜利。',
    '"Rome wasn\'t built in a day." 罗马不是一天建成的，学习英语也是如此。',
    '"Where there is a will, there is a way." 有志者，事竟成！',
    '"The best way to predict the future is to create it." 未来要靠自己创造！',
]


def gen_teacher_msg(name, total, total_max, scores=None, sections=None, test_name="",
                    english_quote=False, class_avg=None):
    """个性化班主任寄语：固定包含本次分数、最强板块、最弱板块、下一步目标。
    同一学生每次稳定，不同学生因分数/板块/姓名 hash 不同而不重复。
    """
    rate = total / total_max if total_max > 0 else 0

    details = []
    if scores and sections:
        for sec_name, max_score in sections:
            got = scores.get(sec_name, 0)
            r = got / max_score if max_score > 0 else 0
            details.append({"name": sec_name, "score": got, "max": max_score, "rate": r})
    details.sort(key=lambda x: (-x["rate"], _stable_num(name + x["name"])))
    strongest = details[0] if details else None
    weakest = details[-1] if details else None
    weakest_sec = weakest["name"] if weakest else None

    if rate >= 0.75:
        msg = _pick_variant(name, TEACHER_MSG_GOOD).format(
            name=name, section=weakest_sec or "", test_name=test_name)
    elif rate >= 0.55:
        msg = _pick_variant(name, TEACHER_MSG_MEDIUM).format(
            name=name, section=weakest_sec or "薄弱环节", test_name=test_name)
    else:
        msg = _pick_variant(name, TEACHER_MSG_WEAK).format(
            name=name, section=weakest_sec or "基础", test_name=test_name)

    # 数据句：把本次真实得分写进去，多种句式轮换，避免每份报告一个模板
    if strongest and weakest:
        ss = f"{strongest['score']:g}/{strongest['max']:g}"
        ws = f"{weakest['score']:g}/{weakest['max']:g}"
        data_tpls = [
            "这次{sn}拿到 {ss}，是本次最稳的一块；{wn} {ws}，是接下来要重点补的地方。",
            "从板块看，{sn} {ss} 最扎实；{wn} {ws} 还有提升空间，咱们就盯住它练。",
            "老师注意到你{sn}拿到了 {ss}，这块保持住；{wn}只拿到 {ws}，接下来重点攻它。",
            "这次你的{sn}很稳（{ss}），{wn}（{ws}）是下一步的突破口。",
        ]
        data = _pick_variant(name + "d", data_tpls).format(
            sn=strongest["name"], ss=ss, wn=weakest["name"], ws=ws)

        # 班级位置：让每个孩子看到的参照不同
        if class_avg is not None:
            diff = total - class_avg
            diff_r = round(diff)
            if diff_r >= 5:
                cmp_tpls = [
                    f"总分比班级平均分高 {diff_r:g} 分，已经在队伍前面了。",
                    f"高出班级平均 {diff_r:g} 分，这个优势要继续守住。",
                ]
            elif diff_r <= -5:
                cmp_tpls = [
                    f"总分比班级平均分低 {abs(diff_r):g} 分，差距主要就在{weakest['name']}。",
                    f"离班级平均还差 {abs(diff_r):g} 分，把{weakest['name']}补上来就能追近。",
                ]
            else:
                cmp_tpls = [
                    "总分和班级平均分基本持平，往上一步就差一个薄弱板块。",
                    "目前处在班级平均线附近，突破点很明确。",
                ]
            data += _pick_variant(name + "c", cmp_tpls)

        # 下一步目标
        target = min(total_max, total + 5)
        target_tpls = [
            "咱们先把下一个小目标定在 {target:g} 分，一步步来。",
            "下次先冲 {target:g} 分，老师陪你一起把这块补上。",
            "下一个目标不高：先到 {target:g} 分，重点就练{wn}。",
            "咱们把目标定在 {target:g} 分，主要靠{wn}多拿几分。",
        ]
        data += _pick_variant(name + "t", target_tpls).format(target=target, wn=weakest["name"])
        msg = msg + " " + data

    quote = _pick_variant(name + "q", ENGLISH_QUOTES) if english_quote else ""
    return msg, quote


def gen_badges(scores, sections):
    badges = []
    for sec_name, max_score in sections:
        if sec_name not in scores:
            continue
        r = scores[sec_name] / max_score
        if r >= 1.0:
            badges.append(("🌟", f"{sec_name}满分"))
        elif r >= 0.9:
            badges.append(("📚", f"{sec_name}达人"))
    if not badges:
        badges.append(("💪", "坚持之星"))
    return badges


def gen_suggestions(scores, sections, section_advice):
    ranked = sorted(sections, key=lambda x: scores.get(x[0], 0) / x[1])
    suggestions = []
    for sec_name, _ in ranked[:3]:
        info = section_advice.get(sec_name, DEFAULT_ADVICE)
        suggestions.append((info["topic"], info["advice"]))
    return suggestions


def star_rating(score, max_score):
    full = int(score / max_score * 5) if max_score > 0 else 0
    return "★" * full + "☆" * (5 - full)


def get_section_color(sec_name):
    """根据板块名返回主题色（背景色 + 左边框色）。
    不同板块用不同颜色区分，便于一眼识别薄弱板块。
    """
    if any(k in sec_name for k in ["单词", "词汇"]):
        return "#E3F2FD", "#42A5F5"  # 浅蓝
    if any(k in sec_name for k in ["短语", "词组"]):
        return "#E0F7FA", "#00ACC1"  # 浅青
    if any(k in sec_name for k in ["单项选择", "单选"]):
        return "#F3E5F5", "#8E24AA"  # 浅紫
    if "语法填空" in sec_name:
        return "#E0F2F1", "#00897B"  # 浅薄荷
    if any(k in sec_name for k in ["填空"]):
        return "#FFF3E0", "#FB8C00"  # 浅橙
    if any(k in sec_name for k in ["翻译"]):
        return "#E8F5E9", "#43A047"  # 浅绿
    if "任务型" in sec_name and "阅读" in sec_name:
        return "#F1F8E9", "#689F38"  # 浅草绿（区别于普通阅读）
    if any(k in sec_name for k in ["阅读"]):
        return "#FCE4EC", "#D81B60"  # 浅粉
    if any(k in sec_name for k in ["完形"]):
        return "#FFFDE7", "#FBC02D"  # 浅黄
    if any(k in sec_name for k in ["六选五", "七选五", "五选五", "5选5", "6选5", "7选5"]):
        return "#E8EAF6", "#3F51B5"  # 浅靛
    if any(k in sec_name for k in ["作文", "书面表达", "写作"]):
        return "#FFF8E1", "#FFB300"  # 浅金
    return "#F5F9FF", "#42A5F5"  # 默认


# ============================================================
#  HTML 报告模板
# ============================================================
def build_html(name, scores, total, sections, total_max, cfg, logo_b64, mascot_b64, edited=None,
               class_avg=None):
    """生成报告 HTML。
    edited: 可选 dict，覆盖默认生成的内容。支持字段：
        overall / teacher_msg / teacher_quote / suggestions(list of [topic, advice])
    """
    test_name = cfg["test_name"]
    semester = cfg["semester"]
    subject = cfg["subject"]
    teacher = cfg["teacher"]
    date_str = cfg["date"]
    section_advice = cfg["section_advice"]
    edited = edited or {}

    overall = edited.get("overall") or gen_overall_eval(name, total, total_max, scores, sections, test_name)
    badges = gen_badges(scores, sections)
    suggestions = edited.get("suggestions") or gen_suggestions(scores, sections, section_advice)
    default_msg = gen_teacher_msg(name, total, total_max, scores=scores,
                                  sections=sections, test_name=test_name,
                                  english_quote=True, class_avg=class_avg)
    t_msg = edited.get("teacher_msg") or default_msg[0]
    t_quote = edited.get("teacher_quote") or default_msg[1]

    # 分项评分行（每个板块按主题色区分）
    score_rows = ""
    for sec_name, max_score in sections:
        s = scores.get(sec_name, 0)
        stars = star_rating(s, max_score)
        bg_color, border_color = get_section_color(sec_name)
        score_rows += f"""
        <tr style="background:{bg_color};">
            <td class="sec-name" style="border-left:4px solid {border_color};">{sec_name}</td>
            <td class="sec-score">{s:g}<span class="max">/{max_score:g}</span></td>
            <td class="sec-rate">{(s/max_score*100) if max_score > 0 else 0:.0f}%</td>
            <td class="sec-star">{stars}</td>
        </tr>"""

    badge_html = "".join(
        f'<div class="badge"><span class="badge-emoji">{e}</span><span class="badge-text">{l}</span></div>'
        for e, l in badges
    )
    sug_html = "".join(f'<li><span class="sug-topic">{t}：</span>{a}</li>' for t, a in suggestions)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:"PingFang SC","Microsoft YaHei","Helvetica Neue",sans-serif; background:linear-gradient(135deg,#E3F2FD 0%,#E1F5FE 50%,#E8EAF6 100%); padding:30px; width:820px; }}
.report {{ background:#FFF; border-radius:24px; overflow:hidden; box-shadow:0 8px 32px rgba(33,150,243,0.15); position:relative; }}
.header {{ background:linear-gradient(135deg,#42A5F5 0%,#1E88E5 50%,#1565C0 100%); padding:28px 36px 24px; color:white; position:relative; overflow:hidden; }}
.header::before {{ content:"✨"; position:absolute; top:10px; right:20px; font-size:28px; opacity:0.5; }}
.header::after {{ content:"✨"; position:absolute; bottom:8px; right:60px; font-size:18px; opacity:0.4; }}
.brand-row {{ display:flex; align-items:center; gap:10px; margin-bottom:14px; }}
.brand-logo {{ height:40px; }}
.brand-name {{ font-size:14px; font-weight:600; letter-spacing:1px; opacity:0.95; }}
.title-row {{ display:flex; align-items:center; gap:12px; }}
.title-icon {{ font-size:32px; }}
.title {{ font-size:26px; font-weight:700; }}
.subtitle {{ font-size:14px; opacity:0.9; margin-top:4px; }}
.body {{ padding:28px 36px 28px; }}
.section {{ margin-bottom:22px; }}
.sec-title {{ font-size:17px; font-weight:700; color:#1565C0; margin-bottom:12px; display:flex; align-items:center; gap:8px; }}
.sec-title-icon {{ font-size:20px; }}
.eval-box {{ background:#E1F5FE; border-left:4px solid #42A5F5; border-radius:12px; padding:16px 20px; font-size:14px; line-height:1.8; color:#0D47A1; }}
.score-table {{ width:100%; border-collapse:collapse; border-radius:12px; overflow:hidden; box-shadow:0 0 0 1px #BBDEFB; }}
.score-table th {{ background:#BBDEFB; color:#1565C0; font-size:13px; padding:10px 16px; text-align:left; }}
.score-table td {{ padding:10px 16px; font-size:15px; border-bottom:1px solid #E3F2FD; }}
.sec-name {{ color:#1565C0; font-weight:600; }}
.sec-score {{ font-weight:700; color:#0D47A1; font-size:18px; }}
.sec-score .max {{ font-size:13px; color:#9E9E9E; font-weight:400; }}
.sec-rate {{ color:#42A5F5; font-size:14px; }}
.sec-star {{ color:#FFA726; letter-spacing:2px; }}
.score-table tr:last-child td {{ border-bottom:none; }}
.total-row td {{ background:#E1F5FE; font-weight:700; border-top:2px solid #42A5F5; }}
.badges {{ display:flex; flex-wrap:wrap; gap:12px; }}
.badge {{ background:linear-gradient(135deg,#FFF3E0,#FFE0B2); border:1.5px solid #FFB74D; border-radius:20px; padding:8px 18px; display:flex; align-items:center; gap:6px; box-shadow:0 2px 8px rgba(255,167,38,0.2); }}
.badge-emoji {{ font-size:18px; }}
.badge-text {{ font-size:14px; font-weight:600; color:#E65100; }}
.sug-list {{ list-style:none; counter-reset:sug; }}
.sug-list li {{ counter-increment:sug; background:#E1F5FE; border-radius:10px; padding:12px 16px 12px 48px; margin-bottom:10px; font-size:14px; line-height:1.7; color:#0D47A1; position:relative; }}
.sug-list li::before {{ content:counter(sug); position:absolute; left:14px; top:50%; transform:translateY(-50%); width:24px; height:24px; background:#42A5F5; color:white; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:700; }}
.sug-topic {{ color:#1565C0; font-weight:700; }}
.msg-box {{ background:linear-gradient(135deg,#E3F2FD,#BBDEFB); border-radius:16px; padding:20px 24px; }}
.msg-text {{ font-size:14px; line-height:1.8; color:#1565C0; margin-bottom:10px; }}
.msg-quote {{ font-size:15px; font-style:italic; color:#0D47A1; font-weight:600; text-align:right; }}
.footer {{ background:#E1F5FE; padding:16px 36px; display:flex; justify-content:space-between; align-items:center; border-top:1px solid #BBDEFB; }}
.footer-info {{ font-size:13px; color:#1565C0; }}
.footer-mascot {{ height:60px; }}
.heart {{ position:absolute; color:#90CAF9; opacity:0.3; font-size:16px; }}
.heart-1 {{ top:200px; left:12px; }}
.heart-2 {{ bottom:100px; right:12px; font-size:20px; }}
</style></head><body>
<div class="report">
    <div class="header">
        <div class="brand-row">
            <img class="brand-logo" src="data:image/png;base64,{logo_b64}" alt="logo">
            <span class="brand-name">飞扬精准自学 · AFTER-SCHOOL CARE CENTER</span>
        </div>
        <div class="title-row">
            <span class="title-icon">📖</span>
            <div>
                <div class="title">{name}同学{semester}{subject}学习反馈</div>
                <div class="subtitle">{semester} · {test_name} · 满分{total_max:g}分</div>
            </div>
        </div>
    </div>
    <div class="body">
        <span class="heart heart-1">💙</span>
        <span class="heart heart-2">💙</span>
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">📋</span>综合评价</div>
            <div class="eval-box">{overall}</div>
        </div>
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">📊</span>分项评分</div>
            <table class="score-table">
                <tr><th>板块</th><th>得分</th><th>得分率</th><th>星级</th></tr>
                {score_rows}
                <tr class="total-row">
                    <td class="sec-name" style="font-size:16px;">总分</td>
                    <td class="sec-score" style="font-size:20px;">{total:g}<span class="max">/{total_max:g}</span></td>
                    <td class="sec-rate">{(total/total_max*100) if total_max > 0 else 0:.0f}%</td>
                    <td class="sec-star">{star_rating(total, total_max)}</td>
                </tr>
            </table>
        </div>
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">🏆</span>学习亮点</div>
            <div class="badges">{badge_html}</div>
        </div>
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">💡</span>成长建议</div>
            <ol class="sug-list">{sug_html}</ol>
        </div>
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">💌</span>教师寄语</div>
            <div class="msg-box">
                <div class="msg-text">{t_msg}</div>
                <div class="msg-quote">{t_quote}</div>
            </div>
        </div>
    </div>
    <div class="footer">
        <div class="footer-info">教师：{teacher}　|　{date_str}</div>
        <img class="footer-mascot" src="data:image/png;base64,{mascot_b64}" alt="mascot">
    </div>
</div>
</body></html>"""
    return html


# ============================================================
#  报告生成
# ============================================================
def generate_reports(students, sections, cfg, logo_b64, mascot_b64, output_dir, edited_map=None,
                     progress=None):
    """用 Playwright 逐个生成报告图片，返回文件名列表。
    edited_map: 可选 {学生姓名: {overall/teacher_msg/teacher_quote/suggestions}}，覆盖默认内容。
    progress: 可选回调 (已完成数, 总数, 当前学生姓名)，用于前端轮询进度。
    """
    total_max = sum(m for _, m in sections)
    class_avg = sum(t for _, _, t in students) / len(students) if students else None
    edited_map = edited_map or {}
    generated = []
    total = len(students)
    # 每生成 8 份重启一次浏览器，释放 Chromium 累积的内存（云服务器 512MB 也能跑完整班）
    BROWSER_CHUNK = 8

    with sync_playwright() as p:
        browser = None
        page = None
        try:
            for idx, (name, scores, total_score) in enumerate(students):
                if browser is None or idx % BROWSER_CHUNK == 0:
                    if browser is not None:
                        try:
                            page.close()
                            browser.close()
                        except Exception:
                            pass
                        page = None
                        browser = None
                        gc.collect()
                    browser = p.chromium.launch(
                        args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage",
                              "--disable-gpu", "--single-process", "--disable-software-rasterizer"]
                    )
                    page = browser.new_page(viewport={"width": 880, "height": 1200})

                edited = edited_map.get(name, {})
                html = build_html(name, scores, total_score, sections, total_max, cfg, logo_b64, mascot_b64,
                                  edited=edited, class_avg=class_avg)

                # 直接设置页面内容（避免中文路径问题）
                page.set_content(html, wait_until="load")
                page.wait_for_timeout(600)

                # 生成 PNG
                img_name = f"{name}_反馈报告.png"
                img_path = os.path.join(output_dir, img_name)
                el = page.query_selector(".report")
                el.screenshot(path=img_path)
                el = None

                # 生成 PDF（从 PNG 转换，保证排版一致）
                pdf_name = f"{name}_反馈报告.pdf"
                pdf_path = os.path.join(output_dir, pdf_name)
                try:
                    with Image.open(img_path) as im:
                        if im.mode == "RGBA":
                            im = im.convert("RGB")
                        im.save(pdf_path, "PDF", resolution=150.0)
                except Exception as e:
                    print(f"  ⚠ PDF 生成失败 {name}: {e}")

                generated.append((name, total_score, total_max, img_name, pdf_name))
                print(f"  ✓ {name}  {total_score:g}/{total_max:g}")
                if progress:
                    progress(idx + 1, total, name)
        finally:
            try:
                if page is not None:
                    page.close()
                if browser is not None:
                    browser.close()
            except Exception:
                pass

    return generated


# ============================================================
#  Flask 路由
# ============================================================
@app.route("/")
def index():
    default_advice = load_default_advice()
    return render_template("index.html", default_advice=default_advice)


@app.route("/api/analyze", methods=["POST"])
def analyze():
    """AJAX: 上传成绩表文件（Excel/CSV），返回解析出的板块信息"""
    if "excel" not in request.files:
        return jsonify({"error": "请选择成绩表文件（Excel 或 CSV）"}), 400

    f = request.files["excel"]
    if not f.filename:
        return jsonify({"error": "请选择成绩表文件（Excel 或 CSV）"}), 400

    # 根据扩展名保存
    ext = os.path.splitext(f.filename)[1].lower() or ".xlsx"
    if ext not in (".xlsx", ".xls", ".csv", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"):
        ext = ".xlsx"
    tmp_path = os.path.join(UPLOAD_DIR, f"upload_{datetime.now().strftime('%H%M%S')}{ext}")
    f.save(tmp_path)

    try:
        students, sections = load_excel(tmp_path)
        default_advice = load_default_advice()

        section_info = []
        for sec_name, max_score in sections:
            advice = default_advice.get(sec_name, DEFAULT_ADVICE)
            section_info.append({
                "name": sec_name,
                "max": max_score,
                "topic": advice["topic"],
                "advice": advice["advice"],
            })

        total_max = sum(m for _, m in sections)
        return jsonify({
            "ok": True,
            "file_id": os.path.basename(tmp_path),
            "sections": section_info,
            "total_max": total_max,
            "student_count": len(students),
            "students": [{"name": n, "total": t} for n, _, t in students],
        })
    except Exception as e:
        return jsonify({"error": f"解析失败: {str(e)}"}), 500


@app.route("/api/analyze_pdf", methods=["POST"])
def analyze_pdf():
    """AJAX: 上传试卷 PDF，返回每个板块的考点分析"""
    if "pdf" not in request.files:
        return jsonify({"error": "请选择 PDF 文件"}), 400

    f = request.files["pdf"]
    if not f.filename:
        return jsonify({"error": "请选择 PDF 文件"}), 400

    # 获取板块名列表
    section_names_raw = request.form.get("section_names", "")
    try:
        section_names = json.loads(section_names_raw) if section_names_raw else []
    except json.JSONDecodeError:
        section_names = []

    # 保存 PDF 到临时文件
    tmp_path = os.path.join(UPLOAD_DIR, f"pdf_{datetime.now().strftime('%H%M%S')}.pdf")
    f.save(tmp_path)

    try:
        text = extract_pdf_text(tmp_path)
        if not text.strip():
            return jsonify({"error": "PDF 内容为空，可能是扫描件，需 OCR 处理"}), 400

        advice_map = analyze_pdf_content(text, section_names)
        return jsonify({
            "ok": True,
            "advice_map": advice_map,
            "preview": text[:500],
        })
    except Exception as e:
        return jsonify({"error": f"PDF 分析失败: {str(e)}"}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.route("/preview", methods=["POST"])
def preview():
    """预览编辑：生成所有学生报告的文本内容，供用户编辑后再生成"""
    # 获取表单数据
    file_id = request.form.get("file_id", "")
    excel_path = os.path.join(UPLOAD_DIR, file_id)

    # 如果没有 file_id（直接上传模式），从 request.files 获取
    if not file_id or not os.path.isfile(excel_path):
        if "excel" in request.files:
            f = request.files["excel"]
            ext = os.path.splitext(f.filename)[1].lower() or ".xlsx"
            if ext not in (".xlsx", ".xls", ".csv", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"):
                ext = ".xlsx"
            excel_path = os.path.join(UPLOAD_DIR, f"direct_{datetime.now().strftime('%H%M%S')}{ext}")
            f.save(excel_path)
        else:
            return render_template("index.html", error="请先上传成绩表文件（Excel 或 CSV）", default_advice=load_default_advice())

    # 基本配置
    date_val = request.form.get("date", "auto")
    if date_val == "auto" or not date_val:
        date_val = datetime.now().strftime("%Y年%m月%d日")

    cfg = {
        "class_name": request.form.get("class_name", ""),
        "semester": request.form.get("semester", ""),
        "subject": request.form.get("subject", "英语"),
        "test_name": request.form.get("test_name", ""),
        "teacher": request.form.get("teacher", ""),
        "date": date_val,
        "section_advice": {},
    }

    # 板块建议（从表单 JSON）
    advice_json = request.form.get("section_advice", "")
    if advice_json:
        try:
            cfg["section_advice"] = json.loads(advice_json)
        except json.JSONDecodeError:
            cfg["section_advice"] = load_default_advice()
    else:
        cfg["section_advice"] = load_default_advice()

    # 解析 Excel
    try:
        students, sections = load_excel(excel_path)
    except Exception as e:
        return render_template("index.html", error=f"Excel 解析失败: {str(e)}", default_advice=load_default_advice())

    if not students:
        return render_template("index.html", error="Excel 中没有找到学生数据", default_advice=load_default_advice())

    # 生成每位学生的文本内容（不渲染图片，只生成文字）
    total_max = sum(m for _, m in sections)
    class_avg = sum(t for _, _, t in students) / len(students) if students else None
    student_contents = []
    for name, scores, total in students:
        overall = gen_overall_eval(name, total, total_max, scores, sections, cfg["test_name"])
        badges = gen_badges(scores, sections)
        suggestions = gen_suggestions(scores, sections, cfg["section_advice"])
        t_msg, t_quote = gen_teacher_msg(name, total, total_max, scores=scores,
                                         sections=sections, test_name=cfg["test_name"],
                                         english_quote=True, class_avg=class_avg)

        score_detail = []
        for sec_name, max_score in sections:
            s = scores.get(sec_name, 0)
            score_detail.append({
                "name": sec_name,
                "score": s,
                "max": max_score,
                "rate": round(s / max_score * 100, 1) if max_score > 0 else 0,
            })

        student_contents.append({
            "name": name,
            "total": total,
            "total_max": total_max,
            "rate": round(total / total_max * 100, 1) if total_max > 0 else 0,
            "overall": overall,
            "badges": [f"{e} {l}" for e, l in badges],
            "suggestions": [{"topic": t, "advice": a} for t, a in suggestions],
            "teacher_msg": t_msg,
            "teacher_quote": t_quote,
            "scores": score_detail,
        })

    # 把原始表单数据传给 preview.html，确认生成时一起转发到 /generate
    form_data = {
        "file_id": os.path.basename(excel_path),
        "class_name": cfg["class_name"],
        "semester": cfg["semester"],
        "subject": cfg["subject"],
        "test_name": cfg["test_name"],
        "teacher": cfg["teacher"],
        "date": request.form.get("date", "auto"),
        "section_advice": advice_json,
        "app_version": APP_VERSION,
    }

    # 处理上传的 logo/mascot（保存为临时文件，确认生成时再读取）
    logo_path = ""
    mascot_path = ""
    if "logo" in request.files and request.files["logo"].filename:
        logo_path = os.path.join(UPLOAD_DIR, f"logo_preview_{datetime.now().strftime('%H%M%S')}.png")
        request.files["logo"].save(logo_path)
    if "mascot" in request.files and request.files["mascot"].filename:
        mascot_path = os.path.join(UPLOAD_DIR, f"mascot_preview_{datetime.now().strftime('%H%M%S')}.png")
        request.files["mascot"].save(mascot_path)

    print(f"预览：{len(student_contents)} 位学生，等待用户编辑后生成")
    return render_template("preview.html",
                          students=student_contents,
                          cfg=cfg,
                          form_data=form_data,
                          total_max=total_max,
                          logo_path=logo_path,
                          mascot_path=mascot_path)


@app.route("/generate", methods=["POST"])
def generate():
    """生成全部报告"""
    # 获取表单数据
    file_id = request.form.get("file_id", "")
    excel_path = os.path.join(UPLOAD_DIR, file_id)

    # 如果没有 file_id（直接上传模式），从 request.files 获取
    if not file_id or not os.path.isfile(excel_path):
        if "excel" in request.files:
            f = request.files["excel"]
            ext = os.path.splitext(f.filename)[1].lower() or ".xlsx"
            if ext not in (".xlsx", ".xls", ".csv", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"):
                ext = ".xlsx"
            excel_path = os.path.join(UPLOAD_DIR, f"direct_{datetime.now().strftime('%H%M%S')}{ext}")
            f.save(excel_path)
        else:
            return render_template("index.html", error="请先上传成绩表文件（Excel 或 CSV）", default_advice=load_default_advice())

    # 基本配置
    date_val = request.form.get("date", "auto")
    if date_val == "auto" or not date_val:
        date_val = datetime.now().strftime("%Y年%m月%d日")

    cfg = {
        "class_name": request.form.get("class_name", ""),
        "semester": request.form.get("semester", ""),
        "subject": request.form.get("subject", "英语"),
        "test_name": request.form.get("test_name", ""),
        "teacher": request.form.get("teacher", ""),
        "date": date_val,
        "section_advice": {},
    }

    # 板块建议（从表单 JSON）
    advice_json = request.form.get("section_advice", "")
    if advice_json:
        try:
            cfg["section_advice"] = json.loads(advice_json)
        except json.JSONDecodeError:
            cfg["section_advice"] = load_default_advice()
    else:
        cfg["section_advice"] = load_default_advice()

    # Logo 和 Mascot
    logo_b64 = None
    mascot_b64 = None

    # 1) 优先从直接上传的文件读取
    if "logo" in request.files and request.files["logo"].filename:
        logo_tmp = os.path.join(UPLOAD_DIR, "logo_tmp.png")
        request.files["logo"].save(logo_tmp)
        logo_b64 = remove_white_bg(logo_tmp)
        os.remove(logo_tmp)

    if "mascot" in request.files and request.files["mascot"].filename:
        mascot_tmp = os.path.join(UPLOAD_DIR, "mascot_tmp.png")
        request.files["mascot"].save(mascot_tmp)
        mascot_b64 = img_to_b64(mascot_tmp)
        os.remove(mascot_tmp)

    # 2) 否则用预览页保存的临时文件
    if not logo_b64:
        logo_path_form = request.form.get("logo_path", "")
        if logo_path_form and os.path.exists(logo_path_form):
            logo_b64 = remove_white_bg(logo_path_form)
            os.remove(logo_path_form)

    if not mascot_b64:
        mascot_path_form = request.form.get("mascot_path", "")
        if mascot_path_form and os.path.exists(mascot_path_form):
            mascot_b64 = img_to_b64(mascot_path_form)
            os.remove(mascot_path_form)

    # 3) 使用默认图片
    if not logo_b64:
        default_logo = os.path.join(BASE_DIR, "2.png")
        if os.path.exists(default_logo):
            logo_b64 = remove_white_bg(default_logo)

    if not mascot_b64:
        default_mascot = os.path.join(BASE_DIR, "3.PNG")
        if os.path.exists(default_mascot):
            mascot_b64 = img_to_b64(default_mascot)

    # 解析 Excel
    try:
        students, sections = load_excel(excel_path)
    except Exception as e:
        return render_template("index.html", error=f"Excel 解析失败: {str(e)}", default_advice=load_default_advice())

    if not students:
        return render_template("index.html", error="Excel 中没有找到学生数据", default_advice=load_default_advice())

    # 用户编辑后的内容（来自预览页）
    edited_map = {}
    edited_json = request.form.get("edited_content", "")
    # 预览页版本校验：旧预览页里的文本是旧逻辑生成的，提交会覆盖新分析结果
    preview_version = request.form.get("app_version", "")
    if edited_json and preview_version != APP_VERSION:
        return render_template(
            "index.html",
            error="你打开的是旧版本预览页，里面的考点和寄语都是旧的。请重新上传成绩表和试卷 PDF，生成新的预览后再导出。",
            default_advice=load_default_advice())
    if edited_json:
        try:
            edited_map = json.loads(edited_json)
        except json.JSONDecodeError:
            pass

    # 生成报告
    total_max = sum(m for _, m in sections)
    print(f"开始生成 {len(students)} 份报告...")

    # 每次生成使用独立目录，避免不同老师看到彼此的学生报告
    run_id, run_dir = new_run_dir()

    # 后台线程生成 + 前端轮询进度：人数多时请求不会超时，也不会卡住页面
    job = {
        "status": "running",
        "done": 0,
        "total": len(students),
        "current": students[0][0] if students else "",
        "reports": [],
        "total_max": total_max,
        "cfg": cfg,
        "error": "",
    }
    with JOBS_LOCK:
        JOBS[run_id] = job

    def _progress(done, total, name):
        job["done"] = done
        job["current"] = name

    def _run():
        try:
            with GENERATION_LOCK:
                generated = generate_reports(students, sections, cfg, logo_b64, mascot_b64, run_dir,
                                            edited_map=edited_map, progress=_progress)
            job["reports"] = generated
            with open(os.path.join(run_dir, "job.json"), "w", encoding="utf-8") as f:
                json.dump({"reports": generated, "total_max": total_max, "cfg": cfg}, f, ensure_ascii=False)
            job["status"] = "done"
            print(f"完成！{len(generated)} 份报告已保存到 {run_dir}")
        except Exception as e:
            job["status"] = "error"
            job["error"] = str(e)
            print(f"生成失败 {run_id}: {e}")

    threading.Thread(target=_run, daemon=True).start()

    return render_template("generating.html", run_id=run_id, total=len(students))


@app.route("/progress/<run_id>")
def progress(run_id):
    """轮询生成进度"""
    with JOBS_LOCK:
        job = JOBS.get(run_id)
        if not job:
            return jsonify({"status": "error", "error": "任务不存在或服务已重启，请重新生成"}), 404
        return jsonify({
            "status": job["status"],
            "done": job["done"],
            "total": job["total"],
            "current": job["current"],
            "error": job["error"],
        })


@app.route("/results/<run_id>")
def results(run_id):
    """生成完成后展示报告列表"""
    with JOBS_LOCK:
        job = JOBS.get(run_id)
        data = None
        if job and job["status"] == "done":
            data = {"reports": job["reports"], "total_max": job["total_max"], "cfg": job["cfg"]}
    if data is None:
        job_file = os.path.join(OUTPUT_DIR, run_id, "job.json")
        if os.path.isfile(job_file):
            with open(job_file, encoding="utf-8") as f:
                data = json.load(f)
    if not data:
        return "报告还没有生成完成，或链接已过期，请重新生成", 404
    return render_template("results.html", reports=data["reports"], total_max=data["total_max"],
                           cfg=data["cfg"], run_id=run_id)


def _safe_report_path(run_id, filename):
    """只允许访问当前批次目录下的报告文件"""
    if not re.fullmatch(r'[0-9A-Za-z_]+', run_id or ''):
        return None
    filename = os.path.basename(filename or '')
    path = os.path.join(OUTPUT_DIR, run_id, filename)
    return path if os.path.isfile(path) else None


@app.route("/reports/<run_id>/<path:filename>")
def serve_report(run_id, filename):
    path = _safe_report_path(run_id, filename)
    if path:
        return send_file(path, mimetype="image/png")
    return "文件不存在或已过期，请重新生成", 404


@app.route("/reports_pdf/<run_id>/<path:filename>")
def serve_report_pdf(run_id, filename):
    path = _safe_report_path(run_id, filename)
    if path:
        return send_file(path, mimetype="application/pdf")
    return "文件不存在或已过期，请重新生成", 404


@app.route("/download_all")
def download_all():
    """打包下载当前批次的报告（PNG + PDF）"""
    run_id = request.args.get("run", "")
    if not re.fullmatch(r'[0-9A-Za-z_]+', run_id or ''):
        return "下载链接已失效，请重新生成", 404
    run_dir = os.path.join(OUTPUT_DIR, run_id)
    if not os.path.isdir(run_dir):
        return "下载链接已失效，请重新生成", 404

    zip_path = os.path.join(run_dir, "_all_reports.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(os.listdir(run_dir)):
            if fname.endswith((".png", ".pdf")) and not fname.startswith("_"):
                zf.write(os.path.join(run_dir, fname), fname)
    return send_file(zip_path, as_attachment=True, download_name="反馈报告_全部.zip")


if __name__ == "__main__":
    print("=" * 50)
    print("  测评反馈报告生成器")
    port = int(os.environ.get("PORT", 8080))
    print(f"  访问: http://localhost:{port}")
    print("=" * 50)
    app.run(debug=False, host="0.0.0.0", port=port)
