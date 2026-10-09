#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测评反馈报告生成器 - Web版
=========================
启动: python3 app.py
访问: http://localhost:5000
"""

import base64
import io
import json
import os
import re
import zipfile
from datetime import datetime

import openpyxl
from flask import Flask, render_template, request, jsonify, send_file
from PIL import Image
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max upload
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    """从表头列表解析出板块名和满分，返回 [(name, max), ...] 和姓名列索引、总分列索引
    支持多种满分格式：
      - 单词(30) / 单词（30）  圆括号
      - 单词[30] / 单词【30】  方括号
      - 单词/30  斜杠
      - 单词 30分  后缀分
      - 单词 满分30 / 单词 满分：30  满分前缀
      - 单词（满分30）  满分+数字
    """
    sections = []
    score_indices = []
    name_idx = None
    total_idx = None
    for i, h in enumerate(headers):
        h = str(h) if h else ""
        # 姓名列
        if h in ("姓名", "名字", "学生", "name", "Name"):
            name_idx = i
            continue
        # 总分列（支持「总分」「总分(120)」「合计」「总分（满分）」等）
        h_clean = re.sub(r"\s*[（(].*?[)）]", "", h).strip()
        if h_clean in ("总分", "合计", "Total", "total") or \
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

        if max_score is not None and name:
            sections.append((name, max_score))
            score_indices.append(i)
    return sections, score_indices, name_idx, total_idx


def load_excel(path):
    """返回 (students, sections)
    students = [(姓名, {板块: 得分}, 总分), ...]
    sections  = [(板块名, 满分), ...]
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []

    headers = [str(h) if h else "" for h in rows[0]]
    sections, score_indices, name_idx, total_idx = parse_headers(headers)

    if name_idx is None:
        name_idx = 0

    students = []
    for row in rows[1:]:
        name = row[name_idx] if name_idx < len(row) else None
        if not name:
            continue
        scores = {}
        for sec_idx, col_idx in enumerate(score_indices):
            if col_idx < len(row):
                sec_name = sections[sec_idx][0]
                val = row[col_idx]
                scores[sec_name] = float(val) if val is not None else 0

        if total_idx is not None and total_idx < len(row) and row[total_idx] is not None:
            total = float(row[total_idx])
        else:
            total = sum(scores.values())

        students.append((str(name), scores, total))

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
    """提取 PDF 全文"""
    text = ""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
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


def split_pdf_by_sections(text):
    """按罗马数字/中文数字切分试卷文本，返回 [(板块标题, 板块内容), ...]"""
    pattern = re.compile(
        r'(?:^|\n)\s*([IVXLCDM]{1,5}|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+|[一二三四五六七八九十]+)[\.、．]\s*([^\n]+)'
    )
    matches = list(pattern.finditer(text))
    if not matches:
        return [("全文", text)]
    sections = []
    for i, m in enumerate(matches):
        title = m.group(2).strip()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        sections.append((title, content))
    return sections


def map_to_section_key(title, content=""):
    """把试卷板块标题映射到 Excel 中的板块名"""
    t = title
    if any(k in t for k in ["单项选择", "单选"]):
        return "单选"
    # 语法填空必须在填空之前检测，否则会被"填空"匹配
    if "语法填空" in t:
        return "语法填空"
    # 完形填空必须在填空之前检测，否则会被"填空"匹配
    if any(k in t for k in ["完形"]):
        return "完形"
    if any(k in t for k in ["填空"]):
        return "填空"
    if any(k in t for k in ["翻译"]):
        return "翻译"
    # 六选五/七选五/五选五
    if any(k in t for k in ["六选五", "七选五", "五选五", "5选5", "6选5", "7选5"]):
        return "六选五"
    # 任务型阅读必须在阅读之前检测
    if "任务型" in t and "阅读" in t:
        return "任务型阅读"
    if any(k in t for k in ["阅读"]):
        return "阅读"
    if any(k in t for k in ["作文", "书面表达", "写作"]):
        return "作文"
    if any(k in t for k in ["短语", "词组"]):
        return "短语"
    if any(k in t for k in ["单词", "词汇"]):
        return "单词"
    # 英汉互译：根据内容判断单词还是短语
    if "互译" in t:
        if re.search(r'\b(?:n|v|adj|adv|prep|conj)\.', content):
            return "单词"
        if re.search(r'\b\w+\s+\w+', content):
            return "短语"
    return None


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


def detect_clauses(text):
    """检测从句，返回 [从句类型]"""
    text_l = text.lower()
    detail = []
    # 优先看中文提示词（试卷括号里常标注）
    if "宾语从句" in text:
        detail.append("宾语从句")
    if "定语从句" in text:
        detail.append("定语从句")
    if "状语从句" in text:
        detail.append("状语从句")
    if "同位语" in text:
        detail.append("同位语从句")
    if "主语从句" in text:
        detail.append("主语从句")
    if "表语从句" in text:
        detail.append("表语从句")
    # 没有中文提示时，从结构检测
    if not detail:
        if re.search(r'\b(?:i|we|he|she|they|you)\s+(?:think|believe|hope|wonder|know|said|asked|told|didn\'t know)\b', text_l) or \
           re.search(r'\bdo you (?:think|know|believe)\b', text_l):
            detail.append("宾语从句")
        if re.search(r'\b(?:which|who|whom|whose|that)\s+(?:is|are|was|were|has|have|will|can|do|does|did)\b', text_l) or \
           re.search(r'\bthe (?:\w+\s+)?(?:which|who|whom|that|whose)\b', text_l):
            detail.append("定语从句")
        if re.search(r'\b(?:because|although|though|since|unless|as long as|so that|in order that|so\s+\w+\s+that)\b', text_l):
            detail.append("状语从句")
    return detail


def detect_special_sentences(text):
    """检测特殊句式，返回 [句式名]"""
    text_l = text.lower()
    detail = []
    # 感叹句
    if re.search(r'\bwhat\s+(?:a|an|the)?\s+\w+', text_l) and '!' in text or \
       re.search(r'\bhow\s+\w+\s+!', text_l):
        detail.append("感叹句")
    # 强调句
    if re.search(r'\bit (?:is|was)\s+\w+\s+that\b', text_l):
        detail.append("强调句")
    # 倒装句
    if re.search(r'\b(?:never|hardly|seldom|not only|only by|only when|only after)\b', text_l) and \
       re.search(r'\b(?:have|has|had|did|do|does|is|are|was|were|will|can|could|should)\s+\w+\b', text_l):
        detail.append("倒装句")
    # would rather do A than do B
    if re.search(r'\bwould rather\s+\w+\s+than\s+\w+', text_l):
        detail.append("would rather do A than do B")
    # It takes sb. ... to do
    if re.search(r'\bit takes?\s+\w+\s+\w+\s+to\s+\w+', text_l):
        detail.append("It takes sb. ... to do")
    # too...to...
    if re.search(r'\btoo\s+\w+\s+to\s+\w+', text_l):
        detail.append("too...to... 句型")
    # not...until...
    if re.search(r'\bnot\s+\w+\s+until\b', text_l) or re.search(r"\bwon't\s+\w+\s+until\b", text_l):
        detail.append("not...until... 句型")
    # 主将从现（if条件句）
    if re.search(r'\bif\s+\w+\s+\w+\b.*\b(?:will|won\'t|shall)\b', text_l):
        detail.append("主将从现（if条件句）")
    return detail


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
    """单选板块 - 主要考时态、语态、从句、词法"""
    tenses = detect_tenses(block_text)
    passives = detect_passive_voice(block_text)
    clauses = detect_clauses(block_text)
    specials = detect_special_sentences(block_text)
    modals = detect_modal_verbs(block_text)
    comparisons = detect_comparisons(block_text)

    all_points = []
    if tenses:
        all_points.extend(tenses)
    if passives:
        all_points.append("被动语态")
    if clauses:
        all_points.extend(clauses)
    if specials:
        all_points.extend(specials)
    if modals:
        all_points.append("情态动词")
    if comparisons:
        all_points.extend(comparisons)

    if not all_points:
        topic = "语法综合"
        advice = ("做题时先判断考点（时态/语态/从句/介词），再选最佳答案。"
                  "建议每天练5道单选，做完后核对考点并归纳错题。")
        return {"topic": topic, "advice": advice}

    extras = ""
    if passives:
        extras += "被动语态要特别注意时态搭配。"
    if clauses:
        extras += "从句要判断类型再选引导词。"
    if modals:
        extras += "情态动词注意表推测/表必要的用法区分。"

    if tenses and len(tenses) >= 2:
        topic = "时态综合判断"
        advice = (f"做题养成「先找时间标志词，再判断时态」的习惯。本次重点考{'、'.join(tenses)}。"
                  + extras +
                  "建议每天练5道，做完后归纳错题考点。")
    elif tenses:
        topic = f"{tenses[0]}判断"
        advice = (f"做题先找时间标志词（already/yet/by the time/last year/next week等），"
                  f"再判断时态。本次重点考{'、'.join(tenses)}。" + extras +
                  "建议每天练5道，做完后归纳错题考点。")
    elif passives:
        topic = "被动语态"
        advice = (f"做题先判断主动/被动（主语是动作执行者还是承受者），"
                  f"再选时态搭配（{'、'.join(passives)}）。建议每天练5道，归纳错题考点。")
    elif clauses:
        topic = "从句运用"
        advice = (f"做题先判断从句类型（{'、'.join(clauses)}），再选正确引导词。"
                  + extras + "建议每天练5道，归纳错题考点。")
    elif specials:
        topic = "特殊句式"
        advice = (f"本次涉及{'、'.join(specials)}，掌握句式结构是关键。" + extras +
                  "建议每天练5道，归纳错题考点。")
    elif comparisons:
        topic = "比较等级"
        advice = (f"本次涉及{'、'.join(comparisons)}，注意比较级/最高级/同级比较的搭配。"
                  "建议每天练5道。")
    else:
        topic = "语法综合"
        advice = ("做题时先判断考点（时态/语态/从句/介词），再选最佳答案。"
                  "建议每天练5道，归纳错题考点。")

    return {"topic": topic, "advice": advice}


def analyze_fill_blank_section(block_text):
    """填空板块 - 主要考非谓语+被动"""
    nonfinites = detect_nonfinite(block_text)
    passives = detect_passive_voice(block_text)

    all_points = []
    if nonfinites:
        all_points.extend(nonfinites)
    if passives:
        all_points.append("被动语态")

    if not all_points:
        topic = "动词形式变化"
        advice = ("做题时先判断主动/被动，再根据动词前后结构判断用 doing / to do / done，"
                  "建议每天练3道填空专项。")
        return {"topic": topic, "advice": advice}

    has_passive = "被动语态" in all_points
    has_nonfinite = bool(nonfinites)

    if has_passive and has_nonfinite:
        topic = "非谓语动词+被动语态"
        passive_str = "（" + "、".join(passives) + "）" if passives else ""
        advice = (f"做题时先判断主动/被动，再判断动词形式（doing / to do / done）。"
                  f"本次涉及{'、'.join(nonfinites)}与被动语态{passive_str}。"
                  "被动语态注意时态搭配（be + 过去分词），非谓语注意前后结构。"
                  "建议每天练3道填空专项。")
    elif has_passive:
        topic = "被动语态"
        passive_str = "、".join(passives) if passives else "be + 过去分词"
        advice = (f"做题先判断主动/被动（主语是动作执行者还是承受者），"
                  f"被动结构 be + 过去分词，注意时态搭配（{passive_str}）。"
                  "建议每天练3道填空专项。")
    elif has_nonfinite:
        topic = "非谓语动词"
        advice = (f"做题时根据动词前后结构判断用 doing / to do / done。"
                  f"本次涉及{'、'.join(nonfinites)}。建议每天练3道填空专项，归纳固定搭配。")
    else:
        topic = "动词形式变化"
        advice = ("做题时先判断主动/被动，再根据动词前后结构判断用 doing / to do / done，"
                  "建议每天练3道填空专项。")

    return {"topic": topic, "advice": advice}


def analyze_translation_section(block_text):
    """翻译板块 - 主要考从句和特殊句式"""
    clause_hints = detect_clauses(block_text)
    specials = detect_special_sentences(block_text)

    if clause_hints:
        topic = "复合句结构"
        extras = "同时注意" + "、".join(specials) + "的句式特征。" if specials else ""
        advice = (f"下笔前先想清楚是什么句型（{'、'.join(clause_hints)}），"
                  "主从句怎么搭配，再按结构写。" + extras +
                  "建议先列结构（主句/从句引导词），再填具体内容。")
    elif specials:
        topic = "特殊句式翻译"
        advice = (f"下笔前先想清楚句型（{'、'.join(specials)}），按结构翻译。"
                  "注意中英文语序差异，建议先列结构再填内容。")
    else:
        topic = "句子翻译"
        advice = ("下笔前先想清楚句型结构，注意主谓一致和语序，再按结构写。"
                  "先列主句骨架，再加修饰成分。")

    return {"topic": topic, "advice": advice}


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


def analyze_grammar_fill_section(block_text):
    """语法填空板块 - 定词性 → 变词形
    核心做题逻辑（人教版）：
    ①定词性——空格前后缺什么词性？主语位缺n/谓语位缺v/修饰位缺adj或adv
    ②变词形——根据语境变形（时态/被动/非谓语/单复数/比较级）
    """
    forms = []
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

    topic = "语法填空（定词性变词形）"
    forms_str = "、".join(forms) if forms else "时态/语态/非谓语"
    advice = (f"做题两步走：①定词性——空格前后缺什么？主语位缺名词n，谓语位缺动词v，"
              f"修饰位缺形容词adj或副词adv；②变词形——根据句子语境变形。"
              f"本次重点考点：{forms_str}。"
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
    tenses = detect_tenses(block_text)
    clauses = detect_clauses(block_text)
    passives = detect_passive_voice(block_text)

    all_points = []
    if tenses:
        all_points.extend(tenses)
    if passives:
        all_points.append("被动语态")
    if clauses:
        all_points.extend(clauses)

    if all_points:
        topic = "综合语法"
        advice = f"本板块涉及{'、'.join(all_points[:5])}。建议针对薄弱点专项练习，每天坚持找找感觉。"
    else:
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

    # 2. 把试卷板块映射到 Excel 板块名
    sec_text_map = {}
    for title, content in section_blocks:
        key = map_to_section_key(title, content)
        if key and key not in sec_text_map:
            sec_text_map[key] = content

    # 3. 对每个板块做针对性分析
    for sec_name in section_names:
        block_text = sec_text_map.get(sec_name, text)
        analyzer = SECTION_ANALYZERS.get(sec_name)
        if analyzer:
            result[sec_name] = analyzer(block_text)
        else:
            result[sec_name] = analyze_generic_section(block_text, sec_name)

    return result


# ============================================================
#  个性化内容生成
# ============================================================
def gen_overall_eval(name, total, total_max, scores, sections, test_name):
    rate = total / total_max if total_max > 0 else 0
    rates = {s: scores.get(s, 0) / m for s, m in sections if m > 0}
    strongest = max(rates, key=rates.get) if rates else ""
    weakest = min(rates, key=rates.get) if rates else ""
    word_rate = rates.get("单词", 0)

    if rate >= 0.75:
        text = (f"本次{test_name}取得 <b>{total:g}</b> 分，整体表现优异。"
                f"{strongest}板块掌握扎实，说明平时有认真记背，值得表扬。"
                f"{'翻译板块表现突出，句子组织能力较强。' if rates.get('翻译', 0) >= 0.7 else ''}"
                f"语法运用板块仍有提升空间，补上来分数还能再上一个台阶。")
    elif rate >= 0.65:
        text = (f"本次{test_name}取得 <b>{total:g}</b> 分，整体表现良好。"
                f"{'词汇基础扎实，短语搭配记得牢。' if word_rate >= 0.85 else '词汇短语基础尚可，个别不太熟的需要巩固。'}"
                f"{weakest}板块失分较多，这是需要重点突破的方向。"
                f"孩子底子是有的，薄弱板块补上来分数还能再上一个台阶。")
    elif rate >= 0.55:
        text = (f"本次{test_name}取得 <b>{total:g}</b> 分，整体表现中等。"
                f"{'词汇基础尚可，但短语搭配还需加强。' if word_rate >= 0.7 else '词汇基础还需巩固，部分知识点记得不够牢。'}"
                f"{weakest}板块是主要失分点，需要多做练习找找感觉。"
                f"孩子其实不笨，就是有些基础没跟上，咱们一步步来，肯定能提上来的。")
    else:
        text = (f"本次{test_name}取得 <b>{total:g}</b> 分，成绩不太理想，"
                f"主要原因是{'词汇基础薄弱' if word_rate < 0.7 else '基础不够扎实'}影响了整体发挥。"
                f"{weakest}板块也因基础受限而失分较多。"
                f"孩子现在最需要的是先把基础打牢，基础上来了其他板块也会跟着好起来。"
                f"别着急，咱们一起帮孩子把基础补上。")
    return text


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


def gen_teacher_msg(name, total, total_max):
    rate = total / total_max if total_max > 0 else 0
    if rate >= 0.75:
        return (f"{name}同学，你的努力大家都看得到，继续保持，薄弱板块再精进一步就更棒了！",
                '"Practice makes perfect!" 熟能生巧，你的坚持终将换来质的飞跃。')
    elif rate >= 0.65:
        return (f"{name}同学，底子不错，只要薄弱板块补上来，分数还能再上一个台阶，加油！",
                '"Every expert was once a beginner." 每位高手都曾是初学者，坚持你就一定能行。')
    elif rate >= 0.55:
        return (f"{name}同学，你其实不笨，就是有些基础没跟上，咱们一步步来，肯定能提上来的！",
                '"Slow and steady wins the race." 稳扎稳打，终会成功。')
    elif rate >= 0.45:
        return (f"{name}同学，基础关先过了，其他板块会跟着好起来。多鼓励自己，咱们一起努力！",
                '"A journey of a thousand miles begins with a single step." 千里之行始于足下。')
    else:
        return (f"{name}同学，现在最重要的是先把基础打牢，每天坚持一点点，你会看到变化的！",
                '"Fall seven times, stand up eight." 跌倒了就爬起来，坚持就是胜利。')


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
def build_html(name, scores, total, sections, total_max, cfg, logo_b64, mascot_b64, edited=None):
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
    default_msg = gen_teacher_msg(name, total, total_max)
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
def generate_reports(students, sections, cfg, logo_b64, mascot_b64, output_dir, edited_map=None):
    """用 Playwright 逐个生成报告图片，返回文件名列表。
    edited_map: 可选 {学生姓名: {overall/teacher_msg/teacher_quote/suggestions}}，覆盖默认内容。
    """
    total_max = sum(m for _, m in sections)
    edited_map = edited_map or {}
    generated = []

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 880, "height": 1200})

        for name, scores, total in students:
            edited = edited_map.get(name, {})
            html = build_html(name, scores, total, sections, total_max, cfg, logo_b64, mascot_b64, edited=edited)
            html_path = os.path.join(output_dir, f"_tmp_{name}.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)

            page.goto(f"file://{html_path}")
            page.wait_for_timeout(600)

            img_name = f"{name}_反馈报告.png"
            img_path = os.path.join(output_dir, img_name)
            el = page.query_selector(".report")
            el.screenshot(path=img_path)
            os.remove(html_path)

            generated.append((name, total, total_max, img_name))
            print(f"  ✓ {name}  {total:g}/{total_max:g}")

        browser.close()

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
    """AJAX: 上传 Excel 文件，返回解析出的板块信息"""
    if "excel" not in request.files:
        return jsonify({"error": "请选择 Excel 文件"}), 400

    f = request.files["excel"]
    if not f.filename:
        return jsonify({"error": "请选择 Excel 文件"}), 400

    # 保存到临时文件
    tmp_path = os.path.join(UPLOAD_DIR, f"upload_{datetime.now().strftime('%H%M%S')}.xlsx")
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
    if not os.path.exists(excel_path):
        if "excel" in request.files:
            f = request.files["excel"]
            excel_path = os.path.join(UPLOAD_DIR, f"direct_{datetime.now().strftime('%H%M%S')}.xlsx")
            f.save(excel_path)
        else:
            return render_template("index.html", error="请先上传 Excel 文件", default_advice=load_default_advice())

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
    student_contents = []
    for name, scores, total in students:
        overall = gen_overall_eval(name, total, total_max, scores, sections, cfg["test_name"])
        badges = gen_badges(scores, sections)
        suggestions = gen_suggestions(scores, sections, cfg["section_advice"])
        t_msg, t_quote = gen_teacher_msg(name, total, total_max)

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
    if not os.path.exists(excel_path):
        if "excel" in request.files:
            f = request.files["excel"]
            excel_path = os.path.join(UPLOAD_DIR, f"direct_{datetime.now().strftime('%H%M%S')}.xlsx")
            f.save(excel_path)
        else:
            return render_template("index.html", error="请先上传 Excel 文件", default_advice=load_default_advice())

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
    if edited_json:
        try:
            edited_map = json.loads(edited_json)
        except json.JSONDecodeError:
            pass

    # 生成报告
    total_max = sum(m for _, m in sections)
    print(f"开始生成 {len(students)} 份报告...")

    generated = generate_reports(students, sections, cfg, logo_b64, mascot_b64, OUTPUT_DIR, edited_map=edited_map)

    print(f"完成！{len(generated)} 份报告已保存到 {OUTPUT_DIR}")

    return render_template("results.html", reports=generated, total_max=total_max, cfg=cfg)


@app.route("/reports/<filename>")
def serve_report(filename):
    path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(path):
        return send_file(path, mimetype="image/png")
    return "文件不存在", 404


@app.route("/download_all")
def download_all():
    """打包下载所有报告"""
    zip_path = os.path.join(OUTPUT_DIR, "_all_reports.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(OUTPUT_DIR):
            if fname.endswith(".png") and not fname.startswith("_"):
                zf.write(os.path.join(OUTPUT_DIR, fname), fname)
    return send_file(zip_path, as_attachment=True, download_name="反馈报告_全部.zip")


if __name__ == "__main__":
    print("=" * 50)
    print("  测评反馈报告生成器")
    port = int(os.environ.get("PORT", 8080))
    print(f"  访问: http://localhost:{port}")
    print("=" * 50)
    app.run(debug=False, host="0.0.0.0", port=port)
