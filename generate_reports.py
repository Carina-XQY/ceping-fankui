#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""为FV2班13位学生生成个性化测评反馈报告（图片格式）"""

import base64
import os
import io
from PIL import Image
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "反馈报告")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def remove_white_bg(path):
    """去除图片白色背景，返回透明背景的 base64"""
    img = Image.open(path).convert("RGBA")
    data = img.getdata()
    new_data = []
    for item in data:
        r, g, b, a = item
        # 接近白色的像素设为透明（含浅灰虚线边框）
        if r > 235 and g > 235 and b > 235:
            new_data.append((r, g, b, 0))
        else:
            new_data.append((r, g, b, a))
    img.putdata(new_data)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

# ---------- 学生数据 ----------
STUDENTS = [
    ("蔡宜熔", 27, 24, 5, 4, 7),
    ("陈诺颖", 29, 36, 9, 6, 11.5),
    ("宫楚儿", 30, 38, 7, 5, 16.5),
    ("李娢静", 30, 34, 5, 1, 11),
    ("李振洋", 27, 40, 6, 4, 15.5),
    ("马嘉阳", 8, 15, 9, 2, 5),
    ("孙苑津", 17, 26, 4, 3, 2.5),
    ("唐晟宇", 29, 28, 2, 0, 12),
    ("田嘉炜", 21, 26, 2, 2, 1),
    ("张智轩", 22, 17, 5, 2, 2),
    ("周厚烨", 27, 34, 10, 1, 9),
    ("周睿杰", 28, 38, 4, 3, 8.5),
    ("周芯宇", 28, 28, 7, 1, 0),
]

MAX_SCORES = {"单词": 30, "短语": 40, "单选": 15, "填空": 15, "翻译": 20}
TOTAL_MAX = 120

# ---------- 图片转base64 ----------
def img_to_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

LOGO_B64 = remove_white_bg(os.path.join(BASE_DIR, "2.png"))
MASCOT_B64 = img_to_b64(os.path.join(BASE_DIR, "3.PNG"))

# ---------- 个性化内容生成 ----------
def get_overall_eval(total, scores):
    """综合评价"""
    rate = total / TOTAL_MAX
    word_rate = scores["单词"] / MAX_SCORES["单词"]
    if rate >= 0.75:
        return (f"本次Unit 1-4小测取得 <b>{total}</b> 分，整体表现优异。"
                f"词汇短语基础扎实，说明平时有认真记背，值得表扬。"
                f"{'翻译板块表现突出，句子组织能力较强。' if scores['翻译']/MAX_SCORES['翻译']>=0.7 else ''}"
                f"语法运用板块仍有提升空间，补上来分数还能再上一个台阶。")
    elif rate >= 0.65:
        return (f"本次Unit 1-4小测取得 <b>{total}</b> 分，整体表现良好。"
                f"{'词汇基础扎实，短语搭配记得牢。' if word_rate>=0.85 else '词汇短语基础尚可，个别不太熟的需要巩固。'}"
                f"语法运用板块失分较多，尤其是非谓语动词和被动语态，这是全班的共性弱点，咱们慢慢补。"
                f"孩子底子是有的，语法补上来分数还能再上一个台阶。")
    elif rate >= 0.55:
        return (f"本次Unit 1-4小测取得 <b>{total}</b> 分，整体表现中等。"
                f"{'词汇基础尚可，但短语搭配还需加强。' if word_rate>=0.7 else '词汇基础还需巩固，部分单词记得不够牢。'}"
                f"语法运用板块失分较多，非谓语动词和被动语态分不清用法，翻译复合句结构也不太熟练。"
                f"孩子其实不笨，就是有些基础没跟上，咱们一步步来，肯定能提上来的。")
    else:
        return (f"本次Unit 1-4小测取得 <b>{total}</b> 分，成绩不太理想，主要原因是词汇基础薄弱影响了整体发挥。"
                f"{'单词和短语是当前比较大的短板，好多搭配还没记熟。' if word_rate<0.7 else '词汇短语基础还需加强。'}"
                f"语法运用和翻译板块也因词汇受限而失分较多。"
                f"孩子现在最需要的是先过好词汇关，词汇上来了其他板块也会跟着好起来。别着急，咱们一起帮孩子把基础补上。")


def get_badges(scores):
    """学习亮点"""
    badges = []
    if scores["单词"] / MAX_SCORES["单词"] >= 0.93:
        badges.append(("📚", "词汇达人"))
    if scores["短语"] / MAX_SCORES["短语"] >= 0.9:
        badges.append(("📝", "短语之星"))
    if scores["单选"] / MAX_SCORES["单选"] >= 0.8:
        badges.append(("⚡", "语法先锋"))
    if scores["填空"] / MAX_SCORES["填空"] >= 0.67:
        badges.append(("🔍", "填空能手"))
    if scores["翻译"] / MAX_SCORES["翻译"] >= 0.7:
        badges.append(("✍️", "翻译达人"))
    if scores["单词"] == MAX_SCORES["单词"] or scores["短语"] == MAX_SCORES["短语"]:
        badges.append(("🌟", "满分标兵"))
    if not badges:
        badges.append(("💪", "坚持之星"))
    return badges


def get_suggestions(scores):
    """成长建议"""
    suggestions = []
    sections = [
        ("填空", "非谓语动词+被动语态", "做题时先判断主动/被动，再根据动词前后的结构判断用 doing / to do / done，建议每天练 3 道填空专项。"),
        ("单选", "时态判断", "做题养成「先找时间标志词，再判断时态」的习惯，重点区分现在完成时、过去进行时和过去完成时。"),
        ("翻译", "复合句结构", "下笔前先想清楚是什么句型（宾语从句/定语从句/状语从句），主从句怎么搭配，再按结构写。"),
        ("短语", "短语搭配", "把 Unit 1-4 的短语按「带来/发生/受益/推迟」等归类整理，搭配着记比单独背更牢固。"),
        ("单词", "核心词汇", "每天朗读+默写 15-20 分钟，当天学的词当天过一遍，用「三次记忆法」巩固。"),
    ]
    # 按得分率从低到高排序
    sections_sorted = sorted(sections, key=lambda x: scores[x[0]] / MAX_SCORES[x[0]])
    for _, topic, advice in sections_sorted[:3]:
        suggestions.append((topic, advice))
    return suggestions


def get_teacher_msg(total, name):
    """教师寄语"""
    if total >= 90:
        msgs = [
            f"{name}同学，你的努力大家都看得到，词汇扎实、翻译有条理，继续保持，语法再精进一步就更棒了！",
            f"\"Practice makes perfect!\" 熟能生巧，你的坚持终将换来质的飞跃。",
        ]
    elif total >= 80:
        msgs = [
            f"{name}同学，底子不错，只要语法这块补上来，分数还能再上一个台阶，加油！",
            f"\"Every expert was once a beginner.\" 每位高手都曾是初学者，坚持你就一定能行。",
        ]
    elif total >= 70:
        msgs = [
            f"{name}同学，你其实不笨，就是有些基础没跟上，咱们一步步来，肯定能提上来的！",
            f"\"Slow and steady wins the race.\" 稳扎稳打，终会成功。",
        ]
    elif total >= 60:
        msgs = [
            f"{name}同学，词汇关先过了，其他板块会跟着好起来。多鼓励自己，咱们一起努力！",
            f"\"A journey of a thousand miles begins with a single step.\" 千里之行始于足下。",
        ]
    else:
        msgs = [
            f"{name}同学，现在最重要的是先把词汇基础打牢，每天坚持一点点，你会看到变化的！",
            f"\"Fall seven times, stand up eight.\" 跌倒了就爬起来，坚持就是胜利。",
        ]
    return msgs


def star_rating(score, max_score):
    """星级评分"""
    rate = score / max_score
    full = int(rate * 5)
    half = 1 if (rate * 5 - full) >= 0.5 else 0
    empty = 5 - full - half
    return "★" * full + "☆" * (5 - full)


# ---------- HTML 模板 ----------
def build_html(name, scores, total):
    overall = get_overall_eval(total, scores)
    badges = get_badges(scores)
    suggestions = get_suggestions(scores)
    teacher_msg = get_teacher_msg(total, name)

    # 分项评分行
    score_rows = ""
    for section in ["单词", "短语", "单选", "填空", "翻译"]:
        s = scores[section]
        m = MAX_SCORES[section]
        stars = star_rating(s, m)
        score_rows += f"""
        <tr>
            <td class="sec-name">{section}</td>
            <td class="sec-score">{s}<span class="max">/{m}</span></td>
            <td class="sec-rate">{(s/m*100):.0f}%</td>
            <td class="sec-star">{stars}</td>
        </tr>"""

    # 徽章
    badge_html = ""
    for emoji, label in badges:
        badge_html += f'<div class="badge"><span class="badge-emoji">{emoji}</span><span class="badge-text">{label}</span></div>'

    # 建议列表
    sug_html = ""
    for i, (topic, advice) in enumerate(suggestions, 1):
        sug_html += f'<li><span class="sug-topic">{topic}：</span>{advice}</li>'

    today = "2026年10月9日"

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{
    font-family: "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
    background: linear-gradient(135deg, #E3F2FD 0%, #E1F5FE 50%, #E8EAF6 100%);
    padding: 30px;
    width: 820px;
}}
.report {{
    background: #FFFFFF;
    border-radius: 24px;
    overflow: hidden;
    box-shadow: 0 8px 32px rgba(33, 150, 243, 0.15);
    position: relative;
}}

/* ---- 顶部横幅 ---- */
.header {{
    background: linear-gradient(135deg, #42A5F5 0%, #1E88E5 50%, #1565C0 100%);
    padding: 28px 36px 24px;
    color: white;
    position: relative;
    overflow: hidden;
}}
.header::before {{
    content: "✨";
    position: absolute;
    top: 10px; right: 20px;
    font-size: 28px;
    opacity: 0.5;
}}
.header::after {{
    content: "✨";
    position: absolute;
    bottom: 8px; right: 60px;
    font-size: 18px;
    opacity: 0.4;
}}
.brand-row {{
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 14px;
}}
.brand-logo {{
    height: 40px;
}}
.brand-name {{
    font-size: 14px;
    font-weight: 600;
    letter-spacing: 1px;
    opacity: 0.95;
}}
.title-row {{
    display: flex;
    align-items: center;
    gap: 12px;
}}
.title-icon {{
    font-size: 32px;
}}
.title {{
    font-size: 26px;
    font-weight: 700;
}}
.subtitle {{
    font-size: 14px;
    opacity: 0.9;
    margin-top: 4px;
}}

/* ---- 内容区 ---- */
.body {{ padding: 28px 36px 28px; }}

.section {{ margin-bottom: 22px; }}
.sec-title {{
    font-size: 17px;
    font-weight: 700;
    color: #1565C0;
    margin-bottom: 12px;
    display: flex;
    align-items: center;
    gap: 8px;
}}
.sec-title-icon {{ font-size: 20px; }}

/* 综合评价 */
.eval-box {{
    background: #E1F5FE;
    border-left: 4px solid #42A5F5;
    border-radius: 12px;
    padding: 16px 20px;
    font-size: 14px;
    line-height: 1.8;
    color: #0D47A1;
}}

/* 分项评分 */
.score-table {{
    width: 100%;
    border-collapse: collapse;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 0 0 1px #BBDEFB;
}}
.score-table th {{
    background: #BBDEFB;
    color: #1565C0;
    font-size: 13px;
    padding: 10px 16px;
    text-align: left;
}}
.score-table td {{
    padding: 10px 16px;
    font-size: 15px;
    border-bottom: 1px solid #E3F2FD;
}}
.sec-name {{ color: #1565C0; font-weight: 600; }}
.sec-score {{ font-weight: 700; color: #0D47A1; font-size: 18px; }}
.sec-score .max {{ font-size: 13px; color: #9E9E9E; font-weight: 400; }}
.sec-rate {{ color: #42A5F5; font-size: 14px; }}
.sec-star {{ color: #FFA726; letter-spacing: 2px; }}
.score-table tr:last-child td {{ border-bottom: none; }}
.total-row td {{
    background: #E1F5FE;
    font-weight: 700;
    border-top: 2px solid #42A5F5;
}}

/* 学习亮点 */
.badges {{
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
}}
.badge {{
    background: linear-gradient(135deg, #FFF3E0, #FFE0B2);
    border: 1.5px solid #FFB74D;
    border-radius: 20px;
    padding: 8px 18px;
    display: flex;
    align-items: center;
    gap: 6px;
    box-shadow: 0 2px 8px rgba(255, 167, 38, 0.2);
}}
.badge-emoji {{ font-size: 18px; }}
.badge-text {{ font-size: 14px; font-weight: 600; color: #E65100; }}

/* 成长建议 */
.sug-list {{
    list-style: none;
    counter-reset: sug;
}}
.sug-list li {{
    counter-increment: sug;
    background: #E1F5FE;
    border-radius: 10px;
    padding: 12px 16px 12px 48px;
    margin-bottom: 10px;
    font-size: 14px;
    line-height: 1.7;
    color: #0D47A1;
    position: relative;
}}
.sug-list li::before {{
    content: counter(sug);
    position: absolute;
    left: 14px; top: 50%;
    transform: translateY(-50%);
    width: 24px; height: 24px;
    background: #42A5F5;
    color: white;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 13px;
    font-weight: 700;
}}
.sug-topic {{ color: #1565C0; font-weight: 700; }}

/* 教师寄语 */
.msg-box {{
    background: linear-gradient(135deg, #E3F2FD, #BBDEFB);
    border-radius: 16px;
    padding: 20px 24px;
    position: relative;
}}
.msg-text {{
    font-size: 14px;
    line-height: 1.8;
    color: #1565C0;
    margin-bottom: 10px;
}}
.msg-quote {{
    font-size: 15px;
    font-style: italic;
    color: #0D47A1;
    font-weight: 600;
    text-align: right;
}}

/* ---- 底部 ---- */
.footer {{
    background: #E1F5FE;
    padding: 16px 36px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid #BBDEFB;
}}
.footer-info {{
    font-size: 13px;
    color: #1565C0;
}}
.footer-mascot {{
    height: 60px;
}}

/* 心形装饰 */
.heart {{
    position: absolute;
    color: #90CAF9;
    opacity: 0.3;
    font-size: 16px;
}}
.heart-1 {{ top: 200px; left: 12px; }}
.heart-2 {{ bottom: 100px; right: 12px; font-size: 20px; }}
</style>
</head>
<body>
<div class="report">
    <!-- 顶部横幅 -->
    <div class="header">
        <div class="brand-row">
            <img class="brand-logo" src="data:image/png;base64,{LOGO_B64}" alt="logo">
            <span class="brand-name">飞扬精准自学 · AFTER-SCHOOL CARE CENTER</span>
        </div>
        <div class="title-row">
            <span class="title-icon">📖</span>
            <div>
                <div class="title">{name}同学秋季学期英语学习反馈</div>
                <div class="subtitle">初三秋季 · Lesson 5 Unit 1-4 小测 · 满分{TOTAL_MAX}分</div>
            </div>
        </div>
    </div>

    <!-- 内容区 -->
    <div class="body">
        <span class="heart heart-1">💙</span>
        <span class="heart heart-2">💙</span>

        <!-- 综合评价 -->
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">📋</span>综合评价</div>
            <div class="eval-box">{overall}</div>
        </div>

        <!-- 分项评分 -->
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">📊</span>分项评分</div>
            <table class="score-table">
                <tr><th>板块</th><th>得分</th><th>得分率</th><th>星级</th></tr>
                {score_rows}
                <tr class="total-row">
                    <td class="sec-name" style="font-size:16px;">总分</td>
                    <td class="sec-score" style="font-size:20px;">{total}<span class="max">/{TOTAL_MAX}</span></td>
                    <td class="sec-rate">{(total/TOTAL_MAX*100):.0f}%</td>
                    <td class="sec-star">{star_rating(total, TOTAL_MAX)}</td>
                </tr>
            </table>
        </div>

        <!-- 学习亮点 -->
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">🏆</span>学习亮点</div>
            <div class="badges">{badge_html}</div>
        </div>

        <!-- 成长建议 -->
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">💡</span>成长建议</div>
            <ol class="sug-list">
                {sug_html}
            </ol>
        </div>

        <!-- 教师寄语 -->
        <div class="section">
            <div class="sec-title"><span class="sec-title-icon">💌</span>教师寄语</div>
            <div class="msg-box">
                <div class="msg-text">{teacher_msg[0]}</div>
                <div class="msg-quote">{teacher_msg[1]}</div>
            </div>
        </div>
    </div>

    <!-- 底部 -->
    <div class="footer">
        <div class="footer-info">教师：赵老师　|　{today}</div>
        <img class="footer-mascot" src="data:image/png;base64,{MASCOT_B64}" alt="mascot">
    </div>
</div>
</body>
</html>"""
    return html


# ---------- 生成报告 ----------
def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 880, "height": 1200})

        for name, w, ph, ch, fb, tr in STUDENTS:
            scores = {"单词": w, "短语": ph, "单选": ch, "填空": fb, "翻译": tr}
            total = w + ph + ch + fb + tr

            html = build_html(name, scores, total)
            html_path = os.path.join(OUTPUT_DIR, f"{name}_temp.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)

            page.goto(f"file://{html_path}")
            page.wait_for_timeout(800)

            # 截取 .report 元素
            report_el = page.query_selector(".report")
            img_path = os.path.join(OUTPUT_DIR, f"{name}_反馈报告.png")
            report_el.screenshot(path=img_path)

            # 清理临时HTML
            os.remove(html_path)
            print(f"✓ {name}  {total}/120  → {img_path}")

        browser.close()

    print(f"\n全部完成！共 {len(STUDENTS)} 份报告已保存到：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()
