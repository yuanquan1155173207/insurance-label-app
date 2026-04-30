import pdfplumber, fitz, re, os, io, math
import pandas as pd
from dataclasses import dataclass
from collections import Counter

def find_chinese_font():
    repo_fonts = [
        "font.otf", "font.ttc", "font.ttf",
        os.path.join(os.path.dirname(__file__), "font.otf"),
        os.path.join(os.path.dirname(__file__), "font.ttc"),
        os.path.join(os.path.dirname(__file__), "font.ttf"),
    ]
    for p in repo_fonts:
        if os.path.exists(p): return p
    candidates = [
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/Library/Fonts/Arial Unicode MS.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    ]
    for p in candidates:
        if os.path.exists(p): return p
    return None

@dataclass
class CriticalIllnessPolicy:
    insured_name:              str   = ""
    insured_age:               int   = 0
    applicant_name:            str   = ""
    currency:                  str   = "港幣"
    annual_premium:            float = 0
    payment_years:             int   = 0
    coverage_age:              int   = 100
    continuous_cancer_monthly: float = 0
    base_sum_insured:          float = 0
    extra_sum_insured:         float = 0
    extra_years:               int   = 10
    extra_ratio:               int   = 50

RED    = (0.85, 0.05, 0.05)
ORANGE = (0.90, 0.45, 0.00)
GREEN  = (0.05, 0.50, 0.05)

def _write(page, text, x, y, color, font_path, fontsize=10):
    kw = dict(fontsize=fontsize, color=color)
    if font_path and os.path.exists(font_path):
        kw["fontfile"] = font_path
        kw["fontname"] = "cjk"
    try:    return page.insert_text((x, y), text, **kw)
    except: return -1

def _write_centered(page, text, y, color, font_path, fontsize=10):
    pw = page.rect.width
    x  = max(10, (pw - len(text) * fontsize * 0.82) / 2)
    return _write(page, text, x, y, color, font_path, fontsize)

def _format_wan(amount, currency=""):
    if not amount or amount <= 0: return "XXX"
    cur = currency or ""
    wan = amount / 10000
    if wan >= 1 and wan == int(wan):   return f"{int(wan)}W{cur}"
    elif wan >= 1:                      return f"{wan:.0f}W{cur}"
    else:                               return f"{int(amount):,}{cur}"

def _draw_red_box(fitz_page, rect, line_width=1.5):
    shape = fitz_page.new_shape()
    shape.draw_rect(rect)
    shape.finish(color=RED, fill=None, width=line_width)
    shape.commit()

def _draw_underline(fitz_page, rect, line_width=1.5):
    shape = fitz_page.new_shape()
    shape.draw_line(fitz.Point(rect.x0, rect.y1), fitz.Point(rect.x1, rect.y1))
    shape.finish(color=RED, width=line_width, closePath=False)
    shape.commit()

def _draw_arrowhead(page, x0, y0, x1, y1, size=6):
    dx = x1-x0; dy = y1-y0
    length = math.hypot(dx, dy)
    if length == 0: return
    ux = dx/length; uy = dy/length
    vx = -uy; vy = ux
    p1  = fitz.Point(x1-size*ux+size*0.5*vx, y1-size*uy+size*0.5*vy)
    p2  = fitz.Point(x1-size*ux-size*0.5*vx, y1-size*uy-size*0.5*vy)
    tip = fitz.Point(x1, y1)
    shape = page.new_shape()
    shape.draw_polyline([p1, tip, p2])
    shape.finish(color=RED, fill=RED, width=1.0, closePath=False)
    shape.commit()

def _group_by_rows(words, tolerance=4):
    if not words: return []
    sorted_words = sorted(words, key=lambda w: w["top"])
    rows = []
    for w in sorted_words:
        placed = False
        for row in rows:
            if abs(row["y"] - w["top"]) <= tolerance:
                row["words"].append(w)
                row["y"] = (row["y"]*(len(row["words"])-1)+w["top"])/len(row["words"])
                placed = True; break
        if not placed:
            rows.append({"y": w["top"], "words": [w]})
    rows.sort(key=lambda r: r["y"])
    return [(r["y"], sorted(r["words"], key=lambda w: w["x0"])) for r in rows]


# ═══════════════════════════════════════════════════════════════
# 页面类型识别
# ═══════════════════════════════════════════════════════════════
def _is_cover_page(full_text):
    return (
        "保障摘要"        in full_text and
        ("投保時每年總保費" in full_text or "每年總保費" in full_text) and
        "基本計劃"        in full_text and
        "說明摘要"    not in full_text
    )

def _is_summary_page(words_text):
    return (
        "說明摘要"     in words_text and
        "補充說明摘要" not in words_text and
        "(1)+(2)"      in words_text and
        "(3)+(4)"      in words_text and
        "100歲"        in words_text and
        "已繳保費"     in words_text
    )

def _is_multi_page(full_text):
    return (
        "多重保險賠償" in full_text and
        "次索償"       in full_text and
        ("9 次索償" in full_text or "9次索償" in full_text) and
        "600%"         in full_text
    )

def _is_cancer_page(full_text):
    return (
        "持續癌症" in full_text and
        ("每月" in full_text or "5%" in full_text) and
        "要求條件" in full_text
    )

def _is_supplement_no_withdrawal(full_text):
    return (
        "補充說明摘要"     in full_text and
        "沒有行使保單選項" in full_text and
        "已繳保費"         in full_text and
        "悲觀"         not in full_text and
        "樂觀"         not in full_text and
        "最高貸款額"   not in full_text and
        "解釋附註"     not in full_text
    )

def _is_supplement_with_withdrawal(full_text):
    return (
        "補充說明摘要"     in full_text and
        "提取款項"         in full_text and
        "退保發還金額"     in full_text and
        "已繳保費"         in full_text and
        "沒有行使保單選項" not in full_text and
        "終期紅利之面值"   not in full_text and
        "悲觀"         not in full_text and
        "樂觀"         not in full_text and
        "最高貸款額"   not in full_text and
        "解釋附註"     not in full_text
    )


# ═══════════════════════════════════════════════════════════════
# 个人信息遮盖（简化坐标策略：每页无条件遮右上角 + 左下角）
# ═══════════════════════════════════════════════════════════════
def redact_personal_info(doc: fitz.Document, is_savings: bool = False) -> fitz.Document:
    """
    改进策略：
    ① 右上角：动态检测正文标题位置作为下边界，宽度收窄到 30%
    ② 左下角：优先基于签名关键词定位，找不到才用兜底
    ③ 保单号文字精确遮盖
    ④ 封面 logo 区域单独处理
    """
    WHITE = (1, 1, 1)

    # 提取保单号
    policy_number = None
    for page in doc:
        text = page.get_text("text")
        m = re.search(r"[A-Z]{2}\d{6}-\d{8,}-\d", text)
        if m:
            policy_number = m.group(0)
            break

    # 正文标题关键词（用于定位右上角遮盖下边界）
    title_keywords = [
        "保障摘要", "說明摘要", "補充說明摘要", "解釋附註",
        "保費徵費摘要", "提取款項之補充摘要",
        "不同投資回報下的說明",
        "愛唯守危疾保障",  # 重疾险封面
        "盛利",            # 储蓄险产品名
    ]

    for page_num in range(len(doc)):
        page = doc[page_num]
        pw   = page.rect.width
        ph   = page.rect.height
        text = page.get_text("text")
        redact_rects = []

        # ═══════════════════════════════════════════════════════
        # 1️⃣ 动态检测：找到正文第一个标题的 y 坐标作为遮盖下界
        # ═══════════════════════════════════════════════════════
        title_top_y = None
        for kw in title_keywords:
            hits = page.search_for(kw)
            if hits:
                # 取所有命中里最靠上的那个
                top_hit = min(hits, key=lambda r: r.y0)
                # 只要标题在页面上半部（y < 25%），才作为下边界参考
                if top_hit.y0 < ph * 0.25:
                    if title_top_y is None or top_hit.y0 < title_top_y:
                        title_top_y = top_hit.y0

        # 右上角遮盖的下边界：
        # - 如果找到标题 → 遮到标题上方 3pt
        # - 找不到       → 保守用 8%（比原来的 10% 更紧）
        if title_top_y is not None:
            top_bottom_y = max(title_top_y - 3, ph * 0.04)
        else:
            top_bottom_y = ph * 0.08

        # ═══════════════════════════════════════════════════════
        # 2️⃣ 右上角遮盖：宽度从 45% 收窄到 30%
        # ═══════════════════════════════════════════════════════
        # 条形码一般在最右侧，30% 宽度足够覆盖
        redact_rects.append(fitz.Rect(pw * 0.70, 0, pw, top_bottom_y))

        # ═══════════════════════════════════════════════════════
        # 3️⃣ 封面页特殊处理（标题上方的 logo 条区域）
        # ═══════════════════════════════════════════════════════
        if is_savings:
            is_savings_cover = ("保障摘要" in text and "保障項目" in text)
            if is_savings_cover and title_top_y is not None and title_top_y > ph * 0.08:
                # 标题上方、logo 右侧（从 40% 宽度开始）
                redact_rects.append(
                    fitz.Rect(pw * 0.40, 0, pw * 0.70, title_top_y - 3)
                )
        else:
            is_ci_cover = _is_cover_page(text)
            if is_ci_cover:
                hits_title = page.search_for("愛唯守危疾保障")
                if hits_title and hits_title[0].y0 > ph * 0.08:
                    redact_rects.append(
                        fitz.Rect(pw * 0.35, 0, pw * 0.70, hits_title[0].y0 - 3)
                    )

        # ═══════════════════════════════════════════════════════
        # 4️⃣ 保单号文字：精确遮盖（只遮文字所在行右半部分）
        # ═══════════════════════════════════════════════════════
        if policy_number:
            for r in page.search_for(policy_number):
                # 只遮盖从保单号起始位置到页面右边
                redact_rects.append(
                    fitz.Rect(r.x0 - 3, r.y0 - 1, pw, r.y1 + 2)
                )

        # ═══════════════════════════════════════════════════════
        # 5️⃣ 左下角签名区：动态检测
        # ═══════════════════════════════════════════════════════
        sig_keywords = [
            "被保人姓名", "申請人姓名", "建議被保人",
            "投保人簽署", "申請人簽署",
            "理財顧問", "保險顧問", "Source Code",
        ]
        bottom_threshold = ph * 0.70
        sig_top_y = None
        for kw in sig_keywords:
            for r in page.search_for(kw):
                if r.y0 >= bottom_threshold:
                    if sig_top_y is None or r.y0 < sig_top_y:
                        sig_top_y = r.y0
                    break

        if sig_top_y is not None:
            # 找到签名关键词 → 从该关键词上方 2pt 遮到页面底部
            redact_rects.append(
                fitz.Rect(0, sig_top_y - 2, pw * 0.55, ph)
            )
        else:
            # 找不到 → 兜底只遮页脚（底部 5%，更保守）
            redact_rects.append(
                fitz.Rect(0, ph * 0.95, pw * 0.55, ph)
            )

        # ═══════════════════════════════════════════════════════
        # 6️⃣ 应用遮盖
        # ═══════════════════════════════════════════════════════
        for rect in redact_rects:
            page.add_redact_annot(rect, fill=WHITE)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

    return doc



# ═══════════════════════════════════════════════════════════════
# 重疾险 1. 保障摘要页（首页）标识 —— 两行合并为一行
# ═══════════════════════════════════════════════════════════════
def _annotate_cover(fitz_page, words, policy, font_path):
    hits_extra = fitz_page.search_for("額外保障")
    hits_base  = fitz_page.search_for("愛唯守危疾保障")

    base_str  = _format_wan(policy.base_sum_insured, policy.currency)
    ratio     = policy.extra_ratio or 50
    ext_years = policy.extra_years or 10

    premium     = int(policy.annual_premium) if policy.annual_premium else 0
    premium_str = f"{premium:,}" if premium > 0 else "XXXX"
    years_str   = str(int(policy.payment_years)) if policy.payment_years else "XX"
    age_str     = str(int(policy.coverage_age))  if policy.coverage_age  else "100"

    # ★ 合并成一行
    line_all = f"保額是{base_str}，首{ext_years}年額外贈送{ratio}%保額；年保費是{premium_str}，交{years_str}年，保到{age_str}歲"

    pw = fitz_page.rect.width

    # ── 基本计划行数字定位 ──
    base_num_rect  = None
    extra_row_line = None
    base_row_line  = None

    base_num_str = f"{int(policy.base_sum_insured):,}" if policy.base_sum_insured else "1,000,000"
    hits_base_num = fitz_page.search_for(base_num_str)
    if hits_base_num:
        r = hits_base_num[0]
        base_num_rect = fitz.Rect(r.x0 - 2, r.y0 - 1, r.x1 + 2, r.y1 + 1)
        base_row_line = fitz.Rect(36, r.y0 - 1, pw - 36, r.y1 + 1)

    extra_num_str = f"{int(policy.extra_sum_insured):,}" if policy.extra_sum_insured else "500,000"
    hits_extra_num = fitz_page.search_for(extra_num_str)
    if hits_extra_num:
        r = hits_extra_num[0]
        extra_row_line = fitz.Rect(36, r.y0 - 1, pw - 36, r.y1 + 1)

    # ── 计算写字 y 坐标 ──
    hits_total = fitz_page.search_for("投保時每年總保費")
    total_y    = hits_total[0].y0 if hits_total else fitz_page.rect.height * 0.85

    if extra_row_line:
        table_bottom = extra_row_line.y1
    elif hits_extra:
        table_bottom = hits_extra[0].y1
    elif hits_base:
        table_bottom = hits_base[-1].y1
    else:
        table_bottom = fitz_page.rect.height * 0.42

    # ★ 在表格底部和总保费之间的空白中间写一行
    text_y = table_bottom + (total_y - table_bottom) * 0.45

    x_pos = 36

    # ── 下划线 ──
    if base_row_line:
        _draw_underline(fitz_page, base_row_line, line_width=1.2)
    if extra_row_line:
        _draw_underline(fitz_page, extra_row_line, line_width=1.2)

    # ── 写一行红字 ──
    _write(fitz_page, line_all, x_pos, text_y, RED, font_path, fontsize=12)

    # ── 箭头：从"保額是XXW美金"末尾指向保额数字 ──
    if base_num_rect:
        prefix = f"保額是{base_str}"
        ax0 = x_pos + len(prefix) * 12 * 0.52
        ay0 = text_y - 6

        ax1 = base_num_rect.x0
        ay1 = (base_num_rect.y0 + base_num_rect.y1) / 2

        if ax0 >= ax1 - 10:
            ax0 = ax1 - 20
            ay0 = text_y - 4

        shape2 = fitz_page.new_shape()
        shape2.draw_line(fitz.Point(ax0, ay0), fitz.Point(ax1, ay1))
        shape2.finish(color=RED, width=1.2, closePath=False)
        shape2.commit()
        _draw_arrowhead(fitz_page, ax0, ay0, ax1, ay1)


# ═══════════════════════════════════════════════════════════════
# 重疾险 2. 說明摘要页标识 —— 一行红字写在右侧空白，不挡原文
# ═══════════════════════════════════════════════════════════════
def _annotate_summary(fitz_page, words, policy, font_path):
    w_col12 = next((w for w in words if w["text"] == "(1)+(2)"),  None)
    w_col34 = next((w for w in words if w["text"] == "(3)+(4)"),  None)
    w_100   = next((w for w in words if "100歲"   in w["text"]), None)
    w_note  = next((w for w in words if "上述年齡" in w["text"]), None)

    premium_str = f"{int(policy.annual_premium):,}" if policy.annual_premium else "24,170"
    years_str   = str(int(policy.payment_years))    if policy.payment_years  else "10"

    pw = fitz_page.rect.width

    # ★ 合并成一行
    line_one = f"每年交{premium_str}，交{years_str}年不用再交"

    # 找"說明摘要"或"基本計劃"作为锚点
    hits_title = fitz_page.search_for("說明摘要")
    if not hits_title:
        hits_title = fitz_page.search_for("基本計劃")

    # 找表格表头作为下边界
    hits_header = fitz_page.search_for("保單年度")
    if not hits_header:
        hits_header = fitz_page.search_for("已繳保費")

    if hits_title and hits_header:
        title_y  = hits_title[0].y1
        header_y = hits_header[0].y0
        text_y   = title_y + (header_y - title_y) * 0.55
        # ★ x 从标题右边一段距离开始，避免遮挡"基本計劃 說明摘要"
        text_x = hits_title[0].x1 + 30
        if text_x + len(line_one) * 10 > pw - 20:
            text_x = max(hits_title[0].x1 + 15, pw * 0.30)
        _write(fitz_page, line_one, text_x, text_y, RED, font_path, fontsize=10)
    else:
        w_paid = next((w for w in words if "已繳保費" in w["text"]), None)
        if w_paid:
            _write(fitz_page, line_one, pw * 0.35, w_paid["top"] - 8, RED, font_path, fontsize=10)

    # 列下方的标签
    if w_col12 and w_100:
        _write(fitz_page, "預計的退保價值",
               w_col12["x0"] - 10, w_100["bottom"] + 12, ORANGE, font_path, fontsize=9)
    if w_col34 and w_100:
        _write(fitz_page, "預計的理賠金額",
               w_col34["x0"] - 10, w_100["bottom"] + 12, GREEN, font_path, fontsize=9)

    # 两个红框
    hits_12  = fitz_page.search_for("(1)+(2)")
    hits_34  = fitz_page.search_for("(3)+(4)")
    hits_100 = fitz_page.search_for("100歲")
    if hits_12 and hits_34 and hits_100:
        r12  = hits_12[0]; r34 = hits_34[0]; r100 = hits_100[-1]
        hits_header2 = fitz_page.search_for("退保發還金額")
        table_top    = hits_header2[0].y0 - 2 if hits_header2 else r12.y0 - 18
        table_bottom = r100.y1 + 2
        _draw_red_box(fitz_page, fitz.Rect(r12.x0 - 4, table_top, r12.x1 + 4, table_bottom), 1.5)
        _draw_red_box(fitz_page, fitz.Rect(r34.x0 - 4, table_top, r34.x1 + 4, table_bottom), 1.5)

    slogan_y = (w_note["bottom"] + 20) if w_note else 500
    _write_centered(fitz_page, "有事就賠錢，沒事就當存了筆錢",
                    slogan_y, RED, font_path, fontsize=11)


# ═══════════════════════════════════════════════════════════════
# 重疾险 3. 多重赔付页 & 4. 持续癌症页
# ═══════════════════════════════════════════════════════════════
def _annotate_multi(fitz_page, words, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 24) if content_words else 560
    _write_centered(fitz_page, "計劃本來自帶多次賠付，",                          last_y,    RED, font_path, fontsize=12)
    _write_centered(fitz_page, "意思是萬一理賠過重疾了，不用再交費繼續有保障，",  last_y+18, RED, font_path, fontsize=12)
    _write_centered(fitz_page, "最多能賠9次",                                      last_y+36, RED, font_path, fontsize=12)

def _annotate_cancer(fitz_page, words, policy, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 24) if content_words else 620
    _write_centered(fitz_page, "還有針對大家最擔心的癌症，",                              last_y,    RED, font_path, fontsize=12)
    _write_centered(fitz_page, "如果患癌症了，理賠完重疾後",                              last_y+18, RED, font_path, fontsize=12)
    _write_centered(fitz_page, "如果一年未愈，能每月賠5%的保額，直到康復或最長100個月",   last_y+46, RED, font_path, fontsize=12)


# ═══════════════════════════════════════════════════════════════
# 字段提取
# ═══════════════════════════════════════════════════════════════
def extract_text(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: pages.append(t)
    return "\n".join(pages)

def extract_fields(text):
    patterns = {
        "annual_premium":  [r"投保時每年總保費[：:]\s*([\d,]+\.?\d*)", r"每年總保費[：:]\s*([\d,]+)"],
        "payment_years":   [r"保費繳付年期\s*(\d{1,2})\s*年"],
        "currency":        [r"保單貨幣[：:]\s*(\S+)"],
        "insured_name":    [r"被保人姓名[：:]\s*(.+?)(?:\n|先生|女士)"],
        "insured_age":     [r"年齡[：:]\s*(\d+)"],
        "applicant_name":  [r"申請人姓名[：:]\s*(.+?)(?:\n|年齡)"],
    }
    fields = {}
    for name, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, text, re.MULTILINE | re.DOTALL)
            if m and m.lastindex:
                val = m.group(1).strip().replace(",", "")
                try:    fields[name] = float(val) if "." in val else int(val)
                except: fields[name] = val
                break
    return fields

def extract_fields_from_cover_page(pdf_path, debug=True):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            if not _is_cover_page(full_text): continue
            words = page.extract_words()
            m = re.search(r"投保時每年總保費[：:]?\s*([\d,]+\.?\d*)", full_text)
            if m:
                try: result["annual_premium"] = float(m.group(1).replace(",",""))
                except: pass
            m = re.search(r"保單貨幣[：:]?\s*(港幣|美金|人民幣|人民币)", full_text)
            if m: result["currency"] = m.group(1)
            rows = _group_by_rows(words, tolerance=4)
            def _num(s):
                try: return float(s.replace(",",""))
                except: return None
            base_row = None
            for y, row_ws in rows:
                texts = [w["text"] for w in row_ws]; joined = " ".join(texts)
                if (any(_num(t) and _num(t) >= 100000 for t in texts) and
                    any(re.fullmatch(r"\d{1,2}年", t) for t in texts) and
                    ("HGS" in joined or "危疾保障" in joined)):
                    base_row = row_ws
                    if debug: print(f"📍 基本计划行 y={y:.1f}: {joined}")
                    break
            if base_row is None:
                for y, row_ws in rows:
                    texts = [w["text"] for w in row_ws]
                    if (any(_num(t) and _num(t) >= 100000 for t in texts) and
                        any(re.fullmatch(r"\d{1,2}年", t) for t in texts)):
                        base_row = row_ws
                        if debug: print(f"📍 基本计划行(退化) y={y:.1f}: {' '.join(texts)}")
                        break
            if base_row:
                nums_numeric = []; years_val = None
                for w in base_row:
                    t = w["text"]
                    m_yr = re.fullmatch(r"(\d{1,2})年", t)
                    if m_yr:
                        yr = int(m_yr.group(1))
                        if 1 <= yr <= 40: years_val = yr
                        continue
                    v = _num(t)
                    if v is not None: nums_numeric.append((w["x0"], v))
                nums_numeric.sort(key=lambda t: t[0])
                nums = [v for _, v in nums_numeric]
                if len(nums) >= 1 and nums[0] >= 10000:  result["base_sum_insured"] = nums[0]
                if len(nums) >= 2 and nums[1] > 100 and "annual_premium" not in result:
                    result["annual_premium"] = nums[1]
                if len(nums) >= 3 and 60 <= int(nums[-1]) <= 120: result["coverage_age"] = int(nums[-1])
                if years_val: result["payment_years"] = years_val
            extra_row = None
            for y, row_ws in rows:
                if "額外保障" in " ".join(w["text"] for w in row_ws):
                    extra_row = row_ws
                    if debug: print(f"📍 额外保障行 y={y:.1f}")
                    break
            if extra_row:
                joined_extra = " ".join(w["text"] for w in extra_row)
                m_ext_yr = re.search(r"首\s*(\d{1,2})\s*年", joined_extra)
                if m_ext_yr: result["extra_years"] = int(m_ext_yr.group(1))
                for w in extra_row:
                    v = _num(w["text"])
                    if v is not None and v >= 10000: result["extra_sum_insured"] = v; break
            if result.get("base_sum_insured") and result.get("extra_sum_insured"):
                ratio = result["extra_sum_insured"] / result["base_sum_insured"] * 100
                result["extra_ratio"] = int(round(ratio / 10) * 10)
            if "payment_years" not in result:
                m = re.search(r"(\d{1,2})\s*年\s+1?00\b", full_text)
                if m:
                    yr = int(m.group(1))
                    if 1 <= yr <= 40: result["payment_years"] = yr
            if debug: print(f"✅ 首页提取结果: {result}")
            break
    return result

def extract_fields_from_summary_page(pdf_path):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(); full_text = page.extract_text() or ""
            if "保費繳付年期" not in full_text and "繳費年期" not in full_text: continue
            m = re.search(r"投保時每年總保費[：:]\s*([\d,]+\.?\d*)", full_text)
            if m:
                try: result["annual_premium"] = float(m.group(1).replace(",",""))
                except: pass
            for i, w in enumerate(words):
                if "繳付" in w["text"] or "年期" in w["text"]:
                    for j in range(i+1, min(i+8, len(words))):
                        tok = words[j]["text"].replace(",","").replace("年","").strip()
                        if re.fullmatch(r"\d{1,2}", tok):
                            yr = int(tok)
                            if 1 <= yr <= 30: result["payment_years"] = yr; break
                    if "payment_years" in result: break
            if "payment_years" not in result:
                m2 = re.search(r"保費繳付年期\s*(\d{1,2})\s*年", full_text, re.MULTILINE)
                if m2:
                    yr = int(m2.group(1))
                    if 1 <= yr <= 30: result["payment_years"] = yr
            for w in words:
                if w["text"] in ["美金","港幣","人民幣"]: result["currency"] = w["text"]; break
            if "payment_years" in result: break
    return result


# ═══════════════════════════════════════════════════════════════
# 重疾险主入口
# ═══════════════════════════════════════════════════════════════
def annotate_critical_illness_pdf(input_pdf_path, policy, font_path=None):
    if font_path is None: font_path = find_chinese_font()
    fitz_doc = fitz.open(input_pdf_path)
    fitz_doc = redact_personal_info(fitz_doc, is_savings=False)
    with pdfplumber.open(input_pdf_path) as pl_doc:
        for page_idx in range(len(fitz_doc)):
            fitz_page = fitz_doc[page_idx]
            pl_page   = pl_doc.pages[page_idx]
            words      = pl_page.extract_words()
            words_text = " ".join(w["text"] for w in words)
            full_text  = pl_page.extract_text() or ""
            is_cover   = _is_cover_page(full_text)
            is_summary = _is_summary_page(words_text)
            is_multi   = _is_multi_page(full_text)
            is_cancer  = _is_cancer_page(full_text)
            if not (is_cover or is_summary or is_multi or is_cancer): continue
            print(f"第{page_idx+1}页 cover={is_cover} summary={is_summary} multi={is_multi} cancer={is_cancer}")
            if is_cover:   _annotate_cover(fitz_page, words, policy, font_path)
            if is_summary: _annotate_summary(fitz_page, words, policy, font_path)
            if is_multi:   _annotate_multi(fitz_page, words, font_path)
            if is_cancer:  _annotate_cancer(fitz_page, words, policy, font_path)
    output = io.BytesIO()
    fitz_doc.save(output, garbage=4, deflate=True, clean=True)
    fitz_doc.close()
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# 储蓄险数据提取
# ═══════════════════════════════════════════════════════════════
def extract_supplement_table(pdf_path, log=print):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text or not _is_supplement_no_withdrawal(text): continue
            lines = text.split("\n")
            for line in lines:
                line = line.strip().replace(",","")
                tokens = line.split()
                if not tokens or not tokens[0].isdigit(): continue
                year = int(tokens[0])
                if year < 1 or year > 99 or len(tokens) < 9: continue
                def parse_val(s):
                    try: return int(s.replace("-","0").replace(",",""))
                    except: return 0
                while len(tokens) < 10: tokens.append("0")
                all_rows.append({
                    "year": year, "paid_total": parse_val(tokens[1]),
                    "cash_value": parse_val(tokens[2]), "bonus_cv": parse_val(tokens[3]),
                    "terminal_cv": parse_val(tokens[4]), "surrender_total": parse_val(tokens[5]),
                    "death_benefit": parse_val(tokens[6]), "bonus_fv": parse_val(tokens[7]),
                    "terminal_fv": parse_val(tokens[8]), "death_total": parse_val(tokens[9]),
                })
    if not all_rows: log("⚠️ 未找到数据"); return pd.DataFrame()
    df = pd.DataFrame(all_rows).drop_duplicates(subset="year").sort_values("year").reset_index(drop=True)
    log(f"✅ 成功提取 {len(df)} 年数据")
    return df

def find_key_milestones(df, log=print):
    milestones = []; base_paid = df["paid_total"].max()
    be = df[df["surrender_total"] >= df["paid_total"]]
    if not be.empty:
        y = int(be.iloc[0]["year"])
        milestones.append({"year": y, "label": f"預計第{y}年保本",   "color": (0.85, 0.1, 0.1)})
    d2 = df[df["surrender_total"] >= 2 * base_paid]
    if not d2.empty:
        y = int(d2.iloc[0]["year"])
        milestones.append({"year": y, "label": f"預計第{y}年翻倍",   "color": (0.7, 0.1, 0.8)})
    d4 = df[df["surrender_total"] >= 4 * base_paid]
    if not d4.empty:
        y = int(d4.iloc[0]["year"])
        milestones.append({"year": y, "label": f"預計第{y}年再翻倍", "color": (0.1, 0.55, 0.1)})
    log(f"🔍 识别到 {len(milestones)} 个关键节点")
    return milestones

def _parse_withdrawal_info(pdf_path, log=print):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not _is_supplement_with_withdrawal(text): continue
            m_cur = re.search(r"保單貨幣[：:]\s*(美金|港幣|人民幣)", text)
            if m_cur: result["currency"] = m_cur.group(1)
            lines = text.split("\n"); year_amounts = []
            for line in lines:
                tokens = line.strip().split()
                if not tokens or not tokens[0].isdigit(): continue
                year = int(tokens[0])
                if year < 1 or year > 99: continue
                clean = [t.replace(",","") for t in tokens]; amounts = []
                for t in clean[1:]:
                    if t == "-": amounts.append(0)
                    else:
                        try: amounts.append(int(t))
                        except: amounts.append(None)
                if len(amounts) >= 4 and amounts[3] and amounts[3] > 0:
                    year_amounts.append((year, amounts[3]))
            if year_amounts:
                result["start_year"] = year_amounts[0][0]
                all_amounts = [a for _, a in year_amounts if a > 0]
                if all_amounts: result["annual_amount"] = Counter(all_amounts).most_common(1)[0][0]
            if result: break
    log(f"📊 提取信息解析: {result}")
    return result


# ═══════════════════════════════════════════════════════════════
# 储蓄险：无提取版标注
# ═══════════════════════════════════════════════════════════════
def _annotate_milestone_rows(fitz_page, milestones, font_path,
                              col_header_text="(1)+(2)+(3)", fallback_col_ratio=0.52):
    pw = fitz_page.rect.width
    text_dict = fitz_page.get_text("dict")
    total_col_x1 = None
    for block in text_dict["blocks"]:
        if block.get("type") != 0: continue
        for line in block["lines"]:
            line_text = "".join(s["text"] for s in line["spans"]).replace(" ","")
            if col_header_text.replace(" ","") in line_text:
                total_col_x1 = max(s["bbox"][2] for s in line["spans"]) + 4; break
        if total_col_x1: break
    if total_col_x1 is None: total_col_x1 = pw * fallback_col_ratio
    table_left = pw
    for block in text_dict["blocks"]:
        if block.get("type") != 0: continue
        for line in block["lines"]:
            spans = line["spans"]
            if spans and re.match(r"^\d{1,2}$", spans[0]["text"].strip().replace(",","")):
                table_left = min(table_left, spans[0]["bbox"][0] - 2)
    if table_left == pw: table_left = 10
    for ms in milestones:
        target_year = str(ms["year"]); label = ms["label"]; r, g, b = ms["color"]
        for block in text_dict["blocks"]:
            if block.get("type") != 0: continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans or spans[0]["text"].strip().replace(",","") != target_year: continue
                y0, y1 = line["bbox"][1], line["bbox"][3]; row_h = max(y1-y0, 8)
                shape = fitz_page.new_shape()
                shape.draw_rect(fitz.Rect(table_left, y0-1.5, total_col_x1, y1+1.5))
                shape.finish(color=(r,g,b), fill=None, width=1.2); shape.commit()
                bubble_w = min(max(len(label)*7+10, 80), 120)
                bx0 = total_col_x1+3; bx1 = bx0+bubble_w
                if bx1 > pw-5: bx1 = pw-5; bx0 = bx1-bubble_w
                bubble = fitz.Rect(bx0, y0-2, bx1, y0-2+row_h+4)
                shape2 = fitz_page.new_shape()
                shape2.draw_rect(bubble)
                shape2.finish(fill=(r,g,b), fill_opacity=0.88, color=(r,g,b), width=0.5); shape2.commit()
                kw = dict(fontsize=6.5, color=(1,1,1), align=fitz.TEXT_ALIGN_CENTER)
                if font_path: kw["fontfile"] = font_path; kw["fontname"] = "cjk"
                fitz_page.insert_textbox(bubble, label, **kw)


# ═══════════════════════════════════════════════════════════════
# 储蓄险：有提取版标注
# ═══════════════════════════════════════════════════════════════
def _annotate_withdrawal_page(fitz_page, withdrawal_info, font_path):
    pw = fitz_page.rect.width
    ph = fitz_page.rect.height
    text_dict = fitz_page.get_text("dict")
    col1_x0 = None; col12_x1 = None
    col345_x0 = None; col345_x1 = None
    col_header_y = None
    for block in text_dict["blocks"]:
        if block.get("type") != 0: continue
        for line in block["lines"]:
            line_text = "".join(s["text"] for s in line["spans"]).replace(" ","")
            if "(1)" in line_text and "(1)+(2)" not in line_text and col1_x0 is None:
                for span in line["spans"]:
                    if span["text"].replace(" ","") == "(1)":
                        col1_x0 = span["bbox"][0]-3
                        if col_header_y is None: col_header_y = line["bbox"][1]
                        break
            if "(1)+(2)" in line_text and col12_x1 is None:
                for span in line["spans"]:
                    if "(1)+(2)" in span["text"].replace(" ",""):
                        col12_x1 = span["bbox"][2]+3
                        if col_header_y is None: col_header_y = line["bbox"][1]
                        break
            if "(3)+(4)+(5)" in line_text and col345_x0 is None:
                for span in line["spans"]:
                    if "(3)+(4)+(5)" in span["text"].replace(" ",""):
                        col345_x0 = span["bbox"][0]-3
                        col345_x1 = span["bbox"][2]+3
                        if col_header_y is None: col_header_y = line["bbox"][1]
                        break
    if col1_x0      is None: col1_x0      = pw * 0.22
    if col12_x1     is None: col12_x1     = pw * 0.42
    if col345_x0    is None: col345_x0    = pw * 0.72
    if col345_x1    is None: col345_x1    = pw * 0.92
    if col_header_y is None: col_header_y = ph * 0.15
    last_data_y1 = ph * 0.85
    for block in text_dict["blocks"]:
        if block.get("type") != 0: continue
        for line in block["lines"]:
            spans = line["spans"]
            if spans and re.match(r"^\d{1,2}$", spans[0]["text"].strip()):
                last_data_y1 = max(last_data_y1, line["bbox"][3])
    shape = fitz_page.new_shape()
    shape.draw_rect(fitz.Rect(col1_x0, col_header_y-2, col12_x1, last_data_y1+2))
    shape.finish(color=RED, fill=None, width=1.5); shape.commit()
    shape = fitz_page.new_shape()
    shape.draw_rect(fitz.Rect(col345_x0, col_header_y-2, col345_x1, last_data_y1+2))
    shape.finish(color=RED, fill=None, width=1.5); shape.commit()
    start_year    = withdrawal_info.get("start_year", 0)
    annual_amount = withdrawal_info.get("annual_amount", 0)
    currency      = withdrawal_info.get("currency", "")
    cur_str = {"美金":"USD","港幣":"HKD","人民幣":"RMB"}.get(currency, currency)
    left_text  = f"第{start_year}年開始，每年提取{int(annual_amount):,}{cur_str}" if start_year > 0 and annual_amount > 0 else "開始提取後"
    right_text = "提取後保單預計繼續增值"
    label_y    = last_data_y1 + 20
    kw_red = dict(fontsize=11, color=RED)
    if font_path: kw_red["fontfile"] = font_path; kw_red["fontname"] = "cjk"
    fitz_page.insert_text((20, label_y), left_text, **kw_red)
    fitz_page.insert_text((pw*0.55, label_y), right_text, **kw_red)


# ═══════════════════════════════════════════════════════════════
# 储蓄险主入口
# ═══════════════════════════════════════════════════════════════
def annotate_savings_pdf(input_pdf_path, milestones, font_path=None, log=print):
    if font_path is None: font_path = find_chinese_font()
    withdrawal_info = _parse_withdrawal_info(input_pdf_path, log=log)
    doc = fitz.open(input_pdf_path)
    doc = redact_personal_info(doc, is_savings=True)
    for page in doc:
        full_text = page.get_text("text")
        is_no_wd   = _is_supplement_no_withdrawal(full_text)
        is_with_wd = _is_supplement_with_withdrawal(full_text)
        if not (is_no_wd or is_with_wd): continue
        log(f"  📄 处理页面：no_withdrawal={is_no_wd}  with_withdrawal={is_with_wd}")
        if is_no_wd:
            _annotate_milestone_rows(page, milestones, font_path,
                                     col_header_text="(1)+(2)+(3)", fallback_col_ratio=0.52)
        elif is_with_wd:
            _annotate_withdrawal_page(page, withdrawal_info, font_path)
    output = io.BytesIO()
    doc.save(output, garbage=4, deflate=True, clean=True)
    doc.close()
    return output.getvalue()
