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
        if os.path.exists(p):
            return p
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
        if os.path.exists(p):
            return p
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
    try:
        return page.insert_text((x, y), text, **kw)
    except Exception:
        return -1

def _write_centered(page, text, y, color, font_path, fontsize=10):
    pw = page.rect.width
    x  = max(10, (pw - len(text) * fontsize * 0.82) / 2)
    return _write(page, text, x, y, color, font_path, fontsize)

def _format_wan(amount, currency=""):
    if not amount or amount <= 0:
        return "XXX"
    cur = currency or ""
    wan = amount / 10000
    if wan >= 1 and wan == int(wan):
        return f"{int(wan)}W{cur}"
    elif wan >= 1:
        return f"{wan:.0f}W{cur}"
    else:
        return f"{int(amount):,}{cur}"


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

def _draw_red_box(fitz_page, rect, line_width=1.5):
    shape = fitz_page.new_shape()
    shape.draw_rect(rect)
    shape.finish(color=RED, fill=None, width=line_width)
    shape.commit()

def _draw_underline(fitz_page, rect, line_width=1.5):
    shape = fitz_page.new_shape()
    shape.draw_line(
        fitz.Point(rect.x0, rect.y1),
        fitz.Point(rect.x1, rect.y1)
    )
    shape.finish(color=RED, width=line_width, closePath=False)
    shape.commit()


# ═══════════════════════════════════════════════════════════════
# 行分桶工具
# ═══════════════════════════════════════════════════════════════
def _group_by_rows(words, tolerance=4):
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: w["top"])
    rows = []
    for w in sorted_words:
        placed = False
        for row in rows:
            if abs(row["y"] - w["top"]) <= tolerance:
                row["words"].append(w)
                row["y"] = (row["y"] * (len(row["words"]) - 1) + w["top"]) / len(row["words"])
                placed = True
                break
        if not placed:
            rows.append({"y": w["top"], "words": [w]})
    rows.sort(key=lambda r: r["y"])
    return [(r["y"], sorted(r["words"], key=lambda w: w["x0"])) for r in rows]


# ═══════════════════════════════════════════════════════════════
# 箭头辅助
# ═══════════════════════════════════════════════════════════════
def _draw_arrowhead(page, x0, y0, x1, y1, size=6):
    dx = x1 - x0
    dy = y1 - y0
    length = math.hypot(dx, dy)
    if length == 0:
        return
    ux = dx / length
    uy = dy / length
    vx = -uy
    vy =  ux
    p1  = fitz.Point(x1 - size * ux + size * 0.5 * vx,
                     y1 - size * uy + size * 0.5 * vy)
    p2  = fitz.Point(x1 - size * ux - size * 0.5 * vx,
                     y1 - size * uy - size * 0.5 * vy)
    tip = fitz.Point(x1, y1)
    shape = page.new_shape()
    shape.draw_polyline([p1, tip, p2])
    shape.finish(color=RED, fill=RED, width=1.0, closePath=False)
    shape.commit()


# ═══════════════════════════════════════════════════════════════
# 1. 保障摘要页（首页）标识
# ═══════════════════════════════════════════════════════════════
def _annotate_cover(fitz_page, words, policy, font_path):
    hits_extra = fitz_page.search_for("額外保障")
    hits_base  = fitz_page.search_for("愛唯守危疾保障")

    base_str  = _format_wan(policy.base_sum_insured, policy.currency)
    ratio     = policy.extra_ratio or 50
    ext_years = policy.extra_years or 10
    line1 = f"保額是{base_str}，首{ext_years}年額外贈送{ratio}%保額"

    premium     = int(policy.annual_premium) if policy.annual_premium else 0
    premium_str = f"{premium:,}" if premium > 0 else "XXXX"
    years_str   = str(int(policy.payment_years)) if policy.payment_years else "XX"
    age_str     = str(int(policy.coverage_age))  if policy.coverage_age  else "100"
    line2 = f"年保費是{premium_str}，交{years_str}年，保到{age_str}歲"

    pw = fitz_page.rect.width

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

    available = total_y - table_bottom - 8
    gap       = max(available / 2.2, 20)

    y1 = table_bottom + gap * 0.5
    y2 = y1 + gap

    x1_pos = 36
    x2_pos = pw * 0.40

    if base_row_line:
        _draw_underline(fitz_page, base_row_line, line_width=1.2)
    if extra_row_line:
        _draw_underline(fitz_page, extra_row_line, line_width=1.2)

    _write(fitz_page, line1, x1_pos, y1, RED, font_path, fontsize=12)
    _write(fitz_page, line2, x2_pos, y2, RED, font_path, fontsize=12)

    if base_num_rect:
        ax0 = x1_pos + len(line1) * 12 * 0.52
        ay0 = y1 - 6
        ax1 = base_num_rect.x0
        ay1 = (base_num_rect.y0 + base_num_rect.y1) / 2

        if ax0 >= ax1 - 10:
            ax0 = ax1 - 20
            ay0 = y1 - 4

        shape2 = fitz_page.new_shape()
        shape2.draw_line(fitz.Point(ax0, ay0), fitz.Point(ax1, ay1))
        shape2.finish(color=RED, width=1.2, closePath=False)
        shape2.commit()
        _draw_arrowhead(fitz_page, ax0, ay0, ax1, ay1)


# ═══════════════════════════════════════════════════════════════
# 2. 說明摘要页标识
# ═══════════════════════════════════════════════════════════════
def _annotate_summary(fitz_page, words, policy, font_path):
    w_paid  = next((w for w in words if "已繳保費" in w["text"]), None)
    w_col12 = next((w for w in words if w["text"] == "(1)+(2)"),  None)
    w_col34 = next((w for w in words if w["text"] == "(3)+(4)"),  None)
    w_100   = next((w for w in words if "100歲"   in w["text"]), None)
    w_note  = next((w for w in words if "上述年齡" in w["text"]), None)

    premium_str = f"{int(policy.annual_premium):,}" if policy.annual_premium else "24,170"
    years_str   = str(int(policy.payment_years))    if policy.payment_years  else "10"
    if w_paid:
        by = w_paid["top"]
        _write(fitz_page, f"每年交{premium_str}",     49, by - 12, RED, font_path, fontsize=9)
        _write(fitz_page, f"交{years_str}年不用再交", 49, by - 2,  RED, font_path, fontsize=9)

    if w_col12 and w_100:
        _write(fitz_page, "預計的退保價值",
               w_col12["x0"] - 10, w_100["bottom"] + 12,
               ORANGE, font_path, fontsize=9)

    if w_col34 and w_100:
        _write(fitz_page, "預計的理賠金額",
               w_col34["x0"] - 10, w_100["bottom"] + 12,
               GREEN, font_path, fontsize=9)

    hits_12  = fitz_page.search_for("(1)+(2)")
    hits_34  = fitz_page.search_for("(3)+(4)")
    hits_100 = fitz_page.search_for("100歲")

    if hits_12 and hits_34 and hits_100:
        r12  = hits_12[0]
        r34  = hits_34[0]
        r100 = hits_100[-1]

        hits_header  = fitz_page.search_for("退保發還金額")
        table_top    = hits_header[0].y0 - 2 if hits_header else r12.y0 - 18
        table_bottom = r100.y1 + 2

        rect_box1 = fitz.Rect(r12.x0 - 4, table_top, r12.x1 + 4, table_bottom)
        _draw_red_box(fitz_page, rect_box1, line_width=1.5)

        rect_box2 = fitz.Rect(r34.x0 - 4, table_top, r34.x1 + 4, table_bottom)
        _draw_red_box(fitz_page, rect_box2, line_width=1.5)

    slogan_y = (w_note["bottom"] + 20) if w_note else 500
    _write_centered(fitz_page, "有事就賠錢，沒事就當存了筆錢",
                    slogan_y, RED, font_path, fontsize=11)


# ═══════════════════════════════════════════════════════════════
# 3. 多重赔付页标识
# ═══════════════════════════════════════════════════════════════
def _annotate_multi(fitz_page, words, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 24) if content_words else 560

    _write_centered(fitz_page, "計劃本來自帶多次賠付，",
                    last_y,       RED, font_path, fontsize=12)
    _write_centered(fitz_page, "意思是萬一理賠過重疾了，不用再交費繼續有保障，",
                    last_y + 18,  RED, font_path, fontsize=12)
    _write_centered(fitz_page, "最多能賠9次",
                    last_y + 36,  RED, font_path, fontsize=12)


# ═══════════════════════════════════════════════════════════════
# 4. 持续癌症页标识
# ═══════════════════════════════════════════════════════════════
def _annotate_cancer(fitz_page, words, policy, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 24) if content_words else 620

    _write_centered(fitz_page, "還有針對大家最擔心的癌症，",
                    last_y,       RED, font_path, fontsize=12)
    _write_centered(fitz_page, "如果患癌症了，理賠完重疾後",
                    last_y + 18,  RED, font_path, fontsize=12)
    _write_centered(fitz_page, "如果一年未愈，能每月賠5%的保額，直到康復或最長100個月",
                    last_y + 46,  RED, font_path, fontsize=12)


# ═══════════════════════════════════════════════════════════════
# 个人信息遮盖（两险共用）
# ═══════════════════════════════════════════════════════════════
def _has_barcode_image(page: fitz.Page) -> bool:
    """检测页面右上角区域是否有图片（条形码）"""
    pw = page.rect.width
    ph = page.rect.height
    top_right = fitz.Rect(pw * 0.25, 0, pw, ph * 0.20)
    for img in page.get_image_info():
        bbox = fitz.Rect(img["bbox"])
        if bbox.intersects(top_right) and bbox.width > 50 and bbox.height > 5:
            return True
    return False

def _page_has_barcode_keyword(page: fitz.Page) -> bool:
    """
    通过页面关键词判断是否是需要遮条形码的页面：
    储蓄险封面（保障摘要）、储蓄险数据页（補充說明摘要）、重疾险封面
    """
    text = page.get_text("text")
    return (
        _is_cover_page(text) or
        "補充說明摘要" in text or
        "保障摘要"     in text
    )

def redact_personal_info(doc: fitz.Document) -> fitz.Document:
    WHITE = (1, 1, 1)

    # 预提取保单号（用于文字精确定位）
    policy_number = None
    for page in doc:
        text = page.get_text("text")
        m = re.search(r"[A-Z]{2}\d{6}-\d{8,}-\d", text)
        if m:
            policy_number = m.group(0)
            break

    for page_num in range(len(doc)):
        page = doc[page_num]
        pw   = page.rect.width
        ph   = page.rect.height

        redact_rects    = []
        barcode_covered = False

        # ── 策略1：文字保单号精确定位（数据页）──
        if policy_number:
            hits = page.search_for(policy_number)
            if hits:
                for rect in hits:
                    redact_rects.append(
                        fitz.Rect(pw * 0.25, 0, pw, rect.y1 + 5)
                    )
                barcode_covered = True

        # ── 策略2：图片检测（封面条形码是嵌入图片）──
        if not barcode_covered and _has_barcode_image(page):
            redact_rects.append(
                fitz.Rect(pw * 0.25, 0, pw, ph * 0.15)
            )
            barcode_covered = True

        # ── 策略3：关键词兜底（只要是封面/摘要页，右上角固定遮盖）──
        if not barcode_covered and _page_has_barcode_keyword(page):
            redact_rects.append(
                fitz.Rect(pw * 0.25, 0, pw, ph * 0.15)
            )

        # ── 左下角页脚（被保人姓名）──
        hits_name = page.search_for("被保人姓名")
        if hits_name:
            r = hits_name[0]
            redact_rects.append(fitz.Rect(0, r.y0 - 1, pw * 0.38, ph))
        else:
            redact_rects.append(fitz.Rect(0, ph * 0.960, pw * 0.38, ph))

        for rect in redact_rects:
            page.add_redact_annot(rect, fill=WHITE)
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

    return doc


# ═══════════════════════════════════════════════════════════════
# 字段提取
# ═══════════════════════════════════════════════════════════════
def extract_text(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                pages.append(t)
    return "\n".join(pages)

def extract_fields(text):
    patterns = {
        "annual_premium": [
            r"投保時每年總保費[：:]\s*([\d,]+\.?\d*)",
            r"每年總保費[：:]\s*([\d,]+)",
        ],
        "payment_years": [
            r"保費繳付年期\s*(\d{1,2})\s*年",
            r"保費繳付年期[\s\S]{1,20}?(\d{1,2})\s*年",
            r"(?<!保障)(?<!至年齡\s)(\b[1-9]\d?\b)\s*年(?!\s*保障|\s*至)",
        ],
        "currency": [r"保單貨幣[：:]\s*(\S+)"],
        "continuous_cancer_monthly": [
            r"每月.*?(\d[\d,]+).*?持續癌症",
            r"持續癌症.*?每月.*?(\d[\d,]+)",
        ],
        "insured_name":   [r"被保人姓名[：:]\s*(.+?)(?:\n|先生|女士)"],
        "insured_age":    [r"年齡[：:]\s*(\d+)"],
        "applicant_name": [r"申請人姓名[：:]\s*(.+?)(?:\n|年齡)"],
    }
    fields = {}
    for name, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, text, re.MULTILINE | re.DOTALL)
            if m and m.lastindex:
                val = m.group(1).strip().replace(",", "")
                try:
                    fields[name] = float(val) if "." in val else int(val)
                except Exception:
                    fields[name] = val
                break
    return fields


def extract_fields_from_cover_page(pdf_path, debug=True):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            if not _is_cover_page(full_text):
                continue

            words = page.extract_words()

            m = re.search(r"投保時每年總保費[：:]?\s*([\d,]+\.?\d*)", full_text)
            if m:
                try:
                    result["annual_premium"] = float(m.group(1).replace(",", ""))
                except Exception:
                    pass

            m = re.search(r"保單貨幣[：:]?\s*(港幣|美金|人民幣|人民币)", full_text)
            if m:
                result["currency"] = m.group(1)

            rows = _group_by_rows(words, tolerance=4)

            def _num(s):
                try:
                    return float(s.replace(",", ""))
                except Exception:
                    return None

            base_row = None
            for y, row_ws in rows:
                texts  = [w["text"] for w in row_ws]
                joined = " ".join(texts)
                has_big_num = any(_num(t) and _num(t) >= 100000 for t in texts)
                has_years   = any(re.fullmatch(r"\d{1,2}年", t) for t in texts)
                has_hgs     = ("HGS" in joined or "危疾保障" in joined)
                if has_big_num and has_years and has_hgs:
                    base_row = row_ws
                    if debug:
                        print(f"📍 基本计划行 y={y:.1f}: {joined}")
                    break

            if base_row is None:
                for y, row_ws in rows:
                    texts   = [w["text"] for w in row_ws]
                    has_big = any(_num(t) and _num(t) >= 100000 for t in texts)
                    has_yr  = any(re.fullmatch(r"\d{1,2}年", t) for t in texts)
                    if has_big and has_yr:
                        base_row = row_ws
                        if debug:
                            print(f"📍 基本计划行(退化) y={y:.1f}: {' '.join(texts)}")
                        break

            if base_row:
                nums_numeric = []
                years_val    = None
                for w in base_row:
                    t    = w["text"]
                    m_yr = re.fullmatch(r"(\d{1,2})年", t)
                    if m_yr:
                        yr = int(m_yr.group(1))
                        if 1 <= yr <= 40:
                            years_val = yr
                        continue
                    v = _num(t)
                    if v is not None:
                        nums_numeric.append((w["x0"], v))
                nums_numeric.sort(key=lambda t: t[0])
                nums = [v for _, v in nums_numeric]
                if len(nums) >= 1 and nums[0] >= 10000:
                    result["base_sum_insured"] = nums[0]
                if len(nums) >= 2 and nums[1] > 100 and "annual_premium" not in result:
                    result["annual_premium"] = nums[1]
                if len(nums) >= 3 and 60 <= int(nums[-1]) <= 120:
                    result["coverage_age"] = int(nums[-1])
                if years_val:
                    result["payment_years"] = years_val

            extra_row = None
            for y, row_ws in rows:
                joined = " ".join(w["text"] for w in row_ws)
                if "額外保障" in joined:
                    extra_row = row_ws
                    if debug:
                        print(f"📍 额外保障行 y={y:.1f}: {joined}")
                    break

            if extra_row:
                joined_extra = " ".join(w["text"] for w in extra_row)
                m_ext_yr = re.search(r"首\s*(\d{1,2})\s*年", joined_extra)
                if m_ext_yr:
                    result["extra_years"] = int(m_ext_yr.group(1))
                for w in extra_row:
                    v = _num(w["text"])
                    if v is not None and v >= 10000:
                        result["extra_sum_insured"] = v
                        break

            if result.get("base_sum_insured") and result.get("extra_sum_insured"):
                ratio = result["extra_sum_insured"] / result["base_sum_insured"] * 100
                result["extra_ratio"] = int(round(ratio / 10) * 10)

            if "payment_years" not in result:
                m = re.search(r"(\d{1,2})\s*年\s+1?00\b", full_text)
                if m:
                    yr = int(m.group(1))
                    if 1 <= yr <= 40:
                        result["payment_years"] = yr

            if debug:
                print(f"✅ 首页提取结果: {result}")
            break

    return result


def extract_fields_from_summary_page(pdf_path):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words     = page.extract_words()
            full_text = page.extract_text() or ""

            if "保費繳付年期" not in full_text and "繳費年期" not in full_text:
                continue

            m = re.search(r"投保時每年總保費[：:]\s*([\d,]+\.?\d*)", full_text)
            if m:
                try:
                    result["annual_premium"] = float(m.group(1).replace(",", ""))
                except Exception:
                    pass

            for i, w in enumerate(words):
                if "繳付" in w["text"] or "年期" in w["text"]:
                    for j in range(i + 1, min(i + 8, len(words))):
                        tok = words[j]["text"].replace(",", "").replace("年", "").strip()
                        if re.fullmatch(r"\d{1,2}", tok):
                            yr = int(tok)
                            if 1 <= yr <= 30:
                                result["payment_years"] = yr
                                break
                    if "payment_years" in result:
                        break

            if "payment_years" not in result:
                m2 = re.search(r"保費繳付年期\s*(\d{1,2})\s*年", full_text, re.MULTILINE)
                if m2:
                    yr = int(m2.group(1))
                    if 1 <= yr <= 30:
                        result["payment_years"] = yr

            if "payment_years" not in result:
                for m3 in re.finditer(r"\b(\d{1,2})\s*年", full_text):
                    yr = int(m3.group(1))
                    if 1 <= yr <= 30:
                        result["payment_years"] = yr
                        break

            for w in words:
                if w["text"] in ["美金", "港幣", "人民幣"]:
                    result["currency"] = w["text"]
                    break

            if "payment_years" in result:
                break

    return result


# ═══════════════════════════════════════════════════════════════
# 主入口（重疾险）
# ═══════════════════════════════════════════════════════════════
def annotate_critical_illness_pdf(input_pdf_path, policy, font_path=None):
    if font_path is None:
        font_path = find_chinese_font()

    fitz_doc = fitz.open(input_pdf_path)
    fitz_doc = redact_personal_info(fitz_doc)

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

            if not (is_cover or is_summary or is_multi or is_cancer):
                continue

            print(f"第 {page_idx+1} 页  cover={is_cover}  summary={is_summary}  "
                  f"multi={is_multi}  cancer={is_cancer}")

            if is_cover:
                _annotate_cover(fitz_page, words, policy, font_path)
            if is_summary:
                _annotate_summary(fitz_page, words, policy, font_path)
            if is_multi:
                _annotate_multi(fitz_page, words, font_path)
            if is_cancer:
                _annotate_cancer(fitz_page, words, policy, font_path)

    output = io.BytesIO()
    fitz_doc.save(output, garbage=4, deflate=True, clean=True)
    fitz_doc.close()
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# 储蓄险：数据提取（优先无提取版）
# ═══════════════════════════════════════════════════════════════
def extract_supplement_table(pdf_path, log=print):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            if not _is_supplement_no_withdrawal(text):
                continue

            lines = text.split("\n")
            for line in lines:
                line   = line.strip().replace(",", "")
                tokens = line.split()
                if not tokens or not tokens[0].isdigit():
                    continue
                year = int(tokens[0])
                if year < 1 or year > 99:
                    continue
                if len(tokens) < 9:
                    continue

                def parse_val(s):
                    s = s.replace("-", "0").replace(",", "")
                    try:
                        return int(s)
                    except Exception:
                        return 0

                while len(tokens) < 10:
                    tokens.append("0")

                all_rows.append({
                    "year":            year,
                    "paid_total":      parse_val(tokens[1]),
                    "cash_value":      parse_val(tokens[2]),
                    "bonus_cv":        parse_val(tokens[3]),
                    "terminal_cv":     parse_val(tokens[4]),
                    "surrender_total": parse_val(tokens[5]),
                    "death_benefit":   parse_val(tokens[6]),
                    "bonus_fv":        parse_val(tokens[7]),
                    "terminal_fv":     parse_val(tokens[8]),
                    "death_total":     parse_val(tokens[9]),
                })

    if not all_rows:
        log("⚠️ 未找到数据")
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)
    df = df.drop_duplicates(subset="year").sort_values("year").reset_index(drop=True)
    log(f"✅ 成功提取 {len(df)} 年数据")
    return df


def find_key_milestones(df, log=print):
    milestones = []
    base_paid  = df["paid_total"].max()

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


# ═══════════════════════════════════════════════════════════════
# 储蓄险：解析提取信息
# ═══════════════════════════════════════════════════════════════
def _parse_withdrawal_info(pdf_path, log=print):
    """从有提取版页面解析：提取开始年份、每年提取金额、货币"""
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not _is_supplement_with_withdrawal(text):
                continue

            m_cur = re.search(r"保單貨幣[：:]\s*(美金|港幣|人民幣)", text)
            if m_cur:
                result["currency"] = m_cur.group(1)

            lines = text.split("\n")
            year_amounts = []

            for line in lines:
                tokens = line.strip().split()
                if not tokens or not tokens[0].isdigit():
                    continue
                year = int(tokens[0])
                if year < 1 or year > 99:
                    continue

                clean = [t.replace(",", "") for t in tokens]
                amounts = []
                for t in clean[1:]:
                    if t == "-":
                        amounts.append(0)
                    else:
                        try:
                            amounts.append(int(t))
                        except Exception:
                            amounts.append(None)

                if len(amounts) >= 4 and amounts[3] and amounts[3] > 0:
                    year_amounts.append((year, amounts[3]))

            if year_amounts:
                result["start_year"] = year_amounts[0][0]
                all_amounts = [a for _, a in year_amounts if a > 0]
                if all_amounts:
                    result["annual_amount"] = Counter(all_amounts).most_common(1)[0][0]

            if result:
                break

    log(f"📊 提取信息解析: {result}")
    return result


# ═══════════════════════════════════════════════════════════════
# 储蓄险：无提取版 - 行红框 + 气泡（横向框整行，不变）
# ═══════════════════════════════════════════════════════════════
def _annotate_milestone_rows(fitz_page, milestones, font_path,
                              col_header_text="(3)+(4)+(5)",
                              fallback_col_ratio=0.85):
    """无提取版：横向框住里程碑年份整行"""
    pw        = fitz_page.rect.width
    text_dict = fitz_page.get_text("dict")

    # 找总额列右边界
    total_col_x1 = None
    for block in text_dict["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            line_text = "".join(s["text"] for s in line["spans"]).replace(" ", "")
            if col_header_text.replace(" ", "") in line_text:
                rightmost    = max(s["bbox"][2] for s in line["spans"])
                total_col_x1 = rightmost + 4
                break
        if total_col_x1:
            break

    if total_col_x1 is None:
        total_col_x1 = pw * fallback_col_ratio

    # 找表格最左边界
    table_left = pw
    for block in text_dict["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            spans = line["spans"]
            if not spans:
                continue
            first_text = spans[0]["text"].strip().replace(",", "")
            if re.match(r"^\d{1,2}$", first_text):
                table_left = min(table_left, spans[0]["bbox"][0] - 2)
    if table_left == pw:
        table_left = 10

    for ms in milestones:
        target_year = str(ms["year"])
        label       = ms["label"]
        r, g, b     = ms["color"]

        for block in text_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans:
                    continue
                first_text = spans[0]["text"].strip().replace(",", "")
                if first_text != target_year:
                    continue

                line_bbox = line["bbox"]
                y0, y1    = line_bbox[1], line_bbox[3]
                row_h     = max(y1 - y0, 8)

                # 纯边框矩形，无填充
                box_rect = fitz.Rect(table_left, y0 - 1.5, total_col_x1, y1 + 1.5)
                shape = fitz_page.new_shape()
                shape.draw_rect(box_rect)
                shape.finish(color=(r, g, b), fill=None, width=1.2)
                shape.commit()

                # 气泡标签
                bubble_w = min(max(len(label) * 7 + 10, 80), 120)
                bx0 = total_col_x1 + 3
                bx1 = bx0 + bubble_w
                if bx1 > pw - 5:
                    bx1 = pw - 5
                    bx0 = bx1 - bubble_w
                bubble = fitz.Rect(bx0, y0 - 2, bx1, y0 - 2 + row_h + 4)

                shape2 = fitz_page.new_shape()
                shape2.draw_rect(bubble)
                shape2.finish(fill=(r, g, b), fill_opacity=0.88,
                             color=(r, g, b), width=0.5)
                shape2.commit()

                kw = dict(fontsize=6.5, color=(1, 1, 1),
                         align=fitz.TEXT_ALIGN_CENTER)
                if font_path:
                    kw["fontfile"] = font_path
                    kw["fontname"] = "cjk"
                fitz_page.insert_textbox(bubble, label, **kw)


# ═══════════════════════════════════════════════════════════════
# 储蓄险：有提取版 - 竖向列框（仿照目标图2效果）
# ═══════════════════════════════════════════════════════════════
def _annotate_withdrawal_page(fitz_page, milestones, withdrawal_info, font_path):
    """
    有提取版标注：
    1. 竖向红框圈住 (1)+(2) 提取总额列（从表头到最后一行）
    2. 竖向红框圈住 (3)+(4)+(5) 退保总额列（从表头到最后一行）
    3. 对里程碑年份行加横向红框 + 右侧气泡（与无提取版一致）
    4. 页面底部红字说明
    """
    pw        = fitz_page.rect.width
    ph        = fitz_page.rect.height
    text_dict = fitz_page.get_text("dict")

    # ── 第一步：定位所有列头，找列边界 ──
    col_bounds = {}  # key -> (x0, x1, header_y0, header_y1)

    for block in text_dict["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            line_text = "".join(s["text"] for s in line["spans"]).replace(" ", "")
            line_y0   = line["bbox"][1]
            line_y1   = line["bbox"][3]
            line_x0   = min(s["bbox"][0] for s in line["spans"])
            line_x1   = max(s["bbox"][2] for s in line["spans"])

            if "(1)+(2)" in line_text and "(1)+(2)" not in col_bounds:
                col_bounds["withdrawal"] = {
                    "x0": line_x0 - 4, "x1": line_x1 + 4,
                    "header_y0": line_y0, "header_y1": line_y1
                }
            if "(3)+(4)+(5)" in line_text and "(3)+(4)+(5)" not in col_bounds:
                col_bounds["surrender"] = {
                    "x0": line_x0 - 4, "x1": line_x1 + 4,
                    "header_y0": line_y0, "header_y1": line_y1
                }

    # ── 第二步：找数据区最后一行的 y1 ──
    last_data_y1 = ph * 0.85
    for block in text_dict["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            spans = line["spans"]
            if spans and re.match(r"^\d{1,2}$", spans[0]["text"].strip()):
                last_data_y1 = max(last_data_y1, line["bbox"][3])

    # ── 第三步：画竖向列框 ──
    # (1)+(2) 提取列 → 红色框
    if "withdrawal" in col_bounds:
        cb = col_bounds["withdrawal"]
        col_rect = fitz.Rect(cb["x0"], cb["header_y0"] - 2,
                             cb["x1"], last_data_y1 + 2)
        shape = fitz_page.new_shape()
        shape.draw_rect(col_rect)
        shape.finish(color=RED, fill=None, width=1.5)
        shape.commit()

    # (3)+(4)+(5) 退保总额列 → 红色框
    if "surrender" in col_bounds:
        cb = col_bounds["surrender"]
        col_rect = fitz.Rect(cb["x0"], cb["header_y0"] - 2,
                             cb["x1"], last_data_y1 + 2)
        shape = fitz_page.new_shape()
        shape.draw_rect(col_rect)
        shape.finish(color=RED, fill=None, width=1.5)
        shape.commit()

    # ── 第四步：里程碑行横向红框 + 气泡 ──
    # 找总额列右边界（用于气泡定位）
    total_col_x1 = col_bounds["surrender"]["x1"] if "surrender" in col_bounds else pw * 0.85

    # 找表格最左边界
    table_left = pw
    for block in text_dict["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            spans = line["spans"]
            if not spans:
                continue
            first_text = spans[0]["text"].strip().replace(",", "")
            if re.match(r"^\d{1,2}$", first_text):
                table_left = min(table_left, spans[0]["bbox"][0] - 2)
    if table_left == pw:
        table_left = 10

    for ms in milestones:
        target_year = str(ms["year"])
        label       = ms["label"]
        r, g, b     = ms["color"]

        for block in text_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans:
                    continue
                first_text = spans[0]["text"].strip().replace(",", "")
                if first_text != target_year:
                    continue

                line_bbox = line["bbox"]
                y0, y1    = line_bbox[1], line_bbox[3]
                row_h     = max(y1 - y0, 8)

                # 横向行框（纯边框，无填充）
                box_rect = fitz.Rect(table_left, y0 - 1.5, total_col_x1, y1 + 1.5)
                shape = fitz_page.new_shape()
                shape.draw_rect(box_rect)
                shape.finish(color=(r, g, b), fill=None, width=1.2)
                shape.commit()

                # 气泡标签
                bubble_w = min(max(len(label) * 7 + 10, 80), 120)
                bx0 = total_col_x1 + 3
                bx1 = bx0 + bubble_w
                if bx1 > pw - 5:
                    bx1 = pw - 5
                    bx0 = bx1 - bubble_w
                bubble = fitz.Rect(bx0, y0 - 2, bx1, y0 - 2 + row_h + 4)

                shape2 = fitz_page.new_shape()
                shape2.draw_rect(bubble)
                shape2.finish(fill=(r, g, b), fill_opacity=0.88,
                             color=(r, g, b), width=0.5)
                shape2.commit()

                kw = dict(fontsize=6.5, color=(1, 1, 1),
                         align=fitz.TEXT_ALIGN_CENTER)
                if font_path:
                    kw["fontfile"] = font_path
                    kw["fontname"] = "cjk"
                fitz_page.insert_textbox(bubble, label, **kw)

    # ── 第五步：底部红字说明 ──
    start_year    = withdrawal_info.get("start_year", 0)
    annual_amount = withdrawal_info.get("annual_amount", 0)
    currency      = withdrawal_info.get("currency", "")
    cur_str = {"美金": "USD", "港幣": "HKD", "人民幣": "RMB"}.get(currency, currency)

    if start_year > 0 and annual_amount > 0:
        left_text = f"第{start_year}年開始，每年提取{int(annual_amount):,}{cur_str}"
    elif start_year > 0:
        left_text = f"第{start_year}年開始提取"
    else:
        left_text = "開始提取後"

    right_text = "提取後保單預計繼續增值"
    label_y    = last_data_y1 + 20

    kw_red = dict(fontsize=11, color=RED)
    if font_path:
        kw_red["fontfile"] = font_path
        kw_red["fontname"] = "cjk"

    fitz_page.insert_text((20, label_y), left_text, **kw_red)
    fitz_page.insert_text((pw * 0.55, label_y), right_text, **kw_red)


# ═══════════════════════════════════════════════════════════════
# 储蓄险主入口
# ═══════════════════════════════════════════════════════════════
def annotate_savings_pdf(input_pdf_path, milestones, font_path=None, log=print):
    if font_path is None:
        font_path = find_chinese_font()

    withdrawal_info = _parse_withdrawal_info(input_pdf_path, log=log)

    doc = fitz.open(input_pdf_path)
    doc = redact_personal_info(doc)

    for page in doc:
        full_text = page.get_text("text")

        is_no_wd   = _is_supplement_no_withdrawal(full_text)
        is_with_wd = _is_supplement_with_withdrawal(full_text)

        if not (is_no_wd or is_with_wd):
            continue

        log(f"  📄 处理页面：no_withdrawal={is_no_wd}  with_withdrawal={is_with_wd}")

        if is_no_wd:
            # 无提取版：横向框整行（原有逻辑不变）
            _annotate_milestone_rows(
                page, milestones, font_path,
                col_header_text="(1)+(2)+(3)",
                fallback_col_ratio=0.52
            )

        elif is_with_wd:
            # 有提取版：竖向列框 + 横向行框 + 底部红字
            _annotate_withdrawal_page(
                page, milestones, withdrawal_info, font_path
            )

    output = io.BytesIO()
    doc.save(output, garbage=4, deflate=True, clean=True)
    doc.close()
    return output.getvalue()
