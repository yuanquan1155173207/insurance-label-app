import pdfplumber, fitz, re, os, io
import pandas as pd
from dataclasses import dataclass

def find_chinese_font():
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
    currency:                  str   = "美金"
    annual_premium:            float = 0
    payment_years:             int   = 0
    coverage_age:              int   = 100
    continuous_cancer_monthly: float = 0

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
        "9 次索償"     in full_text and
        "600%"         in full_text
    )

def _is_cancer_page(full_text):
    return "持續癌症" in full_text and "每月賠償" in full_text

def _draw_red_box(fitz_page, rect, line_width=1.5):
    """在指定区域画红色空心矩形框"""
    shape = fitz_page.new_shape()
    shape.draw_rect(rect)
    shape.finish(
        color=RED,
        fill=None,          # 空心，不填充
        width=line_width
    )
    shape.commit()

def _annotate_summary(fitz_page, words, policy, font_path):
    """
    在说明摘要页：
    1. 写"每年交X / 交X年不用再交"
    2. 写"预计的退保价值" / "预计的理赔金额" 标签
    3. 画红框框住 (1)+(2) 总额列 和 (3)+(4) 总额列的数据区域
    4. 写底部 slogan
    """
    w_paid  = next((w for w in words if "已繳保費" in w["text"]), None)
    w_col12 = next((w for w in words if w["text"] == "(1)+(2)"),  None)
    w_col34 = next((w for w in words if w["text"] == "(3)+(4)"),  None)
    w_100   = next((w for w in words if "100歲"   in w["text"]), None)
    w_note  = next((w for w in words if "上述年齡" in w["text"]), None)

    # ── 1. 写保费/年期提示 ──────────────────────────────────────
    premium_str = f"{int(policy.annual_premium):,}" if policy.annual_premium else "24,170"
    years_str   = str(int(policy.payment_years))    if policy.payment_years  else "10"
    if w_paid:
        by = w_paid["top"]
        _write(fitz_page, f"每年交{premium_str}",     49, by - 12, RED, font_path, fontsize=9)
        _write(fitz_page, f"交{years_str}年不用再交", 49, by - 2,  RED, font_path, fontsize=9)

    # ── 2. 写列标签 ─────────────────────────────────────────────
    if w_col12 and w_100:
        _write(fitz_page, "预计的退保价值",
               w_col12["x0"] - 10, w_100["bottom"] + 12,
               ORANGE, font_path, fontsize=9)

    if w_col34 and w_100:
        _write(fitz_page, "预计的理赔金额",
               w_col34["x0"] - 10, w_100["bottom"] + 12,
               GREEN, font_path, fontsize=9)

    # ── 3. 画红框 ────────────────────────────────────────────────
    # 思路：
    #   用 fitz 在页面里搜索 "(1)+(2)" 和 "(3)+(4)" 的精确坐标，
    #   再搜索表格第一行数字（年度"1"）和最后一行（"100歲"）确定上下边界，
    #   列宽用列标题的 x 范围推算。
    #
    #   框1：(1)+(2) 列 —— 从表头到100歲行
    #   框2：(3)+(4) 列 —— 从表头到100歲行

    # 用 fitz 原生搜索精确坐标
    hits_12  = fitz_page.search_for("(1)+(2)")
    hits_34  = fitz_page.search_for("(3)+(4)")
    hits_100 = fitz_page.search_for("100歲")

    # 表格第一数据行：搜索列标题行下方第一个"1"（年度1）
    # 用 pdfplumber words 里找年度"1"所在行的 top 坐标
    w_year1 = next(
        (w for w in words
         if w["text"] == "1" and w["x0"] < 60),   # 年度列在最左侧
        None
    )

    if hits_12 and hits_34 and hits_100:
        r12  = hits_12[0]
        r34  = hits_34[0]
        r100 = hits_100[-1]   # 取最后一个（表格最后一行）

        # 表格顶部：列标题行顶部（(1)+(2) 标题的 y0）
        # 向上再延伸到包含"退保發還金額"大标题行
        # 找"退保發還金額"
        hits_header = fitz_page.search_for("退保發還金額")
        if hits_header:
            table_top = hits_header[0].y0 - 2
        else:
            table_top = r12.y0 - 18   # 兜底：往上18pt

        # 表格底部：100歲行底部
        table_bottom = r100.y1 + 2

        # ── 框1：(1)+(2) 退保总额列 ──────────────────────────
        # 左边界：(1)+(2) 标题的 x0 再往左留点空间
        # 右边界：(1)+(2) 标题的 x1 再往右留点空间
        # 但实际数据列比标题宽，需要找数据的实际右边界
        # 用 w_100 数据行里 (1)+(2) 列对应数字的位置
        # 简化：以 r12 为中心，左右各扩展到合理范围
        box1_x0 = r12.x0 - 4
        box1_x1 = r12.x1 + 4

        # 找实际数据宽度：搜索100歲行的数字
        # (1)+(2) 列的值是 78,949,000，找它
        hits_val = fitz_page.search_for("78,949,000")
        if hits_val:
            # 找在 r12 列范围附近的那个
            for hv in hits_val:
                if abs(hv.x0 - r12.x0) < 60:
                    box1_x0 = min(box1_x0, hv.x0 - 3)
                    box1_x1 = max(box1_x1, hv.x1 + 3)
                    break

        rect_box1 = fitz.Rect(box1_x0, table_top, box1_x1, table_bottom)
        _draw_red_box(fitz_page, rect_box1, line_width=1.5)

        # ── 框2：(3)+(4) 理赔总额列 ──────────────────────────
        box2_x0 = r34.x0 - 4
        box2_x1 = r34.x1 + 4

        # 同样找100歲行 (3)+(4) 列的值
        # (3)+(4) 列的值也是 78,949,000，找右侧那个
        if hits_val and len(hits_val) >= 2:
            for hv in hits_val:
                if abs(hv.x0 - r34.x0) < 60:
                    box2_x0 = min(box2_x0, hv.x0 - 3)
                    box2_x1 = max(box2_x1, hv.x1 + 3)
                    break
        # 如果只搜到一个，用 r34 偏移推算
        if box2_x1 - box2_x0 < 20:
            col_w   = box1_x1 - box1_x0
            box2_x0 = r34.x0 - 4
            box2_x1 = r34.x0 - 4 + col_w

        rect_box2 = fitz.Rect(box2_x0, table_top, box2_x1, table_bottom)
        _draw_red_box(fitz_page, rect_box2, line_width=1.5)

    # ── 4. 底部 slogan ───────────────────────────────────────────
    slogan_y = (w_note["bottom"] + 20) if w_note else 435
    _write_centered(fitz_page, "有事就赔钱，没事就当存了笔钱",
                    slogan_y, RED, font_path, fontsize=11)


def _annotate_multi(fitz_page, words, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 18) if content_words else 700
    _write_centered(fitz_page, "计划本来还带了多次赔付",
                    last_y, (0.85, 0.30, 0.00), font_path, fontsize=11)

def _annotate_cancer(fitz_page, words, policy, font_path):
    content_words = [w for w in words if w["bottom"] < 760]
    last_y = (max(w["bottom"] for w in content_words) + 18) if content_words else 700

    cur = policy.currency or "美金"
    monthly = int(policy.continuous_cancer_monthly) if policy.continuous_cancer_monthly else 0
    if monthly >= 10000:
        monthly_str = f"{monthly // 10000}W{cur}"
    elif monthly > 0:
        monthly_str = f"{monthly:,}{cur}"
    else:
        monthly_str = f"5W{cur}"

    _write_centered(fitz_page, "及针对大家最担心的癌症，有持续癌症赔付",
                    last_y,      RED, font_path, fontsize=10)
    _write_centered(fitz_page, f"万一得癌症了，一年没康复，每月可赔{monthly_str}",
                    last_y + 14, RED, font_path, fontsize=10)

def _find_text_bbox(page: fitz.Page, search: str):
    hits = page.search_for(search)
    return hits[0] if hits else None

def redact_personal_info(doc: fitz.Document) -> fitz.Document:
    WHITE = (1, 1, 1)

    policy_number = None
    if len(doc) > 0:
        first_page = doc[0]
        text = first_page.get_text("text")
        m = re.search(r"[A-Z]{2}\d{6}-\d{10}-\d", text)
        if m:
            policy_number = m.group(0)

    for page_num in range(len(doc)):
        page = doc[page_num]
        pw   = page.rect.width
        ph   = page.rect.height

        if policy_number:
            hits = page.search_for(policy_number)
            for rect in hits:
                cover = fitz.Rect(rect.x0 - 5, rect.y0 - 2, pw, rect.y1 + 2)
                page.draw_rect(cover, color=WHITE, fill=WHITE)
                if page_num == 0:
                    barcode_cover = fitz.Rect(pw * 0.30, rect.y0 - 32, pw, rect.y0 - 1)
                    page.draw_rect(barcode_cover, color=WHITE, fill=WHITE)
        else:
            rect_fallback = fitz.Rect(pw * 0.30, 0, pw, ph * 0.045)
            page.draw_rect(rect_fallback, color=WHITE, fill=WHITE)
            if page_num == 0:
                rect_bc = fitz.Rect(pw * 0.25, 0, pw, ph * 0.095)
                page.draw_rect(rect_bc, color=WHITE, fill=WHITE)

        hits_name = page.search_for("被保人姓名")
        if hits_name:
            r = hits_name[0]
            footer = fitz.Rect(0, r.y0 - 1, pw * 0.38, ph)
            page.draw_rect(footer, color=WHITE, fill=WHITE)
        else:
            footer_fallback = fitz.Rect(0, ph * 0.960, pw * 0.38, ph)
            page.draw_rect(footer_fallback, color=WHITE, fill=WHITE)

    return doc

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

def extract_fields_from_summary_page(pdf_path):
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words      = page.extract_words()
            full_text  = page.extract_text() or ""

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

            is_summary = _is_summary_page(words_text)
            is_multi   = _is_multi_page(full_text)
            is_cancer  = _is_cancer_page(full_text)

            if not (is_summary or is_multi or is_cancer):
                continue

            print(f"第 {page_idx+1} 页  summary={is_summary}  multi={is_multi}  cancer={is_cancer}")

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

def extract_supplement_table(pdf_path, log=print):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            if "補充說明摘要" not in text and "补充说明摘要" not in text:
                continue
            if any(kw in text for kw in ["悲觀", "樂觀", "悲观", "乐观"]):
                continue
            if "最高貸款額" in text and "65歲" in text:
                continue
            if "解釋附註" in text and "已繳保費" not in text and "(續)" not in text:
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

def annotate_savings_pdf(input_pdf_path, milestones, font_path=None, log=print):
    if font_path is None:
        font_path = find_chinese_font()

    doc = fitz.open(input_pdf_path)
    doc = redact_personal_info(doc)

    for page in doc:
        text_dict  = page.get_text("dict")
        page_width = page.rect.width

        total_col_x1 = None
        for block in text_dict["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    t = span["text"].replace(" ", "")
                    if "(1)+(2)+(3)" in t or "(1)＋(2)＋(3)" in t:
                        total_col_x1 = span["bbox"][2] + 8
                        break

        if total_col_x1 is None:
            total_col_x1 = page_width * 0.52

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
                    if spans[0]["text"].strip().replace(",", "") != target_year:
                        continue

                    line_bbox = line["bbox"]
                    y0, y1    = line_bbox[1], line_bbox[3]
                    row_h     = y1 - y0

                    row_right = total_col_x1
                    for span in spans:
                        if span["bbox"][2] <= total_col_x1 + 5:
                            row_right = max(row_right, span["bbox"][2] + 6)

                    highlight = fitz.Rect(10, y0 - 1, row_right, y1 + 1)
                    shape = page.new_shape()
                    shape.draw_rect(highlight)
                    shape.finish(fill=(r, g, b), fill_opacity=0.15,
                                 color=(r, g, b), width=1.0)
                    shape.commit()

                    bubble_w = 100
                    bubble   = fitz.Rect(row_right + 2, y0 - 2,
                                         row_right + 2 + bubble_w, y0 - 2 + row_h + 4)
                    shape2 = page.new_shape()
                    shape2.draw_rect(bubble)
                    shape2.finish(fill=(r, g, b), fill_opacity=0.88,
                                  color=(r, g, b), width=0.5)
                    shape2.commit()

                    kw = dict(fontsize=6.5, color=(1, 1, 1),
                              align=fitz.TEXT_ALIGN_CENTER)
                    if font_path:
                        kw["fontfile"] = font_path
                        kw["fontname"] = "cjk"
                    page.insert_textbox(bubble, label, **kw)

    output = io.BytesIO()
    doc.save(output, garbage=4, deflate=True, clean=True)
    doc.close()
    return output.getvalue()
