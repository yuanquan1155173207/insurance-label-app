
import streamlit as st
import fitz, tempfile, os, re, pandas as pd
from core import (
    extract_text, extract_fields, extract_fields_from_summary_page,
    find_chinese_font,
    CriticalIllnessPolicy, annotate_critical_illness_pdf,
    extract_supplement_table, find_key_milestones, annotate_savings_pdf,
)

st.set_page_config(page_title="保险建议书标注工具", page_icon="🏥", layout="wide")
st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:700;color:#c0392b;text-align:center;margin-bottom:.2rem}
.sub-title{text-align:center;color:#666;margin-bottom:2rem;font-size:.95rem}
.success-box{background:#eafaf1;border-radius:8px;padding:.8rem 1.2rem;border-left:4px solid #27ae60}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🏥 保险建议书自动标注工具</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">上传 PDF → 自动识别关键信息 → 生成标注版 PDF</div>', unsafe_allow_html=True)
st.divider()

with st.sidebar:
    st.header("⚙️ 全局设置")
    auto_font = find_chinese_font()
    if auto_font:
        st.success("✅ 已检测到中文字体")
        st.caption(auto_font)
        use_auto  = st.checkbox("使用自动检测字体", value=True)
        font_path = auto_font if use_auto else (st.text_input("手动输入字体路径") or None)
    else:
        st.warning("⚠️ 未检测到中文字体")
        font_path = st.text_input("手动输入字体路径",
                                   placeholder="C:/Windows/Fonts/msyh.ttc") or None
    st.divider()
    st.subheader("🖼️ 预览设置")
    preview_dpi   = st.slider("预览分辨率 DPI", 72, 200, 130, 13)
    preview_pages = st.multiselect("预览页码", list(range(1,31)),
                                   default=[1,2,3],
                                   format_func=lambda x: f"第 {x} 页")

tab_ci, tab_sv = st.tabs(["🏥 重疾险标注", "💰 储蓄险标注"])

# ════════════════════════════════════════════════════════════════
# TAB 1：重疾险
# ════════════════════════════════════════════════════════════════
with tab_ci:
    col1, col2 = st.columns([1,1], gap="large")
    with col1:
        st.subheader("📤 上传重疾险建议书")
        uploaded_ci = st.file_uploader(
            "支持安盛、友邦、宏利等主流重疾险建议书",
            type=["pdf"], key="ci_uploader")
        if uploaded_ci:
            st.markdown(
                f'<div class="success-box">✅ 已上传：<b>{uploaded_ci.name}</b>'
                f'　📦 {uploaded_ci.size/1024:.1f} KB</div>',
                unsafe_allow_html=True)
    with col2:
        st.subheader("📌 标注效果预览说明")
        st.markdown("""
| 页面 | 标注内容 |
|------|---------|
| 📋 说明摘要页 | 🔴 缴费信息（表头）<br>🟠 预计的退保价值<br>🟢 预计的理赔金额<br>🔴 有事就赔钱，没事就当存了笔钱 |
| 🔁 多重赔付页 | 🟠 计划本来还带了多次赔付 |
| 🎗️ 持续癌症页 | 🔴 每月可赔金额说明 |
        """, unsafe_allow_html=True)

    if uploaded_ci:
        st.divider()
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path  = os.path.join(tmpdir, uploaded_ci.name)
            output_path = os.path.join(tmpdir,
                uploaded_ci.name.replace(".pdf","_重疾标注版.pdf"))
            with open(input_path,"wb") as f: f.write(uploaded_ci.getvalue())

            progress = st.progress(0, text="正在读取 PDF...")
            try:
                progress.progress(20, text="提取文字字段...")
                text   = extract_text(input_path)
                fields = extract_fields(text)
                # 优先用精确提取（直接从说明摘要页读数据）
                precise = extract_fields_from_summary_page(input_path)
                for k, v in precise.items():
                    if v:
                        fields[k] = v

                # 自动推断持续癌症月赔
                cancer_monthly = float(fields.get("continuous_cancer_monthly", 0))
                if cancer_monthly == 0:
                    m = re.search(r"基本保額[^\d]*(\d[\d,]+)", text)
                    if m:
                        base_sum = int(m.group(1).replace(",",""))
                        cancer_monthly = base_sum * 0.05

                policy = CriticalIllnessPolicy(
                    insured_name    = str(fields.get("insured_name","未识别")),
                    insured_age     = int(fields.get("insured_age", 0)),
                    applicant_name  = str(fields.get("applicant_name","未识别")),
                    currency        = str(fields.get("currency","美金")),
                    annual_premium  = float(fields.get("annual_premium", 0)),
                    payment_years   = int(fields.get("payment_years", 0)),
                    coverage_age    = int(fields.get("coverage_age", 100)),
                    continuous_cancer_monthly = cancer_monthly,
                )

                progress.progress(35, text="显示识别结果...")
                st.subheader("📊 自动识别结果")
                c1,c2,c3,c4 = st.columns(4)
                c1.metric("被保人",   policy.insured_name or "未识别")
                c2.metric("年保费",
                    f"{policy.currency} {int(policy.annual_premium):,}"
                    if policy.annual_premium else "未识别")
                c3.metric("缴费年限",
                    f"{policy.payment_years} 年"
                    if policy.payment_years else "未识别")
                c4.metric("癌症月赔",
                    f"{policy.currency} {int(policy.continuous_cancer_monthly):,}"
                    if policy.continuous_cancer_monthly else "未识别")

                # 手动修正区
                with st.expander("✏️ 手动修正识别字段（识别有误时使用）"):
                    col_a, col_b, col_c, col_d = st.columns(4)
                    new_premium = col_a.number_input(
                        "年保费", value=float(policy.annual_premium), min_value=0.0, step=100.0)
                    new_years   = col_b.number_input(
                        "缴费年限", value=int(policy.payment_years), min_value=0, step=1)
                    new_monthly = col_c.number_input(
                        "癌症月赔", value=float(policy.continuous_cancer_monthly),
                        min_value=0.0, step=1000.0)
                    new_currency = col_d.selectbox(
                        "货币", ["美金","港币","人民币"],
                        index=["美金","港币","人民币"].index(policy.currency)
                              if policy.currency in ["美金","港币","人民币"] else 0)
                    if st.button("✅ 应用修正"):
                        policy.annual_premium = new_premium
                        policy.payment_years  = new_years
                        policy.continuous_cancer_monthly = new_monthly
                        policy.currency = new_currency
                        st.success("已更新字段，将使用修正后的值生成标注")

                progress.progress(60, text="生成标注 PDF...")
                ci_pdf_bytes = annotate_critical_illness_pdf(
                    input_path, policy, font_path=font_path)
                with open(output_path,"wb") as f: f.write(ci_pdf_bytes)
                progress.progress(100, text="✅ 完成！")

                st.divider()
                st.subheader("🖼️ 预览标注效果")
                doc = fitz.open(output_path)
                pages_to_show = [p-1 for p in preview_pages if 0 < p <= len(doc)] or [0]
                if pages_to_show:
                    tabs_p = st.tabs([f"第 {p+1} 页" for p in pages_to_show])
                    for tab_p, p_idx in zip(tabs_p, pages_to_show):
                        with tab_p:
                            mat = fitz.Matrix(preview_dpi/72, preview_dpi/72)
                            pix = doc[p_idx].get_pixmap(matrix=mat, alpha=False)
                            st.image(pix.tobytes("png"), use_container_width=True)
                doc.close()

                st.divider()
                st.download_button(
                    "📥 下载重疾险标注版 PDF", data=ci_pdf_bytes,
                    file_name=uploaded_ci.name.replace(".pdf","_重疾标注版.pdf"),
                    mime="application/pdf",
                    use_container_width=True, type="primary")

            except Exception as e:
                progress.empty()
                st.error(f"❌ 处理失败：{e}")
                st.exception(e)

# ════════════════════════════════════════════════════════════════
# TAB 2：储蓄险
# ════════════════════════════════════════════════════════════════
with tab_sv:
    col1, col2 = st.columns([1,1], gap="large")
    with col1:
        st.subheader("📤 上传储蓄险建议书")
        uploaded_sv = st.file_uploader(
            "支持盛利II等储蓄类建议书 PDF",
            type=["pdf"], key="sv_uploader")
        if uploaded_sv:
            st.markdown(
                f'<div class="success-box">✅ 已上传：<b>{uploaded_sv.name}</b>'
                f'　📦 {uploaded_sv.size/1024:.1f} KB</div>',
                unsafe_allow_html=True)
    with col2:
        st.subheader("ℹ️ 标注内容")
        st.markdown("自动检测 **保本 / 翻倍 / 再翻倍** 年度并整行高亮 + 气泡标签")

    if uploaded_sv:
        st.divider()
        with tempfile.TemporaryDirectory() as tmpdir:
            sv_input  = os.path.join(tmpdir, uploaded_sv.name)
            sv_output = os.path.join(tmpdir,
                uploaded_sv.name.replace(".pdf","_储蓄标注版.pdf"))
            with open(sv_input,"wb") as f: f.write(uploaded_sv.getvalue())

            log_msgs = []
            def sv_log(msg): log_msgs.append(msg)

            progress_sv = st.progress(0, text="正在解析文本数据...")
            try:
                progress_sv.progress(30, text="提取年度数据...")
                df = extract_supplement_table(sv_input, log=sv_log)
                for msg in log_msgs: st.caption(msg)

                if df.empty:
                    progress_sv.empty()
                    st.error("❌ 未能识别到有效数据，请确认 PDF 包含「補充說明摘要」页")
                else:
                    progress_sv.progress(55, text="检测关键节点...")
                    milestones = find_key_milestones(df, log=sv_log)
                    st.success(f"✅ 成功解析 {len(df)} 年数据，检测到 {len(milestones)} 个关键节点")

                    st.subheader("📍 关键节点")
                    key_labels = {0:"🟢 保本", 1:"🟡 翻倍 (2x)", 2:"🟠 再翻倍 (4x)"}
                    if milestones:
                        cols = st.columns(len(milestones))
                        for i, (col, ms) in enumerate(zip(cols, milestones)):
                            sv_val = df[df["year"]==ms["year"]]["surrender_total"].values
                            sv_str = f"退保 ${int(sv_val[0]):,}" if len(sv_val) else ""
                            col.metric(key_labels.get(i,f"节点{i+1}"),
                                       f"第 {ms['year']} 年", delta=sv_str)

                    st.subheader("📊 数据预览（前20行）")
                    preview_rows = [
                        {"年度": str(r["year"]),
                         "已缴保费": f"{r['paid_total']:,}",
                         "退保总额": f"{r['surrender_total']:,}",
                         "身故赔付": f"{r['death_total']:,}"}
                        for r in df.head(20).to_dict("records")
                    ]
                    st.dataframe(pd.DataFrame(preview_rows),
                                 use_container_width=True, hide_index=True)

                    st.divider()
                    progress_sv.progress(75, text="生成标注 PDF...")
                    sv_pdf_bytes = annotate_savings_pdf(
                        sv_input, milestones, font_path=font_path, log=sv_log)
                    with open(sv_output,"wb") as f: f.write(sv_pdf_bytes)
                    progress_sv.progress(100, text="✅ 完成！")

                    st.subheader("🖼️ 预览标注效果")
                    doc_sv = fitz.open(sv_output)
                    sv_pages = [p-1 for p in preview_pages if 0 < p <= len(doc_sv)] or [0]
                    sv_tabs  = st.tabs([f"第 {p+1} 页" for p in sv_pages])
                    for tab_p, p_idx in zip(sv_tabs, sv_pages):
                        with tab_p:
                            mat = fitz.Matrix(preview_dpi/72, preview_dpi/72)
                            pix = doc_sv[p_idx].get_pixmap(matrix=mat, alpha=False)
                            st.image(pix.tobytes("png"), use_container_width=True)
                    doc_sv.close()

                    st.divider()
                    st.download_button(
                        "📥 下载储蓄险标注版 PDF", data=sv_pdf_bytes,
                        file_name=uploaded_sv.name.replace(".pdf","_储蓄标注版.pdf"),
                        mime="application/pdf",
                        use_container_width=True, type="primary")

            except Exception as e:
                progress_sv.empty()
                st.error(f"❌ 处理失败：{e}")
                st.exception(e)

st.divider()
st.caption("保险建议书自动标注工具 · 仅供参考，不构成投资建议")
