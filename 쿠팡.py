import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import sys, os, io

sys.path.insert(0, os.path.dirname(__file__))
from utils.data import (
    load_sheet, load_sales, load_ad_report,
    load_new_ad_report, build_product_table, build_campaign_table,
    build_detail_by_opt, build_detail_by_camp_opt,
)

st.set_page_config(
    page_title="루시아이 광고 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #f8f6f0; }
div[data-testid="metric-container"] {
    background: white; border-radius: 12px; padding: 16px 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
</style>
""", unsafe_allow_html=True)

# ── 사이드바 ──────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 루시아이\n광고 데이터 대시보드")
    st.divider()

    st.markdown("**📂 광고 리포트** (필수)")
    ad_file  = st.file_uploader("매출최적화 리포트 (.xlsx)", type=["xlsx"], key="ad")
    new_file = st.file_uploader("신규고객 리포트 (.xlsx)",   type=["xlsx"], key="new")

    st.markdown("**📦 판매데이터** (Wing 엑셀)")
    acc_file = st.file_uploader("월 누적 데이터 (.xlsx)", type=["xlsx"], key="acc",
                                help="Wing → 판매지표 → 기간: 이번달 1일~오늘")
    day_file = st.file_uploader("전일 데이터 (.xlsx)",   type=["xlsx"], key="day",
                                help="Wing → 판매지표 → 기간: 어제 하루")

    all_ready = ad_file and new_file
    if all_ready:
        badges = "✓ 광고"
        if acc_file: badges += "  ✓ 누적"
        if day_file: badges += "  ✓ 전일"
        st.success(badges)
        key = "|".join([
            ad_file.name, new_file.name,
            acc_file.name if acc_file else "acc_local",
            day_file.name if day_file else "day_none",
        ])
        if st.session_state.get("file_key") != key:
            st.session_state["ad_bytes"]  = ad_file.read()
            st.session_state["new_bytes"] = new_file.read()
            st.session_state["acc_bytes"] = acc_file.read() if acc_file else None
            st.session_state["day_bytes"] = day_file.read() if day_file else None
            st.session_state["file_key"]  = key
    else:
        missing = []
        if not ad_file:  missing.append("광고보고서_매출.xlsx")
        if not new_file: missing.append("신규광고.xlsx")
        st.info("필수: " + ", ".join(missing))
    if not acc_file:
        st.caption("누적 미업로드 시 판매데이터 없이 실행")

    st.divider()
    page = st.radio("페이지", ["광고 대시보드", "광고현황 (상품별)"], label_visibility="collapsed")

# ── 데이터 로드 ───────────────────────────────────────
if "ad_bytes" not in st.session_state:
    st.title("📊 루시아이 광고 대시보드")
    st.info("← 왼쪽에서 광고 리포트 2개를 업로드하세요.")
    st.stop()

@st.cache_data(ttl=300, show_spinner=False)  # 마진 시트 5분 캐시
def get_sheet():
    return load_sheet()

@st.cache_data(show_spinner="데이터 처리 중...")
def process(ad_bytes, new_bytes, acc_bytes=None, day_bytes=None):
    opt_mw, opt_mr, opt_reg, opt_type = get_sheet()
    opt2prod, prod2name, acc_rev, acc_qty, day_rev, day_qty, opt_rev_map, opt_qty_map = \
        load_sales(acc_bytes, day_bytes)
    ad_df    = load_ad_report(io.BytesIO(ad_bytes))
    new_df   = load_new_ad_report(io.BytesIO(new_bytes))
    prod_tbl   = build_product_table(ad_df, new_df, opt2prod, prod2name, opt_mw,
                                     acc_rev, acc_qty, day_rev, day_qty)
    camp_tbl   = build_campaign_table(ad_df, opt2prod, prod2name, opt_mw)
    detail_opt = build_detail_by_opt(ad_df, new_df, opt2prod, prod2name, opt_mw, opt_rev_map, opt_qty_map, opt_type)
    detail_co  = build_detail_by_camp_opt(ad_df, new_df, opt2prod, prod2name, opt_mw, opt_rev_map, opt_qty_map, opt_type)
    return prod_tbl, camp_tbl, detail_opt, detail_co

prod_tbl, camp_tbl, detail_opt, detail_co = process(
    st.session_state["ad_bytes"],
    st.session_state["new_bytes"],
    st.session_state.get("acc_bytes"),
    st.session_state.get("day_bytes"),
)

# ═══════════════════════════════════════════════════
# 페이지 1: 광고 대시보드
# ═══════════════════════════════════════════════════
if page == "광고 대시보드":
    st.title("📢 광고 대시보드")
    st.caption("캠페인 단위 수익 성과 분석")

    # KPI
    has_ad = prod_tbl[prod_tbl["누적_광고합계"] > 0]
    총광고  = prod_tbl["누적_광고합계"].sum()
    총매출  = prod_tbl["누적_월매출"].sum()
    총마진  = prod_tbl["누적_마진금"].sum()
    총순이익= (prod_tbl["누적_마진금"] - prod_tbl["누적_광고합계"]).sum()
    마진사용= 총광고 / 총마진 * 100 if 총마진 > 0 else 0
    로아스  = 총매출 / 총광고 * 100 if 총광고 > 0 else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("총 광고비",    f"₩{총광고:,.0f}")
    c2.metric("총 월매출",    f"₩{총매출:,.0f}")
    c3.metric("직접 로아스",  f"{로아스:.0f}%")
    c4.metric("마진금 사용%", f"{마진사용:.1f}%",
              delta="위험" if 마진사용 > 80 else "양호",
              delta_color="inverse" if 마진사용 > 80 else "normal")
    c5.metric("추정 순이익",  f"₩{총순이익:,.0f}")

    st.divider()

    col_l, col_r = st.columns(2)

    with col_l:
        st.subheader("캠페인별 순이익 (상위 20)")
        top = camp_tbl.head(20).copy()
        top["색"] = top["순이익"].apply(lambda x: "흑자" if x >= 0 else "적자")
        fig = px.bar(
            top, x="순이익", y="campaign", orientation="h",
            color="색",
            color_discrete_map={"흑자": "#4CAF50", "적자": "#F44336"},
            labels={"순이익": "순이익(원)", "campaign": ""},
            height=500,
        )
        fig.update_layout(showlegend=False, yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        st.subheader("광고비 vs 마진금사용%")
        valid = camp_tbl[camp_tbl["마진금사용%"].notna() & (camp_tbl["직접마진"] > 0)]
        fig2 = px.scatter(
            valid, x="광고비", y="마진금사용%",
            hover_name="campaign",
            color="순이익",
            color_continuous_scale=["#F44336", "#FFC107", "#4CAF50"],
            size="광고비", size_max=25,
            labels={"광고비": "광고비(원)", "마진금사용%": "마진금사용(%)"},
            height=500,
        )
        fig2.add_hline(y=80,  line_dash="dash", line_color="#FF9800", annotation_text="80%")
        fig2.add_hline(y=100, line_dash="dash", line_color="#F44336", annotation_text="100%")
        st.plotly_chart(fig2, use_container_width=True)

    tab_camp, tab_by_opt, tab_by_co = st.tabs([
        "📊 캠페인별 집계",
        "📦 옵션ID별 합산",
        "🔍 캠페인 × 옵션ID",
    ])

    def color_neg(v):
        return "color: #F44336" if not pd.isna(v) and v < 0 else "color: #4CAF50" if not pd.isna(v) else ""
    def bg_mg(v):
        if pd.isna(v): return ""
        if v > 100: return "background-color: #ffebee"
        if v > 80:  return "background-color: #fff8e1"
        return ""

    money_fmt = {
        "매출광고비":"₩{:,.0f}","신규광고비":"₩{:,.0f}","광고합계":"₩{:,.0f}",
        "월매출":"₩{:,.0f}","직접마진":"₩{:,.0f}","순이익":"₩{:,.0f}",
        "마진금사용%":"{:.1f}%","직접로아스%":"{:.0f}%",
        "노출수":"{:,.0f}","클릭수":"{:,.0f}",
    }

    # ── 탭 1: 캠페인별 집계 ─────────────────────────
    with tab_camp:
        disp = camp_tbl[["campaign","광고비","직접마진","순이익","직접로아스%","마진금사용%"]].copy()
        disp.columns = ["캠페인명","광고비","직접마진","순이익","직접로아스%","마진금사용%"]
        st.dataframe(
            disp.style
            .format({"광고비":"₩{:,.0f}","직접마진":"₩{:,.0f}","순이익":"₩{:,.0f}",
                     "직접로아스%":"{:.0f}%","마진금사용%":"{:.1f}%"}, na_rep="-")
            .map(color_neg, subset=["순이익"])
            .map(bg_mg,     subset=["마진금사용%"]),
            use_container_width=True, height=500
        )

    # ── 탭 2: 옵션ID별 합산 ─────────────────────────
    with tab_by_opt:
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            keyword = st.text_input("옵션명 검색", placeholder="예: 선풍기", key="opt_keyword")
        with col_f2:
            profit_f = st.selectbox("수익 필터", ["전체","흑자만","적자만"], key="opt_profit")

        dopt = detail_opt[detail_opt["광고합계"] > 0].copy()
        if keyword:
            dopt = dopt[dopt["옵션명"].str.contains(keyword, na=False)]
        if profit_f == "흑자만":   dopt = dopt[dopt["순이익"] >= 0]
        elif profit_f == "적자만": dopt = dopt[dopt["순이익"] < 0]

        st.caption(f"총 {len(dopt):,}개 옵션")
        st.dataframe(
            dopt.style.format(money_fmt, na_rep="-")
            .map(color_neg, subset=["순이익"])
            .map(bg_mg,     subset=["마진금사용%"]),
            use_container_width=True, height=600,
            column_config={
                "옵션ID":    st.column_config.TextColumn(width="small"),
                "판매방식":  st.column_config.TextColumn(width="small"),
                "옵션명":    st.column_config.TextColumn(width="large"),
            }
        )

    # ── 탭 3: 캠페인 × 옵션ID ───────────────────────
    with tab_by_co:
        col_g1, col_g2, col_g3 = st.columns([2, 2, 1])
        with col_g1:
            camp_sel = st.multiselect(
                "캠페인 필터", options=sorted(detail_co["캠페인명"].unique()),
                placeholder="전체", key="co_camp"
            )
        with col_g2:
            keyword2 = st.text_input("옵션명 검색", placeholder="예: 정리함", key="co_keyword")
        with col_g3:
            profit_g = st.selectbox("수익 필터", ["전체","흑자만","적자만"], key="co_profit")

        dco = detail_co[detail_co["광고합계"] > 0].copy()
        if camp_sel:    dco = dco[dco["캠페인명"].isin(camp_sel)]
        if keyword2:    dco = dco[dco["옵션명"].str.contains(keyword2, na=False)]
        if profit_g == "흑자만":   dco = dco[dco["순이익"] >= 0]
        elif profit_g == "적자만": dco = dco[dco["순이익"] < 0]

        st.caption(f"총 {len(dco):,}행")
        st.dataframe(
            dco.style.format(money_fmt, na_rep="-")
            .map(color_neg, subset=["순이익"])
            .map(bg_mg,     subset=["마진금사용%"]),
            use_container_width=True, height=600,
            column_config={
                "캠페인ID":  st.column_config.TextColumn(width="small"),
                "옵션ID":    st.column_config.TextColumn(width="small"),
                "판매방식":  st.column_config.TextColumn(width="small"),
                "캠페인명":  st.column_config.TextColumn(width="medium"),
                "옵션명":    st.column_config.TextColumn(width="large"),
            }
        )

# ═══════════════════════════════════════════════════
# 페이지 2: 광고현황 (상품별)
# ═══════════════════════════════════════════════════
else:
    st.title("📋 광고현황")
    st.caption("상품(등록상품ID) 단위 — 운영시트 광고현황 탭 포맷")

    has_ad = prod_tbl[prod_tbl["누적_광고합계"] > 0].copy()

    총광고  = has_ad["누적_광고합계"].sum()
    총매출  = has_ad["누적_월매출"].sum()
    총마진  = has_ad["누적_마진금"].sum()
    마진사용= 총광고 / 총마진 * 100 if 총마진 > 0 else 0
    로아스  = 총매출 / 총광고 * 100 if 총광고 > 0 else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("광고 집행 상품수", f"{len(has_ad)}개")
    c2.metric("마진금 사용%",    f"{마진사용:.1f}%")
    c3.metric("직접 로아스",     f"{로아스:.0f}%")
    c4.metric("추정 순이익",     f"₩{(총마진-총광고):,.0f}")

    st.divider()

    col_l, col_r = st.columns([3,2])
    with col_l:
        st.subheader("상품별 마진금사용% (상위 25)")
        chart_df = has_ad.dropna(subset=["누적_마진금사용%"]).sort_values("누적_마진금사용%", ascending=False).head(25)
        chart_df["구간"] = chart_df["누적_마진금사용%"].apply(
            lambda x: "위험(>100%)" if x > 100 else ("주의(80~100%)" if x > 80 else "양호(<80%)")
        )
        fig = px.bar(
            chart_df, x="상품명", y="누적_마진금사용%", color="구간",
            color_discrete_map={"위험(>100%)":"#F44336","주의(80~100%)":"#FF9800","양호(<80%)":"#4CAF50"},
            labels={"상품명":"","누적_마진금사용%":"마진금사용(%)"},
            height=400,
        )
        fig.add_hline(y=80, line_dash="dash", line_color="#FF9800")
        fig.add_hline(y=100, line_dash="dash", line_color="#F44336")
        fig.update_xaxes(tickangle=45, tickfont_size=9)
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        st.subheader("직접로아스 vs 마진금사용%")
        v2 = has_ad.dropna(subset=["누적_마진금사용%","누적_직접로아스%"])
        fig2 = px.scatter(
            v2, x="누적_직접로아스%", y="누적_마진금사용%",
            hover_name="상품명", color="누적_순이익",
            color_continuous_scale=["#F44336","#FFC107","#4CAF50"],
            height=400,
            labels={"누적_직접로아스%":"직접로아스(%)","누적_마진금사용%":"마진금사용(%)"},
        )
        fig2.add_hline(y=80, line_dash="dash", line_color="orange")
        st.plotly_chart(fig2, use_container_width=True)

    tab1, tab2 = st.tabs(["📅 누적 (월간)", "📆 전일"])

    with tab1:
        acc = has_ad[["상품명","누적_월매출","누적_마진금","누적_매출광고","누적_신규광고",
                      "누적_광고합계","누적_마진금사용%","누적_직접로아스%","누적_순이익"]].copy()
        acc.columns = ["상품명","월매출","마진금","매출광고비","신규광고비",
                       "광고비합계","마진금사용%","직접로아스%","순이익"]
        def bg_mg(v):
            if pd.isna(v): return ""
            if v > 100: return "background-color: #ffebee"
            if v > 80:  return "background-color: #fff8e1"
            return ""
        def c_neg(v):
            return "color: #F44336" if not pd.isna(v) and v < 0 else ""
        st.dataframe(
            acc.style
            .format({"월매출":"₩{:,.0f}","마진금":"₩{:,.0f}","매출광고비":"₩{:,.0f}",
                     "신규광고비":"₩{:,.0f}","광고비합계":"₩{:,.0f}",
                     "마진금사용%":"{:.1f}%","직접로아스%":"{:.0f}%","순이익":"₩{:,.0f}"}, na_rep="-")
            .map(bg_mg, subset=["마진금사용%"])
            .map(c_neg, subset=["순이익"]),
            use_container_width=True, height=600
        )

    with tab2:
        day = prod_tbl[prod_tbl["전일_매출"] > 0][["상품명","전일_매출","전일_마진금"]].copy()
        day.columns = ["상품명","전일매출","전일마진금"]
        st.dataframe(
            day.style.format({"전일매출":"₩{:,.0f}","전일마진금":"₩{:,.0f}"}),
            use_container_width=True, height=400
        )
        st.caption("⚠️ 전일 광고비는 일별 광고 리포트 별도 업로드 필요")

    st.divider()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        has_ad.to_excel(w, sheet_name="광고현황_누적", index=False)
    buf.seek(0)
    st.download_button("📥 Excel 다운로드", buf, "광고현황.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
