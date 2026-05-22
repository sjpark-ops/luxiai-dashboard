import pandas as pd
import os

def num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "").str.replace("%", ""),
        errors="coerce"
    ).fillna(0)

BASE_AD = os.path.join(os.path.dirname(__file__), "..", "..", "쿠팡 광고")
def load_sheet(sheet_id=""):
    import io as _io
    import urllib.request as _ureq

    def _fetch(gid):
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        req = _ureq.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with _ureq.urlopen(req, timeout=15) as r:
            return r.read()

    opt_mw, opt_mr, opt_reg, opt_type = {}, {}, {}, {}

    # 로켓그로스: 상품관리(현황판) — iloc[:,1]=등록ID, [:,2]=옵션ID, [:,18]=마진금, [:,19]=마진율
    try:
        if not sheet_id:
            raise ValueError("no sheet_id")
        raw = _fetch("795587558")
        df = pd.read_csv(_io.BytesIO(raw), encoding="utf-8", header=3, dtype=str)
        opt = df.iloc[:, 2].astype(str).str.strip()
        reg = df.iloc[:, 1].astype(str).str.strip()
        mw  = num(df.iloc[:, 18])
        mr  = num(df.iloc[:, 19]) / 100
        ok  = opt.str.match(r"^\d{6,}$")
        opt_mw.update(zip(opt[ok], mw[ok]))
        opt_mr.update(zip(opt[ok], mr[ok]))
        opt_reg.update(zip(opt[ok], reg[ok]))
        opt_type.update({o: "그로스" for o in opt[ok]})
    except Exception as e:
        raise RuntimeError(f"현황판 로드 실패: {e}") from e

    # 판매자배송 — iloc[:,1]=옵션ID, [:,15]=현재마진율%, [:,16]=현재마진금
    try:
        if not sheet_id:
            raise ValueError("no sheet_id")
        df = pd.read_csv(_io.BytesIO(_fetch("1484224058")), encoding="utf-8", header=3, dtype=str)
        opt = df.iloc[:, 1].astype(str).str.strip()
        mw  = num(df.iloc[:, 16])
        mr  = num(df.iloc[:, 15]) / 100
        ok  = opt.str.match(r"^\d{6,}$")
        for o, w, rv in zip(opt[ok], mw[ok], mr[ok]):
            if o not in opt_mw:
                opt_mw[o] = w
                opt_mr[o] = rv
                opt_reg[o] = o
            opt_type[o] = "판매자배송"  # 판매자배송 탭에 있으면 무조건 판매자배송
    except Exception as e:
        raise RuntimeError(f"판매자배송 로드 실패: {e}") from e

    # 폴백: 로컬 CSV (개발환경, sheet_id 없을 때)
    if not opt_mw and not sheet_id:
        path = os.path.join(BASE_AD, "_sheet.csv")
        df = pd.read_csv(path, encoding="utf-8-sig", header=3, dtype=str)
        opt  = df.iloc[:, 3].astype(str).str.strip()
        reg  = df.iloc[:, 2].astype(str).str.strip()
        typ  = df.iloc[:, 1].astype(str).str.strip()
        mw   = num(df.iloc[:, 18])
        mr   = num(df.iloc[:, 19]) / 100
        ok   = opt.str.match(r"^\d{6,}$")
        opt_mw   = dict(zip(opt[ok], mw[ok]))
        opt_mr   = dict(zip(opt[ok], mr[ok]))
        opt_reg  = dict(zip(opt[ok], reg[ok]))
        opt_type = dict(zip(opt[ok], typ[ok]))

    return opt_mw, opt_mr, opt_reg, opt_type

def _parse_wing_excel(data):
    """Wing 엑셀 1개 → (opt2prod, prod2name, rev_by_prod, qty_by_prod, opt_rev, opt_qty)"""
    import io as _io
    df = pd.read_excel(_io.BytesIO(data), header=0, dtype=str)
    col_map = {c.strip(): i for i, c in enumerate(df.columns)}
    opt  = df.iloc[:, col_map.get("옵션 ID", 0)].astype(str).str.strip()
    name = df.iloc[:, col_map.get("상품명", 2)].astype(str).str.strip()
    prod = df.iloc[:, col_map.get("등록상품ID", 3)].astype(str).str.strip()
    rev  = num(df.iloc[:, col_map.get("매출(원)", 6)])
    qty  = num(df.iloc[:, col_map.get("판매량", 8)])
    valid = opt.str.match(r"^\d{6,}$") & prod.str.match(r"^\d{6,}$")
    opt2prod  = dict(zip(opt[valid], prod[valid]))
    prod2name = dict(zip(prod[valid], name[valid]))
    # 옵션 단위 매출/수량: 동일 옵션ID 여러 행 있을 수 있으므로 groupby sum
    _od = pd.DataFrame({"opt": opt[valid].values, "rev": rev[valid].values, "qty": qty[valid].values})
    opt_rev = _od.groupby("opt")["rev"].sum().to_dict()
    opt_qty = _od.groupby("opt")["qty"].sum().to_dict()
    def agg(pid_col, val_col):
        d = pd.DataFrame({"pid": pid_col, "v": val_col})
        d = d[d["pid"].str.match(r"^\d{6,}$", na=False)]
        return d.groupby("pid")["v"].sum()
    return opt2prod, prod2name, agg(prod, rev), agg(prod, qty), opt_rev, opt_qty


def load_sales(acc_bytes=None, day_bytes=None):
    """
    acc_bytes: 월 누적 Wing 엑셀 bytes (없으면 로컬 _sales.csv 폴백)
    day_bytes: 전일 Wing 엑셀 bytes (없으면 빈 Series)
    """
    empty = pd.Series(dtype=float)

    opt_rev_map, opt_qty_map = {}, {}
    if acc_bytes is not None:
        opt2prod, prod2name, acc_rev, acc_qty, opt_rev_map, opt_qty_map = _parse_wing_excel(acc_bytes)
    else:
        opt2prod, prod2name = {}, {}
        acc_rev = empty
        acc_qty = empty
        # 로컬 CSV 폴백 (개발환경 전용)
        try:
            import io as _io
            path = os.path.join(BASE_AD, "_sales.csv")
            df = pd.read_csv(path, encoding="utf-8-sig", header=1, dtype=str)
            opt  = df.iloc[:, 1].astype(str).str.strip()
            name = df.iloc[:, 3].astype(str).str.strip()
            prod = df.iloc[:, 4].astype(str).str.strip()
            rev  = num(df.iloc[:, 7])
            qty  = num(df.iloc[:, 9])
            valid = opt.str.match(r"^\d{6,}$") & prod.str.match(r"^\d{6,}$")
            opt2prod  = dict(zip(opt[valid], prod[valid]))
            prod2name = dict(zip(prod[valid], name[valid]))
            def agg(pid_col, val_col):
                d = pd.DataFrame({"pid": pid_col, "v": val_col})
                d = d[d["pid"].str.match(r"^\d{6,}$", na=False)]
                return d.groupby("pid")["v"].sum()
            acc_rev = agg(prod, rev)
            acc_qty = agg(prod, qty)
        except Exception:
            pass

    if day_bytes is not None:
        _, _, day_rev, day_qty, _, _ = _parse_wing_excel(day_bytes)
    else:
        day_rev = empty
        day_qty = empty

    return opt2prod, prod2name, acc_rev, acc_qty, day_rev, day_qty, opt_rev_map, opt_qty_map

def load_ad_report(file_obj):
    df = pd.read_excel(file_obj, sheet_name="Sheet1", header=0, dtype=str)
    return pd.DataFrame({
        "campaign_id": df["캠페인 ID"].astype(str).str.strip(),
        "campaign":    df["캠페인명"].astype(str).str.strip(),
        "ad_group":    df["광고그룹"].astype(str).str.strip(),
        "opt":         df["광고집행 옵션ID"].astype(str).str.strip(),
        "prod_name_ad":df["광고집행 상품명"].astype(str).str.strip(),
        "impression":  num(df["노출수"]),
        "click":       num(df["클릭수"]),
        "spend":       num(df["광고비"]),
        "d1_rev":      num(df["직접 전환매출액(1일)"]),
        "d14_rev":     num(df["직접 전환매출액(14일)"]),
        "d1_qty":      num(df["직접 판매수량(1일)"]),
        "d14_qty":     num(df["직접 판매수량(14일)"]),
    })

def load_new_ad_report(file_obj):
    df = pd.read_excel(file_obj, sheet_name=0, header=0, dtype=str)
    opt         = df["광고집행 옵션 ID"].astype(str).str.strip()
    spend       = num(df["집행 광고비"])
    campaign    = df["캠페인 이름"].astype(str).str.strip()
    campaign_id = df["캠페인 ID"].astype(str).str.strip()
    return pd.DataFrame({"opt": opt, "campaign_id": campaign_id, "campaign": campaign, "spend": spend})

def build_product_table(ad_df, new_df, opt2prod, prod2name, opt_mw, acc_rev, acc_qty, day_rev, day_qty):
    """상품(등록상품ID) 단위 광고현황 테이블 생성"""
    # 매출광고비
    ad_df["pid"] = ad_df["opt"].map(opt2prod)
    new_df["pid"] = new_df["opt"].map(opt2prod)

    ad_by_prod  = ad_df.dropna(subset=["pid"]).groupby("pid")["spend"].sum()
    new_by_prod = new_df.dropna(subset=["pid"]).groupby("pid")["spend"].sum()

    # 마진금 합산 (마진금_개당 × 누적판매량)
    def margin_sum(qty_series_by_prod):
        result = {}
        for pid, qty in qty_series_by_prod.items():
            # 이 pid에 속한 옵션들의 마진금 가중합산
            opts = [o for o, p in opt2prod.items() if p == pid and o in opt_mw]
            # 판매량을 옵션별로 분배할 수 없어서 단순 평균 마진율 사용
            if opts:
                avg_mw = sum(opt_mw[o] for o in opts) / len(opts)
                result[pid] = avg_mw * qty
        return pd.Series(result)

    acc_margin = margin_sum(acc_qty)
    day_margin_s = margin_sum(day_qty)

    all_pids = sorted(
        set(acc_rev.index) | set(ad_by_prod.index) | set(new_by_prod.index),
        key=lambda p: prod2name.get(p, p)
    )

    rows = []
    for pid in all_pids:
        name = prod2name.get(pid, pid)
        월매출   = acc_rev.get(pid, 0)
        마진금   = acc_margin.get(pid, 0)
        매출광고 = ad_by_prod.get(pid, 0)
        신규광고 = new_by_prod.get(pid, 0)
        광고합계 = 매출광고 + 신규광고
        마진사용 = 광고합계 / 마진금 * 100 if 마진금 > 0 else None
        로아스   = 월매출 / 광고합계 * 100 if 광고합계 > 0 else None
        순이익   = 마진금 - 광고합계 if 마진금 > 0 else None

        전일매출 = day_rev.get(pid, 0)
        전일마진 = day_margin_s.get(pid, 0)

        rows.append({
            "등록상품ID": pid, "상품명": name,
            "누적_월매출": 월매출, "누적_마진금": 마진금,
            "누적_매출광고": 매출광고, "누적_신규광고": 신규광고,
            "누적_광고합계": 광고합계,
            "누적_마진금사용%": round(마진사용, 1) if 마진사용 else None,
            "누적_직접로아스%": round(로아스, 0) if 로아스 else None,
            "누적_순이익": round(순이익, 0) if 순이익 else None,
            "전일_매출": 전일매출, "전일_마진금": 전일마진,
        })
    return pd.DataFrame(rows)

def _enrich_ad(ad_df, opt2prod, prod2name, opt_mw):
    """광고 리포트에 상품명·마진 컬럼 추가 (내부용)"""
    df = ad_df.copy()
    df["pid"]        = df["opt"].map(opt2prod)
    df["상품명"]     = df["pid"].map(prod2name).fillna(df["prod_name_ad"])
    df["등록상품ID"] = df["pid"].fillna("-")
    df["마진금_개당"]= df["opt"].map(opt_mw).fillna(0)
    df["직접마진14"] = df["마진금_개당"] * df["d14_qty"]
    df["순이익14"]   = df["직접마진14"] - df["spend"]
    df["마진금사용%"]= (df["spend"] / df["직접마진14"] * 100).where(df["직접마진14"] > 0)
    df["직접로아스%"]= (df["d14_rev"] / df["spend"] * 100).where(df["spend"] > 0)
    return df

def _fmt_cols(df):
    """컬럼명 한글 통일 rename"""
    return df.rename(columns={
        "campaign_id": "캠페인ID", "campaign": "캠페인명", "ad_group": "광고그룹",
        "opt": "옵션ID", "impression": "노출수", "click": "클릭수",
        "spend": "광고비", "d1_rev": "직접매출(1일)",
        "d14_rev": "직접매출(14일)", "d14_qty": "직접판매수(14일)",
    })

def build_detail_by_opt(ad_df, new_df, opt2prod, prod2name, opt_mw, opt_rev_map=None, opt_qty_map=None, opt_type=None):
    """옵션ID 단위 합산 — 월 누적 판매량 기준 마진"""
    df = _enrich_ad(ad_df, opt2prod, prod2name, opt_mw)
    opt_to_name = ad_df.groupby("opt")["prod_name_ad"].first().to_dict()

    g = df.groupby(["opt"], as_index=False).agg(
        노출수=("impression","sum"), 클릭수=("click","sum"),
        매출광고비=("spend","sum"),
    )
    new_by = new_df.groupby("opt")["spend"].sum().reset_index().rename(columns={"spend":"신규광고비"})
    g = g.merge(new_by, on="opt", how="left")
    g["신규광고비"] = g["신규광고비"].fillna(0)
    g["광고합계"] = g["매출광고비"] + g["신규광고비"]

    g["옵션명"]   = g["opt"].map(opt_to_name).fillna("")
    g["판매방식"] = g["opt"].map(opt_type).fillna("-") if opt_type else "-"
    g["월매출"]   = g["opt"].map(opt_rev_map).fillna(0) if opt_rev_map else 0
    g["직접마진"] = g["opt"].map(opt_mw).fillna(0) * (g["opt"].map(opt_qty_map).fillna(0) if opt_qty_map else 0)
    g["순이익"]   = g["직접마진"] - g["광고합계"]
    g["마진금사용%"] = (g["광고합계"] / g["직접마진"] * 100).where(g["직접마진"] > 0)
    g["직접로아스%"]  = (g["월매출"] / g["광고합계"] * 100).where(g["광고합계"] > 0)
    return g.rename(columns={"opt":"옵션ID"})[[
        "옵션ID","판매방식","옵션명","노출수","클릭수",
        "매출광고비","신규광고비","광고합계",
        "월매출","직접마진","순이익","마진금사용%","직접로아스%"
    ]].sort_values("광고합계", ascending=False)

def build_detail_by_camp_opt(ad_df, new_df, opt2prod, prod2name, opt_mw, opt_rev_map=None, opt_qty_map=None, opt_type=None):
    """캠페인 × 옵션ID 단위 — 매출최적화 + 신규고객 유니온, 월 누적 마진"""
    df = _enrich_ad(ad_df, opt2prod, prod2name, opt_mw)
    opt_to_name = ad_df.groupby("opt")["prod_name_ad"].first().to_dict()

    # 매출최적화 집계
    g_매출 = df.groupby(["campaign_id","campaign","opt"], as_index=False).agg(
        노출수=("impression","sum"), 클릭수=("click","sum"),
        매출광고비=("spend","sum"),
    )
    g_매출["신규광고비"] = 0.0

    # 신규고객 집계 — campaign_id 포함
    g_신규 = new_df.groupby(["campaign_id","campaign","opt"])["spend"].sum().reset_index()
    g_신규.columns = ["campaign_id","campaign","opt","신규광고비"]
    g_신규["노출수"]     = 0
    g_신규["클릭수"]     = 0
    g_신규["매출광고비"] = 0.0

    cols = ["campaign_id","campaign","opt","노출수","클릭수","매출광고비","신규광고비"]
    g = pd.concat([g_매출[cols], g_신규[cols]], ignore_index=True)
    g["광고합계"] = g["매출광고비"] + g["신규광고비"]

    g["옵션명"]   = g["opt"].map(opt_to_name).fillna("")
    g["판매방식"] = g["opt"].map(opt_type).fillna("-") if opt_type else "-"
    g["월매출"]   = g["opt"].map(opt_rev_map).fillna(0) if opt_rev_map else 0
    g["직접마진"] = g["opt"].map(opt_mw).fillna(0) * (g["opt"].map(opt_qty_map).fillna(0) if opt_qty_map else 0)
    g["순이익"]   = g["직접마진"] - g["광고합계"]
    g["마진금사용%"] = (g["광고합계"] / g["직접마진"] * 100).where(g["직접마진"] > 0)
    g["직접로아스%"]  = (g["월매출"] / g["광고합계"] * 100).where(g["광고합계"] > 0)
    return g.rename(columns={"campaign_id":"캠페인ID","campaign":"캠페인명","opt":"옵션ID"})[[
        "캠페인ID","캠페인명","옵션ID","판매방식","옵션명",
        "노출수","클릭수","매출광고비","신규광고비","광고합계",
        "월매출","직접마진","순이익","마진금사용%","직접로아스%"
    ]].sort_values(["캠페인명","광고합계"], ascending=[True,False])


def build_campaign_table(ad_df, opt2prod, prod2name, opt_mw):
    """캠페인 단위 수익 분석 테이블"""
    ad_df = ad_df.copy()
    ad_df["pid"] = ad_df["opt"].map(opt2prod)
    ad_df["mw"]  = ad_df["opt"].map(opt_mw)
    ad_df["직접마진14"] = ad_df["mw"] * ad_df["d14_qty"]
    ad_df["순이익14"]   = ad_df["직접마진14"] - ad_df["spend"]

    g = ad_df.dropna(subset=["pid"]).groupby("campaign", as_index=False).agg(
        광고비=("spend", "sum"),
        직접마진=("직접마진14", "sum"),
        순이익=("순이익14", "sum"),
        직접매출14=("d14_rev", "sum"),
    )
    g["직접로아스%"] = (g["직접매출14"] / g["광고비"] * 100).round(0)
    g["마진금사용%"] = (g["광고비"] / g["직접마진"] * 100).round(1).where(g["직접마진"] > 0)
    return g.sort_values("순이익", ascending=False)
