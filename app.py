import re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="メンズビューティ販売分析ダッシュボード",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────
# 定数
# ─────────────────────────────────────────────
FISCAL_START_MONTH = 7
PERIOD_46 = (202407, 202506)
PERIOD_47 = (202507, 202606)

MONTH_LABEL = {
    1: "7月", 2: "8月", 3: "9月", 4: "10月", 5: "11月", 6: "12月",
    7: "1月", 8: "2月", 9: "3月", 10: "4月", 11: "5月", 12: "6月",
}

def ym_to_period(ym: int):
    if PERIOD_46[0] <= ym <= PERIOD_46[1]: return 46
    if PERIOD_47[0] <= ym <= PERIOD_47[1]: return 47
    return None

def ym_to_mip(ym: int) -> int:
    m = ym % 100
    return ((m - FISCAL_START_MONTH) % 12) + 1

def mip_to_yyyymm(period: int, mip: int) -> int:
    """期 + 期内月 → YYYYMM（例: 47期 期内月1 → 202507）"""
    start = PERIOD_46[0] if period == 46 else PERIOD_47[0]
    y, m = start // 100, start % 100
    total = y * 12 + (m - 1) + (mip - 1)
    return (total // 12) * 100 + (total % 12 + 1)

def fmt_yen(v):
    if pd.isna(v): return "—"
    if abs(v) >= 1_000_000: return f"¥{v/1_000_000:.1f}M"
    if abs(v) >= 1_000: return f"¥{v/1_000:.0f}K"
    return f"¥{v:.0f}"

def fmt_yoy(v):
    """昨対比を110.5%形式で表示。vは比率(110=前年比110%)"""
    if pd.isna(v): return "—"
    return f"{v:.1f}%"

def make_download_button(df: pd.DataFrame, filename: str, label: str = "📥 CSVダウンロード"):
    """数値を適切な桁数に丸めてCSVダウンロードボタンを表示する。"""
    out = df.copy()
    for col in out.select_dtypes(include="number").columns:
        if any(k in col for k in ["率","昨対","前年比","GAP"]):
            out[col] = out[col].round(1)
        elif any(k in col for k in ["金額","単価","売上"]):
            out[col] = out[col].round(0).astype("Int64", errors="ignore")
        else:
            out[col] = out[col].round(1)
    # utf-8-sig（BOM付き）をbytesで渡すとExcelで文字化けしない
    csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
    st.download_button(label, csv_bytes, file_name=filename, mime="text/csv", use_container_width=True)

# ─────────────────────────────────────────────
# データ読み込み
# ─────────────────────────────────────────────
@st.cache_data
def load_idpos(file) -> pd.DataFrame:
    df = pd.read_csv(file, encoding="cp932", dtype=str)
    df.columns = df.columns.str.strip()

    # 数値列マッピング
    col_map = {
        "売上数量":          "売上数量",
        "売上数量(前期)":    "売上数量_前期",
        "売上税抜金額(円)":          "売上金額",
        "売上税抜金額(円)(前期)":    "売上金額_前期",
        "POS客数":           "POS客数",
        "POS客数(前期)":     "POS客数_前期",
        "ID客数":            "ID客数",
        "ID客数(前期)":      "ID客数_前期",
    }
    for orig, new in col_map.items():
        if orig in df.columns:
            df[new] = pd.to_numeric(df[orig].str.replace(",", ""), errors="coerce").fillna(0)

    df["年月"] = pd.to_numeric(df["年月"], errors="coerce")
    df["期"]   = df["年月"].apply(ym_to_period)
    df["期内月"] = df["年月"].apply(ym_to_mip)
    df = df[df["期"].notna()].copy()
    df["期"] = df["期"].astype(int)
    return df

@st.cache_data
def load_sri(file) -> pd.DataFrame:
    """SRI+ エクスポートExcel（Sheet1, header=12）をロングフォームに変換して返す。
    返り値: JAN(str,13桁), YYYYMM(int), 市場金額(円換算), 市場数量 の列を持つDF。
    値は x1/1,000 表記なので ×1,000 して返す。
    """
    raw = pd.read_excel(file, sheet_name="Sheet1", header=12, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]

    # JAN列は "Unnamed: 2"（3列目）
    jan_col  = raw.columns[2]
    # 月次金額列: "YYYY/M-"、月次数量列: "YYYY/M-.1"
    amt_map = {}  # YYYYMM → col_name
    qty_map = {}
    for c in raw.columns:
        m = re.match(r'^(\d{4})/(\d{1,2})-$', c)
        if m:
            amt_map[int(m.group(1)) * 100 + int(m.group(2))] = c
        m2 = re.match(r'^(\d{4})/(\d{1,2})-\.1$', c)
        if m2:
            qty_map[int(m2.group(1)) * 100 + int(m2.group(2))] = c

    # JAN が13桁数字の行だけ残す
    raw["JAN"] = raw[jan_col].astype(str).str.strip().str.lstrip("0").str.zfill(13)
    df = raw[raw["JAN"].str.match(r'^\d{13}$')].copy()

    def to_num(s):
        return pd.to_numeric(s.astype(str).str.replace(",", "").str.replace("-", "0"), errors="coerce")

    records = []
    for yyyymm, acol in amt_map.items():
        tmp = pd.DataFrame({"JAN": df["JAN"]})
        tmp["YYYYMM"]   = yyyymm
        tmp["市場金額"]  = to_num(df[acol]) * 1000
        tmp["市場数量"]  = to_num(df[qty_map[yyyymm]]) * 1000 if yyyymm in qty_map else np.nan
        records.append(tmp)

    out = pd.concat(records, ignore_index=True)
    return out[out["市場金額"].notna() & (out["市場金額"] > 0)].copy()


@st.cache_data
def load_trmaster(file) -> pd.DataFrame:
    """月次レポートExcelからTRマスタシートを読み込む。
    返り値: JAN, サブカテゴリー名, セグメント名, サブセグメント名
    """
    df = pd.read_excel(file, sheet_name="TRマスタ", header=0, dtype=str)
    df.columns = df.columns.str.strip()
    df["JAN"] = df["JAN"].astype(str).str.lstrip("0").str.zfill(13)
    keep = [c for c in ["JAN", "サブカテゴリー名", "セグメント名", "サブセグメント名"] if c in df.columns]
    return df[keep].drop_duplicates("JAN").reset_index(drop=True)


def compute_market(df_sri_long, df_trmaster,
                   group_cols, subcat_col, seg_col, subseg_col,
                   sel_period, sel_months) -> pd.DataFrame | None:
    """選択期間の市場金額（今期・前年同期）と市場前年比をgroup_cols単位で集計して返す。"""
    if df_sri_long is None or df_trmaster is None:
        return None

    # TRマスタの列名をIDPOSの列名に合わせてリネーム
    col_rename = {}
    if subcat_col and "サブカテゴリー名" in df_trmaster.columns: col_rename["サブカテゴリー名"] = subcat_col
    if seg_col    and "セグメント名"     in df_trmaster.columns: col_rename["セグメント名"]     = seg_col
    if subseg_col and "サブセグメント名" in df_trmaster.columns: col_rename["サブセグメント名"] = subseg_col
    mst = df_trmaster.rename(columns=col_rename)
    seg_cols_available = [c for c in group_cols if c in mst.columns]
    if not seg_cols_available:
        return None

    # 期間 → YYYYMM 変換
    today_yms = [mip_to_yyyymm(sel_period,     m) for m in sel_months]
    prev_yms  = [mip_to_yyyymm(sel_period - 1, m) for m in sel_months]

    joined = df_sri_long.merge(mst[["JAN"] + seg_cols_available], on="JAN", how="inner")

    def agg_by(yms, label):
        return (joined[joined["YYYYMM"].isin(yms)]
                .groupby(seg_cols_available)["市場金額"].sum()
                .reset_index().rename(columns={"市場金額": label}))

    today = agg_by(today_yms, "市場金額_今期")
    prev  = agg_by(prev_yms,  "市場金額_前年")
    mrk = today.merge(prev, on=seg_cols_available, how="outer")
    mrk["市場前年比"] = mrk["市場金額_今期"] / mrk["市場金額_前年"].replace(0, np.nan) * 100

    # group_cols に足りない列を補完
    for c in group_cols:
        if c not in mrk.columns:
            mrk[c] = np.nan
    return mrk

@st.cache_data
def load_master(file) -> pd.DataFrame:
    df = pd.read_csv(file, encoding="cp932", dtype=str)
    df.columns = df.columns.str.strip()

    store_col = next((c for c in df.columns if "店舗CD" in c or "店舗コード" in c), None)
    if store_col is None:
        candidates = [c for c in df.columns if "CD" in c or "コード" in c]
        store_col = candidates[2] if len(candidates) >= 3 else df.columns[4]

    jan_col = next((c for c in df.columns if c == "JAN" or "JAN" in str(c).upper()), df.columns[0])
    df = df.rename(columns={store_col: "店舗CD", jan_col: "JAN"})
    df["JAN"] = df["JAN"].astype(str).str.zfill(13)

    store_count = (df.groupby("JAN")["店舗CD"].nunique().reset_index()
                   .rename(columns={"店舗CD": "採用店舗数"}))

    type_col = next((c for c in df.columns if "タイプ" in c and "CD" not in c), None)
    if type_col:
        type_df = df[["JAN", type_col]].drop_duplicates("JAN").rename(columns={type_col: "タイプ分類"})
        store_count = store_count.merge(type_df, on="JAN", how="left")

    return store_count

# ─────────────────────────────────────────────
# 集計ヘルパー
# ─────────────────────────────────────────────
def aggregate(df: pd.DataFrame, group_cols: list, value_cols: list = None) -> pd.DataFrame:
    if value_cols is None:
        value_cols = ["売上金額", "売上金額_前期", "売上数量", "売上数量_前期", "POS客数", "ID客数"]
    agg_dict = {c: "sum" for c in value_cols if c in df.columns}
    g = df.groupby(group_cols, dropna=False).agg(agg_dict).reset_index()
    if "売上金額" in g.columns and "売上金額_前期" in g.columns:
        g["昨対比"] = (g["売上金額"] / g["売上金額_前期"].replace(0, np.nan)) * 100
    return g

def filter_months(df: pd.DataFrame, period: int, months: list) -> pd.DataFrame:
    """指定期・指定月リストで絞り込む。複数月 = 合算表示。"""
    return df[(df["期"] == period) & (df["期内月"].isin(months))]

# ─────────────────────────────────────────────
# TRマスタ永続化（parquet）
# ─────────────────────────────────────────────
import pathlib
_IDPOS_CACHE    = pathlib.Path(__file__).parent / ".idpos_cache.parquet"
_SRI_CACHE      = pathlib.Path(__file__).parent / ".sri_cache.parquet"
_TRMASTER_CACHE = pathlib.Path(__file__).parent / ".trmaster_cache.parquet"
_MASTER_CACHE   = pathlib.Path(__file__).parent / ".master_cache.parquet"

def save_trmaster(df: pd.DataFrame):
    df.to_parquet(_TRMASTER_CACHE, index=False)

def load_trmaster_cache() -> pd.DataFrame | None:
    if _TRMASTER_CACHE.exists():
        return pd.read_parquet(_TRMASTER_CACHE)
    return None

def save_master(df: pd.DataFrame):
    df.to_parquet(_MASTER_CACHE, index=False)

def load_master_cache() -> pd.DataFrame | None:
    if _MASTER_CACHE.exists():
        return pd.read_parquet(_MASTER_CACHE)
    return None

def save_idpos(df: pd.DataFrame):
    df.to_parquet(_IDPOS_CACHE, index=False)

def load_idpos_cache() -> pd.DataFrame | None:
    if _IDPOS_CACHE.exists():
        return pd.read_parquet(_IDPOS_CACHE)
    return None

def save_sri(df: pd.DataFrame):
    df.to_parquet(_SRI_CACHE, index=False)

def load_sri_cache() -> pd.DataFrame | None:
    if _SRI_CACHE.exists():
        return pd.read_parquet(_SRI_CACHE)
    return None

# ─────────────────────────────────────────────
# サイドバー: ファイルアップロード
# ─────────────────────────────────────────────
with st.sidebar:
    show_upload = st.toggle("📁 データアップロード", value=True, key="show_upload")
    if show_upload:
        idpos_file = st.file_uploader("① IDPOS CSV（更新時のみ）", type=["csv"], key="idpos",
                                      help="アップロードするとローカルに保存され、次回以降は不要です")
        if idpos_file:
            st.caption("✅ IDPOSを保存しました（次回以降は不要）")
        elif _IDPOS_CACHE.exists():
            st.caption("💾 保存済みIDPOSを使用中")
            if st.button("🗑️ IDPOSをリセット", key="idpos_reset"):
                _IDPOS_CACHE.unlink(missing_ok=True)
                st.rerun()

        sri_file = st.file_uploader("② SRI Excel（更新時のみ）", type=["xlsx","xls"], key="sri",
                                    help="アップロードするとローカルに保存され、次回以降は不要です")
        if sri_file:
            st.caption("✅ SRIを保存しました（次回以降は不要）")
        elif _SRI_CACHE.exists():
            st.caption("💾 保存済みSRIを使用中")
            if st.button("🗑️ SRIをリセット", key="sri_reset"):
                _SRI_CACHE.unlink(missing_ok=True)
                st.rerun()
        master_file   = st.file_uploader(
            "③ マスタ CSV（半期更新）", type=["csv"], key="master",
            help="アップロードするとローカルに保存され、次回以降は不要です",
        )
        if master_file:
            st.caption("✅ マスタを保存しました（次回以降は不要）")
        elif _MASTER_CACHE.exists():
            st.caption("💾 保存済みマスタを使用中")
            if st.button("🗑️ マスタをリセット", key="master_reset"):
                _MASTER_CACHE.unlink(missing_ok=True)
                st.rerun()

        monthly_file  = st.file_uploader(
            "④ TRマスタ Excel（半期更新）",
            type=["xlsx","xls"], key="monthly",
            help="アップロードするとローカルに保存され、次回以降は不要です",
        )
        if monthly_file:
            st.caption("✅ TRマスタを保存しました（次回以降は不要）")
        elif _TRMASTER_CACHE.exists():
            st.caption("💾 保存済みTRマスタを使用中")
            if st.button("🗑️ TRマスタをリセット", key="trmaster_reset"):
                _TRMASTER_CACHE.unlink(missing_ok=True)
                st.rerun()
    else:
        idpos_file   = st.session_state.get("idpos")
        sri_file     = st.session_state.get("sri")
        master_file  = st.session_state.get("master")
        monthly_file = st.session_state.get("monthly")

st.title("📊 メンズビューティ販売分析ダッシュボード")

if idpos_file is None and not _IDPOS_CACHE.exists():
    st.info("👈 サイドバーからIDPOS CSVをアップロードしてください")
    st.stop()

# ─────────────────────────────────────────────
# データ読み込み
# ─────────────────────────────────────────────
with st.spinner("データ読み込み中..."):
    if idpos_file:
        df_all = load_idpos(idpos_file)
        save_idpos(df_all)
    else:
        df_all = load_idpos_cache()

    if sri_file:
        df_sri = load_sri(sri_file)
        save_sri(df_sri)
    else:
        df_sri = load_sri_cache()

    # マスタCSV: 新規アップロード → 保存 → 以降はキャッシュから
    if master_file:
        df_mst = load_master(master_file)
        save_master(df_mst)
    else:
        df_mst = load_master_cache()

    # TRマスタ: 新規アップロード → 保存 → 以降はキャッシュから
    if monthly_file:
        df_trmaster = load_trmaster(monthly_file)
        save_trmaster(df_trmaster)
    else:
        df_trmaster = load_trmaster_cache()

# ─────────────────────────────────────────────
# サイドバー: 期間・フィルタ設定
# ─────────────────────────────────────────────
# 列名解決
def find_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None

col_subcat = find_col(df_all, ["サブカテゴリー"])
col_seg    = find_col(df_all, ["セグメント"])
col_subseg = find_col(df_all, ["サブセグメント"])

available_periods = sorted(df_all["期"].unique())
available_months  = sorted(df_all["期内月"].unique())

with st.sidebar:
    st.markdown("---")
    st.header("📅 期・絞り込み")

    sel_period = st.selectbox(
        "期", available_periods, index=len(available_periods)-1,
        format_func=lambda p: f"{p}期（{p+1993}年7月〜{p+1994}年6月）"
    )

    st.markdown("---")
    st.header("🔍 絞り込み")
    st.caption("未選択＝全て表示　階層が連動します")

    # 階層フィルタ: サブカテゴリー
    df_filtered = df_all.copy()
    if col_subcat:
        all_subcats = sorted(df_all[col_subcat].dropna().unique())
        sel_subcats = st.multiselect(
            "① サブカテゴリー", all_subcats, default=[],
            placeholder="すべて（未選択）",
            key="filter_subcat",
        )
        if sel_subcats:
            df_filtered = df_filtered[df_filtered[col_subcat].isin(sel_subcats)]
    else:
        sel_subcats = []

    # セグメント（サブカテに連動）
    if col_seg:
        all_segs = sorted(df_filtered[col_seg].dropna().unique())
        sel_segs = st.multiselect(
            "② セグメント", all_segs, default=[],
            placeholder="すべて（未選択）",
            key="filter_seg",
        )
        if sel_segs:
            df_filtered = df_filtered[df_filtered[col_seg].isin(sel_segs)]
    else:
        sel_segs = []

    # サブセグメント（セグメントに連動）
    if col_subseg:
        all_subsegs = sorted(df_filtered[col_subseg].dropna().unique())
        sel_subsegs = st.multiselect(
            "③ サブセグメント", all_subsegs, default=[],
            placeholder="すべて（未選択）",
            key="filter_subseg",
        )
        if sel_subsegs:
            df_filtered = df_filtered[df_filtered[col_subseg].isin(sel_subsegs)]
    else:
        sel_subsegs = []

    # 絞り込み状況サマリー
    filter_labels = []
    if sel_subcats: filter_labels.append(f"サブカテ: {', '.join(sel_subcats)}")
    if sel_segs:    filter_labels.append(f"セグ: {', '.join(sel_segs)}")
    if sel_subsegs: filter_labels.append(f"サブセグ: {', '.join(sel_subsegs)}")
    st.caption("絞り込みは全タブに反映されます")

# ─────────────────────────────────────────────
# メイン: 月選択UI（複数選択で自動累計）
# ─────────────────────────────────────────────
months_in_period = sorted(df_all[df_all["期"] == sel_period]["期内月"].unique())

with st.container(border=True):
    c_label, c_reset = st.columns([6, 1])
    c_label.markdown("**月を選択** — 複数選択で合算（累計）表示")
    if c_reset.button("リセット", key="month_reset", use_container_width=True):
        st.session_state["month_pills"] = [max(months_in_period)]
        st.rerun()
    sel_months = st.pills(
        "月番号",
        options=months_in_period,
        format_func=lambda m: MONTH_LABEL[m],
        default=[max(months_in_period)],
        selection_mode="multi",
        key="month_pills",
        label_visibility="collapsed",
    )
    if not sel_months:
        sel_months = [max(months_in_period)]

is_cumulative = len(sel_months) > 1
sel_month = max(sel_months)

if is_cumulative:
    period_label = f"{sel_period}期 {MONTH_LABEL[min(sel_months)]}〜{MONTH_LABEL[sel_month]} 合算"
else:
    period_label = f"{sel_period}期 {MONTH_LABEL[sel_month]}"

# ─────────────────────────────────────────────
# 表示期間データ抽出（前期は同行の(前期)列を使用）
# ─────────────────────────────────────────────
df_cur = filter_months(df_filtered, sel_period, sel_months)

# ─────────────────────────────────────────────
# KPIカード（ヘッダー）
# ─────────────────────────────────────────────
total_sales    = df_cur["売上金額"].sum()
total_prev     = df_cur["売上金額_前期"].sum() if "売上金額_前期" in df_cur.columns else 0
total_qty      = df_cur["売上数量"].sum()
total_qty_prev = df_cur["売上数量_前期"].sum() if "売上数量_前期" in df_cur.columns else 0
total_pos      = df_cur["POS客数"].sum()
total_pos_prev = df_cur["POS客数_前期"].sum() if "POS客数_前期" in df_cur.columns else 0
total_id       = df_cur["ID客数"].sum()
total_id_prev  = df_cur["ID客数_前期"].sum() if "ID客数_前期" in df_cur.columns else 0
total_price      = total_sales / total_pos  if total_pos  > 0 else np.nan
total_price_prev = total_prev  / total_pos_prev if total_pos_prev > 0 else np.nan
total_avg_price      = total_sales / total_qty      if total_qty      > 0 else np.nan
total_avg_price_prev = total_prev  / total_qty_prev if total_qty_prev > 0 else np.nan

def yoy_ratio(cur, prev):
    """昨対比（110.5% 形式）。前年なし → None。"""
    if prev <= 0 or pd.isna(cur) or pd.isna(prev): return None
    return cur / prev * 100

def fmt_yoy_ratio(v):
    return f"{v:.1f}%" if v is not None else "—"

col1, col2, col3, col4, col5, col6 = st.columns(6)
for col, label, cur_v, prev_v in [
    (col1, "売上金額（税抜）", total_sales,       total_prev),
    (col2, "売上数量",         total_qty,         total_qty_prev),
    (col3, "平均単価",         total_avg_price,   total_avg_price_prev),
    (col4, "購買単価",         total_price,       total_price_prev),
    (col5, "POS客数",          total_pos,         total_pos_prev),
    (col6, "ID客数",           total_id,          total_id_prev),
]:
    ratio = yoy_ratio(cur_v, prev_v)
    val_str = fmt_yen(cur_v) if label in ("売上金額（税抜）", "平均単価", "購買単価") else f"{cur_v:,.0f}"
    col.metric(label, val_str)
    if ratio is not None:
        color = "#2ca02c" if ratio >= 100 else "#d62728"
        arrow = "▲" if ratio >= 100 else "▼"
        col.markdown(
            f'<div style="font-size:1.3rem; font-weight:bold; color:{color};">'
            f'{arrow} {ratio:.1f}%</div>'
            f'<div style="font-size:0.75rem; color:#888;">昨対比</div>',
            unsafe_allow_html=True,
        )
    else:
        col.caption("昨対比 —")
filter_summary = "　".join(filter_labels) if filter_labels else "絞り込みなし（全て）"
st.caption(f"表示期間: {period_label}　　絞り込み: {filter_summary}")

st.markdown("---")

# ─────────────────────────────────────────────
# タブ
# ─────────────────────────────────────────────
tab_seg, tab_subseg, tab_trend, tab_rank, tab_teiban = st.tabs(
    ["📋 セグメント別", "🔬 サブセグメント別", "📈 月別トレンド", "🏆 単品ランキング", "🏪 定番分析"]
)

# ═══════════════════════════════════════════════
# Tab1: セグメント別（PowerPシートイメージ）
# ═══════════════════════════════════════════════
with tab_seg:
    st.subheader(f"セグメント別サマリー — {period_label}")

    group_cols = [c for c in [col_subcat, col_seg] if c]
    if not group_cols:
        st.warning("サブカテゴリー/セグメント列が見つかりません")
    else:
        # aggregate は売上金額_前期も集計し昨対比を自動計算
        agg = aggregate(df_cur, group_cols)
        # 列名を統一（前期売上 → 売上金額_前期）
        if "売上金額_前期" in agg.columns:
            agg = agg.rename(columns={"売上金額_前期": "前期売上"})

        # 平均単価（金額÷数量）と昨対比
        agg["平均単価"] = agg["売上金額"] / agg["売上数量"].replace(0, np.nan)
        if "前期売上" in agg.columns and "売上数量_前期" in agg.columns:
            agg["前期平均単価"] = agg["前期売上"] / agg["売上数量_前期"].replace(0, np.nan)
            agg["平均単価昨対"] = agg["平均単価"] / agg["前期平均単価"].replace(0, np.nan) * 100

        # 市場データ（SRI × TRマスタ）
        mrk = compute_market(df_sri, df_trmaster, group_cols,
                             col_subcat, col_seg, col_subseg,
                             sel_period, sel_months)
        if mrk is not None:
            agg = agg.merge(mrk[group_cols + ["市場前年比"]], on=group_cols, how="left")
            agg["昨対GAP"] = agg["昨対比"] - agg["市場前年比"]

        # 表示列
        show = [c for c in agg.columns if c in group_cols + [
            "売上金額","前期売上","昨対比","市場前年比","昨対GAP",
            "売上数量","平均単価","平均単価昨対","POS客数","ID客数"]]
        fmt = {
            "売上金額":    "¥{:,.0f}",
            "前期売上":    "¥{:,.0f}",
            "昨対比":      "{:.1f}%",
            "市場前年比":  "{:.1f}%",
            "昨対GAP":     "{:+.1f}pp",
            "売上数量":    "{:,.0f}",
            "平均単価":    "¥{:,.0f}",
            "平均単価昨対": "{:.1f}%",
            "POS客数":     "{:,.0f}",
            "ID客数":      "{:,.0f}",
        }
        fmt_use = {k: v for k, v in fmt.items() if k in show}

        # サブカテゴリー小計行を差し込む
        sum_cols = [c for c in ["売上金額","前期売上","売上数量","売上数量_前期","POS客数","ID客数"] if c in agg.columns]
        if col_subcat and col_seg:
            subtotals = agg.groupby(col_subcat, as_index=False).agg({c: "sum" for c in sum_cols})
            if "売上金額" in subtotals.columns and "前期売上" in subtotals.columns:
                subtotals["昨対比"] = subtotals["売上金額"] / subtotals["前期売上"].replace(0, np.nan) * 100
            subtotals["平均単価"] = subtotals["売上金額"] / subtotals["売上数量"].replace(0, np.nan)
            if "前期売上" in subtotals.columns and "売上数量_前期" in subtotals.columns:
                subtotals["前期平均単価"] = subtotals["前期売上"] / subtotals["売上数量_前期"].replace(0, np.nan)
                subtotals["平均単価昨対"] = subtotals["平均単価"] / subtotals["前期平均単価"].replace(0, np.nan) * 100
            # 市場小計は按分せず再集計
            if mrk is not None:
                mrk_sub = (mrk.groupby(col_subcat)[["市場前年比"]].mean()
                           .reset_index()) if col_subcat in mrk.columns else None
                if mrk_sub is not None:
                    subtotals = subtotals.merge(mrk_sub, on=col_subcat, how="left")
                    subtotals["昨対GAP"] = subtotals["昨対比"] - subtotals["市場前年比"]
            subtotals[col_seg] = "【小計】"
            subtotals = subtotals.reindex(columns=agg.columns)
            combined = pd.concat([agg, subtotals]).sort_values([col_subcat, col_seg]).reset_index(drop=True)
        else:
            combined = agg.reset_index(drop=True)

        display_df = combined[show].reset_index(drop=True)

        # スタイル: 小計行ハイライト + GAP色
        def style_seg_table(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            if col_seg in df.columns:
                is_sub = df[col_seg].astype(str) == "【小計】"
                styles.loc[is_sub] = "background-color:#fff3cd; font-weight:bold"
            if "昨対GAP" in df.columns:
                for idx, val in df["昨対GAP"].items():
                    if pd.isna(val): continue
                    if val >= 3:    styles.loc[idx, "昨対GAP"] = "color:#1a7a1a; font-weight:bold"
                    elif val >= 0:  styles.loc[idx, "昨対GAP"] = "color:#2ca02c"
                    elif val >= -3: styles.loc[idx, "昨対GAP"] = "color:#ff7f0e"
                    else:           styles.loc[idx, "昨対GAP"] = "color:#d62728; font-weight:bold"
            return styles

        make_download_button(display_df, f"セグメント別_{period_label}.csv")
        st.dataframe(
            display_df.style
            .format(fmt_use, na_rep="—")
            .apply(style_seg_table, axis=None),
            use_container_width=True, hide_index=True
        )

        # ウォーターフォール風棒グラフ（構成比ラベル付き）
        total_for_share = agg["売上金額"].sum()
        agg_plot = agg.sort_values("売上金額", ascending=False).copy()
        agg_plot["構成比"] = agg_plot["売上金額"] / total_for_share * 100
        agg_plot["ラベル"] = agg_plot.apply(
            lambda r: f"{r['売上金額']/1e6:.1f}M  ({r['構成比']:.1f}%)", axis=1
        )
        y_col_plot = col_seg if col_seg else col_subcat
        fig = px.bar(
            agg_plot,
            x="売上金額", y=y_col_plot,
            color=col_subcat if col_subcat else col_seg,
            orientation="h",
            title="セグメント別売上（構成比）",
            text="ラベル",
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_yaxes(categoryorder="total ascending")
        fig.update_layout(
            height=max(400, len(agg)*32),
            legend=dict(orientation="h", y=-0.2),
            xaxis=dict(title="売上金額（円）"),
            uniformtext_minsize=9, uniformtext_mode="hide",
        )
        st.plotly_chart(fig, use_container_width=True)

        # 昨対比ヒートマップ
        if col_subcat and col_seg and "昨対比" in agg.columns:
            pivot = agg.pivot_table(index=col_subcat, columns=col_seg, values="昨対比")
            if not pivot.empty:
                fig2 = px.imshow(
                    pivot, text_auto=".1f", aspect="auto",
                    color_continuous_scale=["#d62728","#ffffff","#2ca02c"],
                    color_continuous_midpoint=100,
                    title="昨対比ヒートマップ（%）",
                )
                st.plotly_chart(fig2, use_container_width=True)


# ═══════════════════════════════════════════════
# Tab2: サブセグメント別
# ═══════════════════════════════════════════════
with tab_subseg:
    st.subheader(f"サブセグメント別詳細 — {period_label}")

    if not col_subseg:
        st.info("サブセグメント列が見つかりません")
    else:
        grp = [c for c in [col_subcat, col_seg, col_subseg] if c]
        agg_ss = aggregate(
            df_cur, grp,
            value_cols=["売上金額","売上金額_前期","売上数量","売上数量_前期",
                        "POS客数","POS客数_前期","ID客数","ID客数_前期"],
        )

        # 購買単価・前期購買単価
        if "POS客数" in agg_ss.columns:
            agg_ss["購買単価"] = agg_ss["売上金額"] / agg_ss["POS客数"].replace(0, np.nan)
        if "POS客数_前期" in agg_ss.columns:
            agg_ss["前期購買単価"] = (
                agg_ss.get("売上金額_前期", pd.Series(np.nan, index=agg_ss.index))
                / agg_ss["POS客数_前期"].replace(0, np.nan)
            )

        # 昨対比4指標
        def safe_yoy(cur_col, prev_col, df):
            if cur_col in df.columns and prev_col in df.columns:
                return (df[cur_col] / df[prev_col].replace(0, np.nan) * 100)
            return pd.Series(np.nan, index=df.index)

        agg_ss["金額昨対"]  = safe_yoy("売上金額",   "売上金額_前期",   agg_ss)
        agg_ss["数量昨対"]  = safe_yoy("売上数量",   "売上数量_前期",   agg_ss)
        agg_ss["単価昨対"]  = safe_yoy("購買単価",   "前期購買単価",    agg_ss)
        agg_ss["POS昨対"]   = safe_yoy("POS客数",    "POS客数_前期",    agg_ss)

        # 平均単価（金額÷数量）と昨対比
        agg_ss["平均単価"] = agg_ss["売上金額"] / agg_ss["売上数量"].replace(0, np.nan)
        if "売上金額_前期" in agg_ss.columns and "売上数量_前期" in agg_ss.columns:
            agg_ss["前期平均単価"] = agg_ss["売上金額_前期"] / agg_ss["売上数量_前期"].replace(0, np.nan)
            agg_ss["平均単価昨対"] = agg_ss["平均単価"] / agg_ss["前期平均単価"].replace(0, np.nan) * 100

        agg_ss["ID昨対"] = safe_yoy("ID客数", "ID客数_前期", agg_ss)

        # 市場前年比・GAP（サブセグメント粒度）
        mrk_ss = compute_market(df_sri, df_trmaster, grp,
                                col_subcat, col_seg, col_subseg,
                                sel_period, sel_months)
        if mrk_ss is not None:
            agg_ss = agg_ss.merge(mrk_ss[grp + ["市場前年比"]], on=grp, how="left")
            agg_ss["昨対GAP"] = agg_ss["金額昨対"] - agg_ss["市場前年比"]

        show_cols = grp + [c for c in [
            "売上金額","金額昨対","市場前年比","昨対GAP",
            "売上数量","数量昨対",
            "平均単価","平均単価昨対",
            "購買単価","単価昨対",
            "POS客数","POS昨対",
            "ID客数","ID昨対",
        ] if c in agg_ss.columns]
        fmt2 = {
            "売上金額":   "¥{:,.0f}",
            "金額昨対":   "{:.1f}%",
            "市場前年比": "{:.1f}%",
            "昨対GAP":    "{:+.1f}pp",
            "売上数量":   "{:,.0f}",
            "数量昨対":   "{:.1f}%",
            "平均単価":   "¥{:,.0f}",
            "平均単価昨対": "{:.1f}%",
            "購買単価":   "¥{:,.0f}",
            "単価昨対":   "{:.1f}%",
            "POS客数":    "{:,.0f}",
            "POS昨対":    "{:.1f}%",
            "ID客数":     "{:,.0f}",
            "ID昨対":     "{:.1f}%",
        }

        # 昨対比・GAP列のカラースケール
        def color_yoy_cols(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for col in ["金額昨対","数量昨対","平均単価昨対","単価昨対","POS昨対","ID昨対"]:
                if col not in df.columns: continue
                for idx, val in df[col].items():
                    if pd.isna(val): continue
                    if val >= 105:   styles.loc[idx, col] = "color:#1a7a1a; font-weight:bold"
                    elif val >= 100: styles.loc[idx, col] = "color:#2ca02c"
                    elif val >= 95:  styles.loc[idx, col] = "color:#ff7f0e"
                    else:            styles.loc[idx, col] = "color:#d62728; font-weight:bold"
            if "昨対GAP" in df.columns:
                for idx, val in df["昨対GAP"].items():
                    if pd.isna(val): continue
                    if val >= 3:    styles.loc[idx, "昨対GAP"] = "color:#1a7a1a; font-weight:bold"
                    elif val >= 0:  styles.loc[idx, "昨対GAP"] = "color:#2ca02c"
                    elif val >= -3: styles.loc[idx, "昨対GAP"] = "color:#ff7f0e"
                    else:           styles.loc[idx, "昨対GAP"] = "color:#d62728; font-weight:bold"
            return styles

        make_download_button(agg_ss[show_cols].sort_values("売上金額", ascending=False),
                             f"サブセグメント別_{period_label}.csv")
        st.dataframe(
            agg_ss[show_cols].sort_values("売上金額", ascending=False)
            .style
            .format({k: v for k, v in fmt2.items() if k in show_cols}, na_rep="—")
            .apply(color_yoy_cols, axis=None),
            use_container_width=True, hide_index=True
        )

        # バブルチャート: 金額昨対 × 売上金額
        if "金額昨対" in agg_ss.columns:
            fig = px.scatter(
                agg_ss.dropna(subset=["金額昨対","売上金額"]),
                x="金額昨対", y="売上金額",
                size="売上数量" if "売上数量" in agg_ss.columns else None,
                color=col_seg if col_seg else col_subcat,
                hover_name=col_subseg,
                labels={"昨対比_plot": "昨対比（%）　※新規商品は0%表示"},
                title="売上昨対比 × 売上金額（バブル＝売上数量）",
            )
            fig.add_vline(x=100, line_dash="dash", line_color="gray",
                          annotation_text="前年同期", annotation_position="top right")
            st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════
# Tab3: 月別トレンド
# ═══════════════════════════════════════════════
with tab_trend:
    st.subheader("月別トレンド — 46期 vs 47期")

    trend_type = st.radio("表示タイプ", ["月次（単月）", "累計推移"], horizontal=True, key="trend_type")
    trend_unit = st.radio(
        "集計単位",
        ["全体"] + ([col_subcat] if col_subcat else []) + ([col_seg] if col_seg else []),
        horizontal=True, key="trend_unit"
    )

    def build_monthly(df_base, unit_col=None):
        grp = ["期", "期内月"] + ([unit_col] if unit_col else [])
        out = df_base.groupby(grp)["売上金額"].sum().reset_index()
        if trend_type == "累計推移":
            if unit_col:
                out = out.sort_values(["期", unit_col, "期内月"])
                out["売上金額_累計"] = out.groupby(["期", unit_col])["売上金額"].cumsum()
            else:
                out = out.sort_values(["期", "期内月"])
                out["売上金額_累計"] = out.groupby("期")["売上金額"].cumsum()
        return out

    y_col = "売上金額_累計" if trend_type == "累計推移" else "売上金額"
    y_label = "累計売上金額（円）" if trend_type == "累計推移" else "売上金額（円）"

    if trend_unit == "全体":
        tdf = build_monthly(df_filtered)
        tdf["期ラベル"] = tdf["期"].astype(str) + "期"
        tdf["月ラベル"] = tdf["期内月"].map(MONTH_LABEL)

        fig = go.Figure()
        colors = {46: "#1f77b4", 47: "#ff7f0e"}
        for p in sorted(tdf["期"].unique()):
            d = tdf[tdf["期"] == p].sort_values("期内月")
            fig.add_trace(go.Scatter(
                x=d["期内月"], y=d[y_col],
                name=f"{p}期",
                mode="lines+markers",
                line=dict(color=colors.get(p,"#2ca02c"), width=3),
                marker=dict(size=9),
                customdata=d["月ラベル"],
                hovertemplate="%{customdata}<br>¥%{y:,.0f}<extra>%{fullData.name}</extra>",
            ))
    else:
        unit_col = trend_unit
        items = sorted(df_filtered[unit_col].dropna().unique())
        sel_items = st.multiselect(f"{unit_col}を選択", items, default=items[:6], key="trend_items")
        tdf = build_monthly(df_filtered[df_filtered[unit_col].isin(sel_items)], unit_col)

        palette = px.colors.qualitative.Set2
        fig = go.Figure()
        for i, item in enumerate(sel_items):
            for p in sorted(tdf["期"].unique()):
                d = tdf[(tdf["期"]==p) & (tdf[unit_col]==item)].sort_values("期内月")
                fig.add_trace(go.Scatter(
                    x=d["期内月"], y=d[y_col],
                    name=f"{item} / {p}期",
                    mode="lines+markers",
                    line=dict(color=palette[i % len(palette)],
                              dash="solid" if p == sel_period else "dot", width=2),
                ))

    fig.update_layout(
        title=f"{'累計' if trend_type=='累計推移' else '月次'}売上推移（46期 vs 47期）",
        xaxis=dict(title="期内月", tickvals=list(range(1,13)),
                   ticktext=[MONTH_LABEL[m] for m in range(1,13)]),
        yaxis=dict(title=y_label),
        hovermode="x unified",
        legend=dict(orientation="h", y=-0.25),
    )
    # 選択月に縦線
    fig.add_vline(x=sel_month, line_dash="dash", line_color="red",
                  annotation_text=f"▶ {MONTH_LABEL[sel_month]}", annotation_position="top right")
    st.plotly_chart(fig, use_container_width=True)

    # 昨対比バーチャート（47期の前期列を使用）
    st.markdown("#### 月別昨対比（47期 vs 前年同月）")
    df47 = df_filtered[df_filtered["期"] == sel_period]
    if "売上金額_前期" in df47.columns:
        monthly_cur  = df47.groupby("期内月")["売上金額"].sum()
        monthly_prev = df47.groupby("期内月")["売上金額_前期"].sum()
    else:
        monthly_cur  = df47.groupby("期内月")["売上金額"].sum()
        monthly_prev = df_filtered[df_filtered["期"]==sel_period-1].groupby("期内月")["売上金額"].sum()
    if trend_type == "累計推移":
        monthly_cur  = monthly_cur.sort_index().cumsum()
        monthly_prev = monthly_prev.sort_index().cumsum()
    yoy_s = ((monthly_cur / monthly_prev.replace(0, np.nan)) * 100).reset_index()
    yoy_s.columns = ["期内月", "昨対比(%)"]
    yoy_s["月"] = yoy_s["期内月"].map(MONTH_LABEL)
    fig2 = px.bar(
        yoy_s, x="月", y="昨対比(%)",
        color="昨対比(%)",
        color_continuous_scale=["#d62728","#ffffff","#2ca02c"],
        color_continuous_midpoint=100,
        title="昨対比（%）　100%=前年同期並み",
        text_auto=".1f",
    )
    fig2.add_hline(y=100, line_color="black", line_width=1, line_dash="dash")
    fig2.update_layout(coloraxis_showscale=False, xaxis_title="",
                       yaxis=dict(ticksuffix="%"))
    st.plotly_chart(fig2, use_container_width=True)


# ═══════════════════════════════════════════════
# Tab4: 単品ランキング（単品Rシートイメージ）
# ═══════════════════════════════════════════════
with tab_rank:
    st.subheader(f"単品ランキング — {period_label}")

    c1, c2, c3 = st.columns(3)
    rank_by   = c1.radio("ランキング基準", ["売上金額","売上数量","ID客数"], horizontal=True, key="rank_by")
    top_n     = c2.select_slider("表示件数", [10,20,30,50,100], value=30, key="top_n")
    show_prev = c3.checkbox("前年比較を表示", value=True)

    jan_agg = df_cur.groupby("JAN").agg(
        商品名   =("商品名","first"),
        メーカー  =("メーカー","first") if "メーカー" in df_cur.columns else ("JAN","first"),
        **({col_subcat: (col_subcat,"first")} if col_subcat else {}),
        **({col_seg:    (col_seg,   "first")} if col_seg    else {}),
        売上金額  =("売上金額","sum"),
        売上数量  =("売上数量","sum"),
        POS客数   =("POS客数", "sum"),
        ID客数    =("ID客数",  "sum"),
    ).reset_index()

    # 前期列はCSV内の(前期)列を使用
    jan_prev = df_cur.groupby("JAN").agg(
        前期売上金額=("売上金額_前期","sum"),
        前期売上数量=("売上数量_前期","sum"),
    ).reset_index() if "売上金額_前期" in df_cur.columns else pd.DataFrame()

    if show_prev and not jan_prev.empty:
        jan_agg = jan_agg.merge(jan_prev, on="JAN", how="left")
        jan_agg["昨対比"]   = (jan_agg["売上金額"] / jan_agg["前期売上金額"].replace(0,np.nan)) * 100
        jan_agg["数量昨対"] = (jan_agg["売上数量"] / jan_agg["前期売上数量"].replace(0,np.nan)) * 100

    jan_agg["平均単価"] = jan_agg["売上金額"] / jan_agg["売上数量"].replace(0, np.nan)
    if show_prev and "前期売上金額" in jan_agg.columns and "前期売上数量" in jan_agg.columns:
        jan_agg["前期平均単価"] = jan_agg["前期売上金額"] / jan_agg["前期売上数量"].replace(0, np.nan)
        jan_agg["平均単価昨対"] = jan_agg["平均単価"] / jan_agg["前期平均単価"].replace(0, np.nan) * 100

    if df_mst is not None:
        jan_agg = jan_agg.merge(df_mst[["JAN","採用店舗数"]], on="JAN", how="left")

    # 市場前年比・昨対GAP（JAN単位でSRIデータと結合）
    if df_sri is not None:
        # IDPOSのJANは20桁ゼロ埋め → 末尾13桁がEANコード
        jan_agg["JAN13"] = jan_agg["JAN"].astype(str).str.lstrip("0").str.zfill(13)
        today_yms_rank = [mip_to_yyyymm(sel_period,     m) for m in sel_months]
        prev_yms_rank  = [mip_to_yyyymm(sel_period - 1, m) for m in sel_months]
        sri_today = (df_sri[df_sri["YYYYMM"].isin(today_yms_rank)]
                     .groupby("JAN")["市場金額"].sum().reset_index()
                     .rename(columns={"市場金額": "市場金額_今期"}))
        sri_prev  = (df_sri[df_sri["YYYYMM"].isin(prev_yms_rank)]
                     .groupby("JAN")["市場金額"].sum().reset_index()
                     .rename(columns={"市場金額": "市場金額_前年"}))
        sri_jan = sri_today.merge(sri_prev, on="JAN", how="outer")
        sri_jan["市場前年比"] = sri_jan["市場金額_今期"] / sri_jan["市場金額_前年"].replace(0, np.nan) * 100
        jan_agg = jan_agg.merge(sri_jan[["JAN","市場前年比"]].rename(columns={"JAN":"JAN13"}), on="JAN13", how="left")
        jan_agg = jan_agg.drop(columns=["JAN13"])
        if "昨対比" in jan_agg.columns:
            jan_agg["昨対GAP"] = jan_agg["昨対比"] - jan_agg["市場前年比"]

    jan_agg = jan_agg.sort_values(rank_by, ascending=False).head(top_n).reset_index(drop=True)
    jan_agg.insert(0, "順位", range(1, len(jan_agg)+1))

    # 表示列構築
    base_cols = ["順位","JAN","商品名"]
    if col_subcat and col_subcat in jan_agg.columns: base_cols.append(col_subcat)
    if col_seg    and col_seg    in jan_agg.columns: base_cols.append(col_seg)
    base_cols += ["売上金額","売上数量","平均単価"]
    if show_prev and "前期売上金額" in jan_agg.columns:
        base_cols += ["前期売上金額","昨対比","数量昨対","平均単価昨対"]
    if "市場前年比" in jan_agg.columns: base_cols.append("市場前年比")
    if "昨対GAP"   in jan_agg.columns: base_cols.append("昨対GAP")
    if "採用店舗数" in jan_agg.columns: base_cols.append("採用店舗数")

    base_cols = [c for c in base_cols if c in jan_agg.columns]
    fmt_r = {
        "売上金額":"¥{:,.0f}","前期売上金額":"¥{:,.0f}","売上数量":"{:,.0f}",
        "平均単価":"¥{:,.0f}","昨対比":"{:.1f}%","数量昨対":"{:.1f}%",
        "平均単価昨対":"{:.1f}%","市場前年比":"{:.1f}%","昨対GAP":"{:+.1f}pp",
        "採用店舗数":"{:.0f}",
    }

    def color_rank_table(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for col in ["昨対比","数量昨対","平均単価昨対"]:
            if col not in df.columns: continue
            for idx, val in df[col].items():
                if pd.isna(val): continue
                if val >= 105:   styles.loc[idx, col] = "color:#1a7a1a; font-weight:bold"
                elif val >= 100: styles.loc[idx, col] = "color:#2ca02c"
                elif val >= 95:  styles.loc[idx, col] = "color:#ff7f0e"
                else:            styles.loc[idx, col] = "color:#d62728; font-weight:bold"
        if "昨対GAP" in df.columns:
            for idx, val in df["昨対GAP"].items():
                if pd.isna(val): continue
                if val >= 3:    styles.loc[idx, "昨対GAP"] = "color:#1a7a1a; font-weight:bold"
                elif val >= 0:  styles.loc[idx, "昨対GAP"] = "color:#2ca02c"
                elif val >= -3: styles.loc[idx, "昨対GAP"] = "color:#ff7f0e"
                else:           styles.loc[idx, "昨対GAP"] = "color:#d62728; font-weight:bold"
        return styles

    make_download_button(jan_agg[base_cols], f"単品ランキング_{period_label}.csv")
    st.dataframe(
        jan_agg[base_cols].style
        .format({k:v for k,v in fmt_r.items() if k in base_cols}, na_rep="—")
        .apply(color_rank_table, axis=None),
        use_container_width=True, hide_index=True, height=600,
    )

    # 横棒グラフ
    fig = px.bar(
        jan_agg,
        x=rank_by, y="商品名",
        orientation="h",
        color=col_seg if col_seg and col_seg in jan_agg.columns else (col_subcat if col_subcat else None),
        hover_data=["JAN","メーカー"] if "メーカー" in jan_agg.columns else ["JAN"],
        title=f"TOP{top_n} 単品ランキング（{rank_by}）",
        text_auto=".3s",
    )
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_layout(height=max(500, top_n * 24), legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════
# Tab5: 定番分析（採用店舗数×平均売上）
# ═══════════════════════════════════════════════
with tab_teiban:
    st.subheader(f"定番分析 — {period_label}")

    if df_mst is None:
        st.warning("マスタCSVをアップロードすると定番分析が表示されます")
    else:
        jan_sales = df_cur.groupby("JAN").agg(
            商品名  =("商品名","first"),
            **({col_subcat: (col_subcat,"first")} if col_subcat else {}),
            **({col_seg:    (col_seg,   "first")} if col_seg    else {}),
            売上金額 =("売上金額","sum"),
            売上数量 =("売上数量","sum"),
        ).reset_index()

        jan_sales = jan_sales.merge(df_mst, on="JAN", how="left")
        jan_sales["採用店舗数"] = pd.to_numeric(jan_sales["採用店舗数"], errors="coerce")
        jan_sales["1店舗あたり売上"] = jan_sales["売上金額"] / jan_sales["採用店舗数"].replace(0, np.nan)

        has_yoy = "売上金額_前期" in df_cur.columns
        if has_yoy:
            prev_jan = df_cur.groupby("JAN")["売上金額_前期"].sum().reset_index().rename(columns={"売上金額_前期":"前期売上"})
            jan_sales = jan_sales.merge(prev_jan, on="JAN", how="left")
            jan_sales["昨対比"] = (jan_sales["売上金額"] / jan_sales["前期売上"].replace(0, np.nan)) * 100

        # ── 象限分類（採用店舗数 × 昨対比） ──────────────────
        med_store = jan_sales["採用店舗数"].median()

        if has_yoy:
            def classify_quad(r):
                # 前期売上なし＝新規商品は別扱い
                if pd.isna(r["前期売上"]) or r["前期売上"] <= 0:
                    return "🆕 新規商品"
                growing = r["昨対比"] >= 100 if not pd.isna(r["昨対比"]) else False
                wide    = r["採用店舗数"] >= med_store if not pd.isna(r["採用店舗数"]) else False
                if   wide and growing:     return "🚀 主力成長"
                elif wide and not growing: return "⚠️ 主力低迷"
                elif not wide and growing: return "📈 伸び盛り"
                else:                     return "❌ 課題商品"
            jan_sales["象限"] = jan_sales.apply(classify_quad, axis=1)

            QUAD_COLOR = {
                "🚀 主力成長": "#2ca02c",
                "📈 伸び盛り": "#98df8a",
                "⚠️ 主力低迷": "#ff7f0e",
                "❌ 課題商品": "#d62728",
                "🆕 新規商品": "#1f77b4",
            }
            QUAD_ORDER = ["🚀 主力成長", "📈 伸び盛り", "⚠️ 主力低迷", "❌ 課題商品", "🆕 新規商品"]

            # 新規商品は昨対比NaNだがプロットする（y軸は0で表示）
            plot_df = jan_sales.dropna(subset=["採用店舗数"]).copy()
            plot_df["昨対比_plot"] = plot_df["昨対比"].fillna(0)
            plot_df["昨対比_表示"] = plot_df["昨対比"].map(lambda v: f"{v:.1f}%" if pd.notna(v) else "新規")

            fig = px.scatter(
                plot_df,
                x="採用店舗数",
                y="昨対比_plot",
                size="売上金額",
                color="象限",
                color_discrete_map=QUAD_COLOR,
                category_orders={"象限": QUAD_ORDER},
                hover_name="商品名",
                hover_data={
                    "JAN": True,
                    "売上金額": ":,.0f",
                    "昨対比_表示": True,
                    "採用店舗数": ":.0f",
                    "昨対比": False,
                    "昨対比_plot": False,
                },
                title="成長マトリクス：採用店舗数 × 昨対比（バブル＝売上金額）",
                size_max=55,
            )
            # 基準線
            fig.add_vline(x=med_store, line_dash="dash", line_color="#888",
                          annotation_text=f"採用中央値 {med_store:.0f}店",
                          annotation_position="top right",
                          annotation_font_color="#888")
            fig.add_hline(y=100, line_dash="dash", line_color="#888",
                          annotation_text="前年同期 100%",
                          annotation_position="bottom right",
                          annotation_font_color="#888")

            # 象限ラベル
            x_max = plot_df["採用店舗数"].quantile(0.97)
            y_min = plot_df["昨対比_plot"].quantile(0.03)
            y_max = plot_df["昨対比_plot"].quantile(0.97)
            for txt, x, y, color in [
                ("🚀 主力成長\n採用広×伸びている",  x_max*0.88, y_max*0.92, "#2ca02c"),
                ("⚠️ 主力低迷\n採用広×伸び悩み",   x_max*0.88, max(y_min*1.1, 80),   "#ff7f0e"),
                ("📈 伸び盛り\n採用拡大チャンス",    med_store*0.15, y_max*0.92, "#1a7a1a"),
                ("❌ 課題商品\n要見直し",             med_store*0.15, max(y_min*1.1, 80), "#d62728"),
            ]:
                fig.add_annotation(x=x, y=y, text=txt, showarrow=False,
                                   font=dict(size=9, color=color), opacity=0.55)

            fig.update_layout(
                height=620,
                yaxis=dict(title="昨対比（%）　※新規商品は0%表示", ticksuffix="%"),
                xaxis_title="採用店舗数",
                legend=dict(title="象限", orientation="h", y=-0.15),
            )
        else:
            # 昨対比がない場合は採用店舗数×1店舗あたり売上の旧チャート
            med_y = jan_sales["1店舗あたり売上"].median()
            fig = px.scatter(
                jan_sales.dropna(subset=["採用店舗数","1店舗あたり売上"]),
                x="採用店舗数", y="1店舗あたり売上",
                size="売上金額",
                color=col_seg if col_seg else col_subcat,
                hover_name="商品名",
                hover_data=["JAN","売上金額"],
                title="定番分析：採用店舗数 × 1店舗あたり売上",
                size_max=50,
            )
            fig.add_vline(x=med_store, line_dash="dash", line_color="gray")
            fig.add_hline(y=med_y,     line_dash="dash", line_color="gray")
            fig.update_layout(height=600)

            jan_sales["象限"] = jan_sales.apply(lambda r: (
                "🚀 主力成長" if r["採用店舗数"] >= med_store and r["1店舗あたり売上"] >= med_y
                else "⚠️ 主力低迷" if r["採用店舗数"] >= med_store
                else "📈 伸び盛り" if r["1店舗あたり売上"] >= med_y
                else "❌ 課題商品"
            ), axis=1)

        st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")

        # 象限サマリー＋詳細テーブル
        QUAD_ORDER_SAFE = ["🚀 主力成長","📈 伸び盛り","⚠️ 主力低迷","❌ 課題商品","🆕 新規商品"]
        quad_sum = (jan_sales.groupby("象限")["売上金額"]
                    .agg(["count","sum"]).reset_index()
                    .rename(columns={"count":"品目数","sum":"売上合計"}))
        quad_sum["象限_order"] = quad_sum["象限"].map({q:i for i,q in enumerate(QUAD_ORDER_SAFE)})
        quad_sum = quad_sum.sort_values("象限_order").drop("象限_order", axis=1)

        c_a, c_b = st.columns([1, 2])
        with c_a:
            st.markdown("**象限別サマリー**")
            st.dataframe(
                quad_sum.style.format({"売上合計":"¥{:,.0f}","品目数":"{:.0f}"}),
                use_container_width=True, hide_index=True,
            )
        with c_b:
            cb_title, cb_input = st.columns([3, 1])
            cb_title.markdown("**商品一覧（売上順）**")
            total_stores = cb_input.number_input(
                "全店舗数", min_value=1, value=383, step=1, key="total_stores",
                label_visibility="visible",
            )

            jan_sales_disp = jan_sales.copy()
            jan_sales_disp["採用率(%)"] = (
                jan_sales_disp["採用店舗数"] / total_stores * 100
            ).round(1)

            show_t = ["JAN","商品名","採用店舗数","採用率(%)","売上金額","1店舗あたり売上","象限"]
            if has_yoy: show_t.insert(6, "昨対比")
            if col_seg and col_seg in jan_sales_disp.columns: show_t.insert(2, col_seg)
            show_t = [c for c in show_t if c in jan_sales_disp.columns]
            fmt_t = {
                "売上金額":"¥{:,.0f}","1店舗あたり売上":"¥{:,.0f}",
                "採用店舗数":"{:.0f}","採用率(%)":"{:.1f}%","昨対比":"{:.1f}%",
            }

            def color_teiban_table(df):
                styles = pd.DataFrame("", index=df.index, columns=df.columns)
                # 行背景（象限色）
                quad_colors = {"🚀 主力成長":"#e8f5e9","📈 伸び盛り":"#f1f8e9",
                               "⚠️ 主力低迷":"#fff3e0","❌ 課題商品":"#ffebee","🆕 新規商品":"#e3f2fd"}
                if "象限" in df.columns:
                    for q, c in quad_colors.items():
                        styles.loc[df["象限"] == q] = f"background-color:{c}"
                # 採用率セルだけ上書きで色付け
                if "採用率(%)" in df.columns:
                    for idx, val in df["採用率(%)"].items():
                        if pd.isna(val): continue
                        if val >= 80:   styles.loc[idx, "採用率(%)"] = "background-color:#c8e6c9; color:#1a7a1a; font-weight:bold"
                        elif val >= 50: styles.loc[idx, "採用率(%)"] = "background-color:#fff9c4; color:#8a6d00; font-weight:bold"
                        else:           styles.loc[idx, "採用率(%)"] = "background-color:#ffcdd2; color:#c62828; font-weight:bold"
                return styles

            make_download_button(jan_sales_disp[show_t].sort_values("売上金額", ascending=False),
                                 f"定番分析_{period_label}.csv")
            st.dataframe(
                jan_sales_disp[show_t].sort_values("売上金額", ascending=False)
                .style
                .format({k:v for k,v in fmt_t.items() if k in show_t}, na_rep="—")
                .apply(color_teiban_table, axis=None),
                use_container_width=True, hide_index=True, height=420,
            )

# ─────────────────────────────────────────────
st.markdown("---")
st.caption(f"IDPOS: {len(df_all):,}件 | 絞り込み後: {len(df_filtered):,}件 | 表示: {len(df_cur):,}件")
