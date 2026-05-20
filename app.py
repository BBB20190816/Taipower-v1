"""
app.py  ── 電業文物 Metadata 分析工具（上傳版）
執行方式：streamlit run app.py
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import io
from pathlib import Path

from data_loader import load_from_uploads

# ── 頁面設定 ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="電業文物 Metadata 分析工具",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 主題 CSS（木質棕 #6B4226 × 溫暖米白 #F7F4EE）────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700&family=Noto+Serif+TC:wght@500;600&display=swap');

html, body, [class*="css"] { font-family: 'Noto Sans TC', sans-serif; }
.stApp { background-color: #F7F4EE; }

h1 {
    font-family: 'Noto Serif TC', serif;
    color: #6B4226;
    border-bottom: 2px solid #B8865A;
    padding-bottom: 10px;
}
h2, h3 { color: #3D2210; }

[data-testid="stSidebar"] {
    background-color: #F7F4EE;
    border-right: 2px solid #6B4226;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    font-size: 9px; font-weight: 700; color: #9A7B60;
    text-transform: uppercase; letter-spacing: .08em; margin-top: 14px;
}
[data-testid="stSidebar"] [data-testid="stCaption"] {
    background: #EDE2D2; border-left: 2px solid #6B4226;
    border-radius: 2px; padding: 5px 9px;
    color: #6B4226; font-weight: 500;
}

[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #D9D1C0;
    border-left: 3px solid #6B4226;
    border-radius: 3px;
    padding: 12px 16px;
    box-shadow: 0 1px 4px rgba(107,66,38,.07);
}
[data-testid="stMetricLabel"] {
    font-size: 10px !important;
    text-transform: uppercase; letter-spacing: .05em; color: #9A7B60 !important;
}
[data-testid="stMetricValue"] { color: #6B4226 !important; font-weight: 700 !important; }

.stTabs [data-baseweb="tab-list"] {
    border-bottom: 1.5px solid #C9BBA8; gap: 2px; background: transparent;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 3px 3px 0 0; color: #9A7B60;
    font-size: 12px; font-weight: 500; padding: 8px 14px; background: transparent;
}
.stTabs [aria-selected="true"] {
    background: #EDE2D2 !important; color: #6B4226 !important;
    font-weight: 700 !important; border-bottom: 2px solid #6B4226 !important;
}

.stButton > button {
    background: #6B4226; color: #FFFFFF; border: none; border-radius: 3px;
    font-family: 'Noto Sans TC', sans-serif; font-weight: 600; font-size: 13px;
    box-shadow: 0 1px 4px rgba(107,66,38,.20);
    transition: opacity .15s, box-shadow .15s;
}
.stButton > button:hover { opacity: .88; box-shadow: 0 2px 8px rgba(107,66,38,.28); }
.stButton > button[kind="secondary"] {
    background: transparent; color: #6B4226; border: 1px solid #C9BBA8;
}

[data-testid="stFileUploadDropzone"] {
    border: 1.5px dashed #B8865A; border-radius: 4px; background: #EDE2D2;
}
[data-testid="stFileUploadDropzone"]:hover { border-color: #6B4226; background: #E5D8C6; }

[data-baseweb="tag"] {
    background: #EDE2D2 !important; color: #6B4226 !important;
    border-radius: 2px !important; border: 1px solid #C9BBA8 !important;
}

[data-baseweb="select"] > div,
[data-baseweb="input"] > div {
    border-color: #C9BBA8 !important; border-radius: 3px !important; background: #FFFFFF !important;
}
[data-baseweb="select"] > div:focus-within,
[data-baseweb="input"] > div:focus-within {
    border-color: #6B4226 !important; box-shadow: 0 0 0 2px rgba(107,66,38,.12) !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid #D9D1C0; border-radius: 3px; overflow: hidden;
    box-shadow: 0 1px 4px rgba(107,66,38,.05);
}

[data-testid="stExpander"] {
    border: 1px solid #D9D1C0 !important; border-radius: 3px !important; background: #FFFFFF;
}
[data-testid="stExpander"] summary { font-weight: 500; color: #6B4226; }

[data-testid="stDownloadButton"] > button {
    background: transparent; color: #6B4226;
    border: 1px solid #B8865A; border-radius: 3px; font-weight: 500;
}
[data-testid="stDownloadButton"] > button:hover { background: #EDE2D2; }

[data-testid="stVerticalBlockBorderWrapper"] > div {
    border: 1px solid #D9D1C0 !important; border-radius: 4px !important;
    background: #FFFFFF; box-shadow: 0 1px 4px rgba(107,66,38,.05);
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# 上傳區（尚未上傳時顯示，上傳後收起）
# ══════════════════════════════════════════════════════════════════
if "df_all" not in st.session_state:
    st.title("🏛️ 電業文物 Metadata 分析工具")
    st.markdown("### 請上傳資料檔案")

    with st.container(border=True):
        col_info, col_upload = st.columns([1, 1])

        with col_info:
            st.markdown("""
**上傳說明**

請一次選取全部 Excel 檔案（主表 + 文資分類表），
app 會自動依檔名識別並整合。

- **主表**（含 `metadata維護` 工作表）
- **文資分類表**（含「建議文資分類」欄位）

📁 支援同時選取多個 .xlsx 檔案
🔒 檔案僅在此 session 中處理，不會儲存或上傳至任何伺服器
""")

        with col_upload:
            uploaded = st.file_uploader(
                "選擇 Excel 檔案",
                type=["xlsx"],
                accept_multiple_files=True,
                label_visibility="collapsed",
            )

            if uploaded:
                main_count = sum(
                    1 for f in uploaded
                    if not any(kw in f.name for kw in ["建議文資分類", "圖片資訊"])
                )
                class_count = len(uploaded) - main_count
                st.caption(f"已選取 **{len(uploaded)}** 個檔案（主表 {main_count} 個，分類表 {class_count} 個）")

                if st.button("🚀 開始載入並整合資料", type="primary", use_container_width=True):
                    with st.spinner("整合中，約需 10~15 秒…"):
                        try:
                            df = load_from_uploads(uploaded)
                            st.session_state["df_all"] = df
                            st.rerun()
                        except Exception as e:
                            st.error(f"載入失敗：{e}")
            else:
                st.info("⬆️ 點擊上方或拖曳 xlsx 檔案至此區域")

    st.stop()

# ── 資料已載入 ────────────────────────────────────────────────────────
df_all = st.session_state["df_all"]

# 側欄重置按鈕
with st.sidebar:
    if st.button("🔄 重新上傳資料", use_container_width=True):
        del st.session_state["df_all"]
        st.rerun()
    st.markdown("---")

# ── 側邊欄篩選 ────────────────────────────────────────────────────────
st.sidebar.title("🔍 篩選條件")
st.sidebar.caption(f"資料總筆數：{len(df_all):,} 件")

def ms_ex(label, col):
    """multiselect，標籤顯示選項數，tooltip 預覽所有值。"""
    opts = sorted(df_all[col].dropna().unique().tolist())
    preview = "、".join(str(o) for o in opts[:20])
    if len(opts) > 20:
        preview += f"… 等 {len(opts)} 種"
    return st.multiselect(
        f"{label}（{len(opts)} 種）",
        opts, default=[],
        help=preview,
    )

# ── Group 1：批次與單位 ──────────────────────────────────────────────
with st.sidebar.expander("📦 批次與單位", expanded=True):
    batches = ms_ex("審查批次", "批次")
    owners  = ms_ex("典藏單位", "典藏單位")
    systems = ms_ex("系統別",   "系統別")

# ── Group 2：典藏分類 ────────────────────────────────────────────────
with st.sidebar.expander("🗂️ 典藏分類"):
    types      = ms_ex("典藏類型",   "典藏類型")
    subtypes   = ms_ex("典藏次類型", "典藏次類型")
    originals  = ms_ex("原件與否",   "原件與否")
    agg_levels = ms_ex("藏品層次",   "藏品層次")

# ── Group 3：主題與屬性 ──────────────────────────────────────────────
with st.sidebar.expander("🔖 主題與屬性"):
    subjects  = ms_ex("電業主題", "電業主題")
    attrs     = ms_ex("文物屬性", "文物屬性")
    materials = ms_ex("主要材質", "主要材質")
    events    = ms_ex("事件名稱", "事件名稱")

# ── Group 4：狀態與管理 ──────────────────────────────────────────────
with st.sidebar.expander("✅ 狀態與管理"):
    conditions   = ms_ex("保存狀況",     "保存狀況")
    categories   = ms_ex("建議分類",     "建議分類")
    levels       = ms_ex("分級",         "分級")
    restrictions = ms_ex("使用限制",     "使用限制")
    open_flags   = ms_ex("資料開放狀態", "資料開放狀態")

# ── 已啟用篩選條件摘要 ───────────────────────────────────────────────
_active = sum(1 for v in [
    batches, owners, systems, types, subtypes, originals, agg_levels,
    subjects, attrs, materials, events,
    conditions, categories, levels, restrictions, open_flags,
] if v)
if _active:
    st.sidebar.info(f"⚡ 已啟用 **{_active}** 個篩選條件")

st.sidebar.markdown("---")
st.sidebar.markdown("### 文物年代（起始西元年）")
yr_min = int(df_all["起始西元年"].dropna().min())
yr_max = int(df_all["起始西元年"].dropna().max())
yr_range = st.sidebar.slider("年代範圍", yr_min, yr_max, (yr_min, yr_max))

st.sidebar.markdown("---")
st.sidebar.markdown("### 📐 尺寸篩選")
st.sidebar.caption("單位：公分")
use_dim = st.sidebar.checkbox("啟用尺寸篩選", value=False)
if use_dim:
    l_min_all = float(df_all["長"].dropna().min())
    l_max_all = float(df_all["長"].dropna().max())
    w_min_all = float(df_all["寬"].dropna().min())
    w_max_all = float(df_all["寬"].dropna().max())
    l_range = st.sidebar.slider("長（cm）", l_min_all, l_max_all,
                                (l_min_all, l_max_all), step=0.5)
    w_range = st.sidebar.slider("寬（cm）", w_min_all, w_max_all,
                                (w_min_all, w_max_all), step=0.5)
    use_h = df_all["高厚"].notna().any() and st.sidebar.checkbox("也篩選高（厚）")
    if use_h:
        h_min_all = float(df_all["高厚"].dropna().min())
        h_max_all = float(df_all["高厚"].dropna().max())
        h_range = st.sidebar.slider("高/厚（cm）", h_min_all, h_max_all,
                                    (h_min_all, h_max_all), step=0.5)
    else:
        h_range = None

# ── 套用篩選 ─────────────────────────────────────────────────────────
def af(df, col, chosen):
    return df[df[col].isin(chosen)] if chosen else df

df = df_all.copy()
for col, chosen in [
    ("批次", batches), ("典藏單位", owners), ("系統別", systems),
    ("典藏類型", types), ("典藏次類型", subtypes),
    ("原件與否", originals), ("藏品層次", agg_levels),
    ("電業主題", subjects), ("文物屬性", attrs),
    ("主要材質", materials), ("事件名稱", events),
    ("保存狀況", conditions), ("建議分類", categories), ("分級", levels),
    ("使用限制", restrictions), ("資料開放狀態", open_flags),
]:
    df = af(df, col, chosen)

mask_yr = df["起始西元年"].isna() | (
    (df["起始西元年"] >= yr_range[0]) & (df["起始西元年"] <= yr_range[1])
)
df = df[mask_yr]

if use_dim:
    mask_dim = (
        df["長"].notna() & df["寬"].notna() &
        df["長"].between(l_range[0], l_range[1]) &
        df["寬"].between(w_range[0], w_range[1])
    )
    if use_h and h_range:
        mask_dim &= df["高厚"].notna() & df["高厚"].between(h_range[0], h_range[1])
    df = df[mask_dim]

# ── 主畫面 ────────────────────────────────────────────────────────────
st.title("🏛️ 電業文物 Metadata 分析工具")

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 資料總覽", "🔎 搜尋與清單", "📐 尺寸分析", "↔️ 交叉比對", "📥 匯出"]
)

# ══════════════════════════════════════════════════════════════════
# TAB 1：資料總覽
# ══════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("篩選結果摘要")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("符合件數", f"{len(df):,}")
    c2.metric("佔總件數比例", f"{len(df)/len(df_all)*100:.1f}%")
    c3.metric("涵蓋批次數", df["批次"].nunique())
    c4.metric("典藏類型數", df["典藏類型"].nunique())

    st.markdown("---")
    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown("**各審查批次件數**")
        bc = df["批次"].value_counts().reset_index()
        bc.columns = ["批次", "件數"]
        fig = px.bar(bc, x="件數", y="批次", orientation="h",
                     text="件數", color_discrete_sequence=["#5B8DEF"])
        fig.update_traces(textposition="outside", texttemplate="%{text:,}")
        fig.update_layout(height=300, margin=dict(l=0,r=40,t=20,b=0), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("**保存狀況分布**")
        cc = df["保存狀況"].value_counts().reset_index()
        cc.columns = ["保存狀況", "件數"]
        fig2 = px.pie(cc, names="保存狀況", values="件數",
                      color_discrete_sequence=px.colors.qualitative.Pastel)
        fig2.update_traces(textinfo="label+value+percent",
                           texttemplate="%{label}<br>%{value:,} 件 (%{percent})")
        fig2.update_layout(height=300, margin=dict(l=0,r=0,t=20,b=0),
                           showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    with col_r:
        st.markdown("**建議分類分布**")
        catc = df["建議分類"].value_counts().reset_index()
        catc.columns = ["建議分類", "件數"]
        fig3 = px.bar(catc, x="建議分類", y="件數",
                      text="件數", color_discrete_sequence=["#F4845F"])
        fig3.update_traces(textposition="outside", texttemplate="%{text:,}")
        fig3.update_layout(height=300, margin=dict(l=0,r=0,t=30,b=0), xaxis_tickangle=-30)
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("**文物年代分布（起始西元年）**")
        yr_data = df["起始西元年"].dropna()
        if len(yr_data) > 0:
            fig4 = px.histogram(yr_data, nbins=30, color_discrete_sequence=["#7EC8A4"],
                                text_auto=True)
            fig4.update_traces(texttemplate="%{y:,}", textposition="outside")
            fig4.update_layout(height=300, margin=dict(l=0,r=0,t=30,b=0),
                               xaxis_title="西元年", yaxis_title="件數")
            st.plotly_chart(fig4, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# TAB 2：搜尋與清單
# ══════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("🔎 關鍵字搜尋")

    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        keyword = st.text_input("搜尋關鍵字",
            placeholder="輸入文物名稱、登錄號、關鍵詞、描述…",
            label_visibility="collapsed")
    with col_s2:
        search_fields = st.multiselect(
            "搜尋範圍",
            options=["文物名稱", "文物登錄號", "關鍵詞", "文物綜合簡述",
                     "文化意義", "備註", "別稱", "事件名稱"],
            default=["文物名稱", "文物登錄號", "關鍵詞"],
            label_visibility="collapsed",
        )

    df_search = df.copy()
    if keyword.strip():
        kw = keyword.strip()
        mask = pd.Series(False, index=df_search.index)
        for field in search_fields:
            if field in df_search.columns:
                mask |= df_search[field].astype(str).str.contains(kw, case=False, na=False)
        df_search = df_search[mask]
        st.caption(f"搜尋「**{kw}**」，符合 **{len(df_search):,}** 件")
    else:
        st.caption(f"目前篩選結果：**{len(df_search):,}** 件")

    default_cols = [
        "文物登錄號", "文物名稱", "批次", "典藏類型", "典藏次類型",
        "電業主題", "文物屬性", "保存狀況", "起始西元年", "年號",
        "典藏單位", "主要材質", "關鍵詞", "文物年代類型",
        "建議分類", "分級", "使用限制", "資料開放狀態",
    ]
    chosen_cols = st.multiselect(
        "顯示欄位",
        options=df_search.columns.tolist(),
        default=[c for c in default_cols if c in df_search.columns],
    )
    if chosen_cols:
        st.dataframe(df_search[chosen_cols].reset_index(drop=True),
                     use_container_width=True, height=500)

# ══════════════════════════════════════════════════════════════════
# TAB 3：尺寸分析
# ══════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("📐 尺寸範圍查詢")
    st.markdown("設定長寬範圍，查看符合的文物清單。")

    dim_all = df_all[df_all["長"].notna() & df_all["寬"].notna()]
    dim_filtered = df[df["長"].notna() | df["寬"].notna()].copy()

    if len(dim_all) == 0:
        st.warning("資料中無尺寸資料。")
    else:
        lmin, lmax = float(dim_all["長"].min()), float(dim_all["長"].max())
        wmin, wmax = float(dim_all["寬"].min()), float(dim_all["寬"].max())

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**長（cm）**")
            l_lo = st.number_input("最小（長）", min_value=lmin, max_value=lmax,
                                   value=lmin, step=1.0, key="l_lo")
            l_hi = st.number_input("最大（長）", min_value=lmin, max_value=lmax,
                                   value=lmax, step=1.0, key="l_hi")
        with c2:
            st.markdown("**寬（cm）**")
            w_lo = st.number_input("最小（寬）", min_value=wmin, max_value=wmax,
                                   value=wmin, step=1.0, key="w_lo")
            w_hi = st.number_input("最大（寬）", min_value=wmin, max_value=wmax,
                                   value=wmax, step=1.0, key="w_hi")

        use_h2 = False
        if dim_all["高厚"].notna().any():
            use_h2 = st.checkbox("同時篩選高（厚）", key="use_h2")
            if use_h2:
                hmin = float(dim_all["高厚"].dropna().min())
                hmax = float(dim_all["高厚"].dropna().max())
                h_lo = st.number_input("最小（高/厚）", min_value=hmin, max_value=hmax,
                                       value=hmin, step=1.0, key="h_lo")
                h_hi = st.number_input("最大（高/厚）", min_value=hmin, max_value=hmax,
                                       value=hmax, step=1.0, key="h_hi")

        dim_result = dim_filtered[
            dim_filtered["長"].notna() & dim_filtered["寬"].notna() &
            (dim_filtered["長"] >= l_lo) & (dim_filtered["長"] <= l_hi) &
            (dim_filtered["寬"] >= w_lo) & (dim_filtered["寬"] <= w_hi)
        ].copy()
        if use_h2:
            dim_result = dim_result[
                dim_result["高厚"].notna() &
                (dim_result["高厚"] >= h_lo) & (dim_result["高厚"] <= h_hi)
            ]

        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("符合件數", f"{len(dim_result):,}")
        m2.metric("長 平均", f"{dim_result['長'].mean():.1f} cm" if len(dim_result) else "—")
        m3.metric("寬 平均", f"{dim_result['寬'].mean():.1f} cm" if len(dim_result) else "—")
        m4.metric("高厚 有值件數", f"{dim_result['高厚'].notna().sum():,}")

        if len(dim_result) == 0:
            st.info("無符合條件的文物，請調整尺寸範圍。")
        else:
            st.markdown("**長 × 寬 散佈圖**")
            fig_s = px.scatter(dim_result, x="長", y="寬", color="批次",
                               hover_data=["文物登錄號", "文物名稱", "典藏類型", "建議分類"],
                               color_discrete_sequence=px.colors.qualitative.Safe,
                               labels={"長": "長 (cm)", "寬": "寬 (cm)"})
            fig_s.update_layout(height=400, margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig_s, use_container_width=True)

            show_cols = [c for c in ["文物登錄號", "文物名稱", "批次", "長", "寬", "高厚",
                                     "數量單位", "典藏類型", "建議分類", "保存狀況"]
                         if c in dim_result.columns]
            st.dataframe(dim_result[show_cols].reset_index(drop=True),
                         use_container_width=True, height=400)

            csv = dim_result[show_cols].to_csv(index=False, encoding="utf-8-sig")
            st.download_button("⬇️ 下載此清單（CSV）",
                               data=csv.encode("utf-8-sig"),
                               file_name="尺寸篩選結果.csv", mime="text/csv")

# ══════════════════════════════════════════════════════════════════
# TAB 4：交叉比對
# ══════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("↔️ 兩欄位交叉比對")
    cat_cols = [
        "批次", "典藏類型", "典藏次類型", "原件與否", "藏品層次", "系統別",
        "電業主題", "文物屬性", "保存狀況", "文物年代類型",
        "建議分類", "分級", "典藏單位", "來源類型", "資料開放狀態",
        "主要材質", "數量單位",
    ]
    available = [c for c in cat_cols if c in df.columns]

    col_a, col_b = st.columns(2)
    with col_a:
        x_col = st.selectbox("X 軸欄位", available,
                             index=available.index("保存狀況") if "保存狀況" in available else 0)
    with col_b:
        rest = [c for c in available if c != x_col]
        y_col = st.selectbox("Y 軸欄位", rest,
                             index=rest.index("建議分類") if "建議分類" in rest else 0)

    if x_col and y_col:
        pivot = df.groupby([x_col, y_col], dropna=False).size().reset_index(name="件數")
        pt = pivot.pivot_table(index=x_col, columns=y_col, values="件數", fill_value=0)

        fig_h = px.imshow(pt, text_auto=True, aspect="auto",
                          color_continuous_scale="Blues", labels={"color": "件數"})
        fig_h.update_layout(height=max(300, len(pt)*45+100),
                            margin=dict(l=0,r=0,t=30,b=0),
                            xaxis_title=y_col, yaxis_title=x_col)
        st.plotly_chart(fig_h, use_container_width=True)

        with st.expander("查看數字表格"):
            st.dataframe(pt, use_container_width=True)

        fig_b = px.bar(pivot.dropna(subset=[x_col, y_col]),
                       x=x_col, y="件數", color=y_col, barmode="stack",
                       color_discrete_sequence=px.colors.qualitative.Safe)
        fig_b.update_layout(height=400, margin=dict(l=0,r=0,t=20,b=0),
                            xaxis_tickangle=-30, legend_title=y_col)
        st.plotly_chart(fig_b, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# TAB 5：匯出
# ══════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("📥 匯出篩選結果")
    st.write(f"目前篩選結果：**{len(df):,} 件**")

    default_exp = [
        "文物登錄號", "文物名稱", "批次", "典藏類型", "典藏次類型",
        "電業主題", "文物屬性", "保存狀況", "起始西元年", "年號",
        "長", "寬", "高厚", "主要材質", "數量單位",
        "建議分類", "分級", "典藏單位",
    ]
    export_cols = st.multiselect(
        "選擇匯出欄位",
        options=df.columns.tolist(),
        default=[c for c in default_exp if c in df.columns],
    )
    if export_cols:
        c1, c2 = st.columns(2)
        with c1:
            csv_buf = io.StringIO()
            df[export_cols].to_csv(csv_buf, index=False, encoding="utf-8-sig")
            st.download_button("⬇️ 下載 CSV",
                               data=csv_buf.getvalue().encode("utf-8-sig"),
                               file_name="文物metadata篩選結果.csv", mime="text/csv")
        with c2:
            xl_buf = io.BytesIO()
            with pd.ExcelWriter(xl_buf, engine="openpyxl") as writer:
                df[export_cols].to_excel(writer, index=False, sheet_name="篩選結果")
            st.download_button("⬇️ 下載 Excel",
                               data=xl_buf.getvalue(),
                               file_name="文物metadata篩選結果.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
