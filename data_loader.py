"""
data_loader.py
支援三種模式：
  - load_all(base_dir)        : 從本機目錄讀取（開發用）
  - load_from_uploads(files)  : 從 Streamlit file_uploader 物件讀取
    - 模式 A：上傳 1 個整合檔（merged_metadata_v3 格式）
    - 模式 B：上傳 15 個原始 xlsx（主表 + 文資分類表）
"""
import pandas as pd
from pathlib import Path
from io import BytesIO

# ── 欄位索引（原始主表 metadata維護 結構）────────────────────────────
COL = {
    "典藏類型": 1, "典藏次類型": 3, "原件與否": 5, "藏品層次": 6, "系統別": 7,
    "文物名稱": 8, "別稱": 9, "創作者角色": 10, "創作者姓名": 11,
    "電業主題": 13, "關鍵詞": 14, "文物屬性": 15,
    "文物綜合簡述": 16, "文化意義": 17, "事件名稱": 18,
    "保存狀況": 19, "市場鑑價": 20, "主要材質": 21, "次要材質": 22,
    "文物鑑定狀態": 23, "藏品分級": 24, "備註": 25,
    "出版社": 26, "出版地": 27, "西元出版年": 28,
    "文物年代類型": 29, "起始西元年": 30, "起始西元月": 31, "起始西元日": 32,
    "結束西元年": 33, "結束西元月": 34, "結束西元日": 35, "年號": 36,
    "地點類型": 37, "地點": 38, "地址": 39,
    "保存環境": 44, "縣市": 45, "數量單位": 46,
    "長": 47, "寬": 48, "高厚": 49, "直徑": 50, "重量": 51,
    "詮釋資料識別碼": 52, "國家文化資料庫識別號": 53, "文物普查編號": 54,
    "文物登錄號": 55,
    "來源類型": 56, "清查日期": 57, "來源對象類別": 58,
    "來源對象名稱": 59, "來源對象地址": 60,
    "全集系列名稱": 64, "子項組件": 65, "專案名稱": 66, "專案計畫編號": 67,
    "數位化類別": 68,
    "著作財產權人": 78, "授權狀態": 79, "使用限制": 80,
    "典藏單位": 81, "資料開放狀態": 82,
}

# ── 批次名稱對照（整合檔原始檔名 → 友善標籤）────────────────────────
BATCH_MAP = {
    "111年_metadata_006電力調度處_53案":          "111年-電力調度處(53案)",
    "111年metadata_014秘書處_532案":              "111年-秘書處(532案)",
    "111年metadata_231蘭陽發電廠_437案":          "111年-蘭陽發電廠(437案)",
    "112年第2次審查__詮釋資料_metadata_303案":    "112年-第2次審查(303案)",
    "112年第2次審查__詮釋資料_2023metadata_電源開發處97案": "112年-電源開發處(97案)",
    "113年第3次審查_metadata_400案_final":        "113年-第3次審查(400案)",
    "113年第4次審查_metadata_400案_修正版":       "113年-第4次審查(400案)",
    "113年第5期審查metadata":                     "113年-第5期審查",
    # 原始主表檔名對照（模式 B）
    "111年_metadata_006電力調度處_53案.xlsx":                                     "111年-電力調度處(53案)",
    "111年metadata_014秘書處_532案_更新最後107案426532.xlsx":                     "111年-秘書處(532案)",
    "111年metadata_231蘭陽發電廠_437案.xlsx":                                     "111年-蘭陽發電廠(437案)",
    "112年第2次審查_詮釋資料metadata_303案_修_定稿0401.xlsx":                     "112年-第2次審查(303案)",
    "112年第2次審查_詮釋資料2023metadata_電源開發處97案__0116更新_FINAL_定稿.xlsx": "112年-電源開發處(97案)",
    "113年第3次審查_metadata_400案_final_系統上傳版本.xlsx":                      "113年-第3次審查(400案)",
    "113年第4次審查_metadata_400案_修正版_更新0401.xlsx":                         "113年-第4次審查(400案)",
    "113年第5期審查metadata_0428修改0512.xlsx":                                   "113年-第5期審查",
}

CLASS_KEYWORDS = ["建議文資分類", "圖片資訊及建議"]


def _guess_batch(filename: str) -> str:
    """從檔名猜測友善批次標籤。"""
    for key, label in BATCH_MAP.items():
        if key in filename:
            return label
    return Path(filename).stem


def _is_class_file(filename: str) -> bool:
    return any(kw in filename for kw in CLASS_KEYWORDS)


def _is_merged_file(file_bytes: bytes) -> bool:
    """判斷是否為整合檔：metadata維護 工作表有 87 欄（原始主表只有 83 欄）。"""
    try:
        raw = pd.read_excel(BytesIO(file_bytes), sheet_name="metadata維護", header=None, nrows=1)
        return len(raw.columns) > 83
    except Exception:
        return False


def _normalize(master: pd.DataFrame) -> pd.DataFrame:
    """統一值標準化與型別整理。"""
    for col in ["起始西元年", "結束西元年", "照片數", "長", "寬", "高厚", "直徑", "重量"]:
        if col in master.columns:
            master[col] = pd.to_numeric(master[col], errors="coerce")

    if "建議分類" in master.columns:
        master["建議分類"] = master["建議分類"].str.strip().replace({
            "列冊追踪": "列冊追蹤",
            "（無對應資料）": None,
        })
    if "分級" in master.columns:
        master["分級"] = master["分級"].str.strip().replace({
            "公司級": "Company.公司級文物",
            "單位級": "Unit.單位級文物",
            "（無對應資料）": None,
        })
    if "批次" in master.columns:
        master["批次"] = master["批次"].map(
            lambda x: _guess_batch(str(x)) if pd.notna(x) else x
        )

    return master.drop_duplicates(subset="文物登錄號", keep="last").reset_index(drop=True)


# ── 模式 A：整合檔解析 ───────────────────────────────────────────────
def _parse_merged(file_bytes: bytes) -> pd.DataFrame:
    """
    解析整合檔（merged_metadata_v3 格式）。
    結構：前3列為標頭，第4列起為資料，每件文物有多列（照片），
    col55=文物登錄號，col84=批次，col85=建議分類，col86=分級。
    """
    raw = pd.read_excel(BytesIO(file_bytes), header=None, skiprows=3)
    id_col = 55

    # 只保留登錄號有值的列，計算照片數，取每個登錄號第一列
    data = raw[raw.iloc[:, id_col].notna()].copy()
    photo_counts = data.iloc[:, id_col].value_counts()
    first_rows = data.drop_duplicates(subset=[data.columns[id_col]], keep="first")

    rows = []
    for _, row in first_rows.iterrows():
        rec = {}
        for name, ci in COL.items():
            rec[name] = row.iloc[ci] if ci < len(row) else None
        rec["批次"]    = row.iloc[84] if 84 < len(row) else None
        rec["建議分類"] = row.iloc[85] if 85 < len(row) else None
        rec["分級"]    = row.iloc[86] if 86 < len(row) else None
        rec["照片數"]  = int(photo_counts.get(row.iloc[id_col], 1))
        rows.append(rec)

    return pd.DataFrame(rows)


# ── 模式 B：原始主表解析 ─────────────────────────────────────────────
def _parse_main(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """解析一個原始主表 xlsx（含 metadata維護 工作表）。"""
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name="metadata維護", header=None)
    data = raw.iloc[3:].reset_index(drop=True)
    id_col = COL["文物登錄號"]

    data_with_id = data[data.iloc[:, id_col].notna()].copy()
    photo_counts = data_with_id.iloc[:, id_col].value_counts()
    first_rows = data_with_id.drop_duplicates(
        subset=[data_with_id.columns[id_col]], keep="first"
    )

    batch = _guess_batch(filename)
    rows = []
    for _, row in first_rows.iterrows():
        rec = {"批次": batch}
        for name, ci in COL.items():
            rec[name] = row.iloc[ci] if ci < len(row) else None
        rec["照片數"] = int(photo_counts.get(row.iloc[id_col], 1))
        rows.append(rec)

    return pd.DataFrame(rows)


def _parse_class(file_bytes: bytes) -> pd.DataFrame:
    """解析文資分類表。"""
    xl = pd.ExcelFile(BytesIO(file_bytes))
    sheet = xl.sheet_names[0]
    for s in xl.sheet_names:
        if any(kw in s for kw in ["文資", "分類"]):
            sheet = s
            break
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)
    if "分級" not in df.columns:
        df["分級"] = None
    needed = [c for c in ["文物登錄號", "建議分類", "分級"] if c in df.columns]
    return df[needed]


# ── 公開 API ─────────────────────────────────────────────────────────
def load_from_uploads(uploaded_files) -> pd.DataFrame:
    """
    從 Streamlit file_uploader 結果自動判斷模式並整合。
    模式 A：只有 1 個非分類表 xlsx，且沒有 metadata維護 工作表 → 整合檔
    模式 B：多個 xlsx，含 metadata維護 工作表 → 原始主表 + 分類表
    """
    # 先讀取所有檔案的 bytes
    file_data = []
    for f in uploaded_files:
        file_data.append((f.name, f.read()))

    # 區分整合檔 / 主表 / 分類表
    merged_files = []
    main_files   = []
    class_files  = []

    for fname, fbytes in file_data:
        if _is_class_file(fname):
            class_files.append(fbytes)
        elif _is_merged_file(fbytes):
            merged_files.append(fbytes)
        else:
            try:
                xl = pd.ExcelFile(BytesIO(fbytes))
                if "metadata維護" in xl.sheet_names:
                    main_files.append((fname, fbytes))
            except Exception:
                pass

    # 模式 A：整合檔
    if merged_files:
        master = _parse_merged(merged_files[0])  # 只取第一個整合檔
        return _normalize(master)

    # 模式 B：原始主表 + 分類表
    if main_files:
        frames = [_parse_main(fb, fn) for fn, fb in main_files]
        master = pd.concat(frames, ignore_index=True)

        if class_files:
            class_df = pd.concat(
                [_parse_class(fb) for fb in class_files], ignore_index=True
            ).drop_duplicates(subset="文物登錄號", keep="last")
            master = master.merge(class_df, on="文物登錄號", how="left")

        return _normalize(master)

    raise ValueError(
        "找不到有效的資料檔案。\n"
        "請上傳「整合版 xlsx」或「原始主表 xlsx（含 metadata維護 工作表）」。"
    )


def load_all(base_dir: str = ".") -> pd.DataFrame:
    """從本機目錄讀取（開發用）。"""
    base = Path(base_dir)
    main_frames  = []
    class_frames = []

    for fpath in sorted(base.glob("*.xlsx")):
        fbytes = fpath.read_bytes()
        if _is_class_file(fpath.name):
            try:
                class_frames.append(_parse_class(fbytes))
            except Exception:
                pass
        elif not _is_merged_file(fbytes):
            try:
                xl = pd.ExcelFile(BytesIO(fbytes))
                if "metadata維護" in xl.sheet_names:
                    main_frames.append(_parse_main(fbytes, fpath.name))
            except Exception:
                pass

    if not main_frames:
        raise FileNotFoundError(f"在 {base_dir} 找不到主表檔案。")

    master = pd.concat(main_frames, ignore_index=True)
    if class_frames:
        class_df = pd.concat(class_frames, ignore_index=True).drop_duplicates(
            subset="文物登錄號", keep="last"
        )
        master = master.merge(class_df, on="文物登錄號", how="left")

    return _normalize(master)
