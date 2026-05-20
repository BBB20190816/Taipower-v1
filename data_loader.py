"""
data_loader.py
支援兩種模式：
  - load_all(base_dir)        : 從本機目錄讀取（開發用）
  - load_from_uploads(files)  : 從 Streamlit file_uploader 物件讀取（部署用）
"""
import pandas as pd
from pathlib import Path
from io import BytesIO

# ── 欄位索引 ────────────────────────────────────────────────────────
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

# 分類表的辨識關鍵字（出現在檔名中即判定為分類表）
CLASS_KEYWORDS = ["建議文資分類", "圖片資訊及建議"]

# 本機模式的批次對照表（依檔名）
BATCH_MAP = {
    "111年_metadata_006電力調度處_53案":          "111年-電力調度處(53案)",
    "111年metadata_014秘書處_532案":              "111年-秘書處(532案)",
    "111年metadata_231蘭陽發電廠_437案":          "111年-蘭陽發電廠(437案)",
    "112年第2次審查_詮釋資料metadata_303案":      "112年-第2次審查(303案)",
    "112年第2次審查_詮釋資料2023metadata_電源開發處97案": "112年-電源開發處(97案)",
    "113年第3次審查_metadata_400案_final":        "113年-第3次審查(400案)",
    "113年第4次審查_metadata_400案_修正版":       "113年-第4次審查(400案)",
    "113年第5期審查metadata":                     "113年-第5期審查",
}


def _guess_batch(filename: str) -> str:
    """從檔名猜測批次標籤。"""
    for key, label in BATCH_MAP.items():
        if key in filename:
            return label
    # 找不到對照時，用檔名本身（去掉副檔名）
    return Path(filename).stem


def _is_class_file(filename: str) -> bool:
    """判斷是否為文資分類表。"""
    return any(kw in filename for kw in CLASS_KEYWORDS)


def _parse_main_bytes(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    解析一個主表 Excel（bytes）。
    回傳：每件文物一列的 DataFrame。
    """
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name="metadata維護", header=None)
    data = raw.iloc[3:].reset_index(drop=True)
    id_col = COL["文物登錄號"]

    id_series = data.iloc[:, id_col]
    data_with_id = data[id_series.notna()].copy()
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


def _parse_class_bytes(file_bytes: bytes) -> pd.DataFrame:
    """解析一個文資分類表（bytes），回傳 登錄號→建議分類、分級。"""
    xl = pd.ExcelFile(BytesIO(file_bytes))
    # 取第一個含「文資」或「分類」的工作表，找不到就取第一個
    target = xl.sheet_names[0]
    for s in xl.sheet_names:
        if any(kw in s for kw in ["文資", "分類"]):
            target = s
            break

    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target)
    if "分級" not in df.columns:
        df["分級"] = None
    # 確保欄位存在
    needed = [c for c in ["文物登錄號", "建議分類", "分級"] if c in df.columns]
    return df[needed]


def _finalize(master: pd.DataFrame, class_frames: list) -> pd.DataFrame:
    """合併分類表、清理型別與值，回傳最終 DataFrame。"""
    if class_frames:
        class_df = pd.concat(class_frames, ignore_index=True)
        class_df = class_df.drop_duplicates(subset="文物登錄號", keep="last")
        master = master.merge(class_df, on="文物登錄號", how="left")

    for col in ["起始西元年", "結束西元年", "照片數", "長", "寬", "高厚", "直徑", "重量"]:
        if col in master.columns:
            master[col] = pd.to_numeric(master[col], errors="coerce")

    if "建議分類" in master.columns:
        master["建議分類"] = master["建議分類"].str.strip().replace({"列冊追踪": "列冊追蹤"})
    if "分級" in master.columns:
        master["分級"] = master["分級"].str.strip().replace(
            {"公司級": "Company.公司級文物", "單位級": "Unit.單位級文物"}
        )

    return master.drop_duplicates(subset="文物登錄號", keep="last").reset_index(drop=True)


# ── 公開 API ────────────────────────────────────────────────────────

def load_from_uploads(uploaded_files) -> pd.DataFrame:
    """
    從 Streamlit file_uploader 的結果讀取並整合。
    uploaded_files: list of UploadedFile（st.file_uploader 回傳值）
    """
    main_frames = []
    class_frames = []

    for f in uploaded_files:
        file_bytes = f.read()
        if _is_class_file(f.name):
            try:
                class_frames.append(_parse_class_bytes(file_bytes))
            except Exception:
                pass  # 分類表解析失敗不中斷
        else:
            try:
                xl = pd.ExcelFile(BytesIO(file_bytes))
                if "metadata維護" in xl.sheet_names:
                    main_frames.append(_parse_main_bytes(file_bytes, f.name))
            except Exception:
                pass  # 非主表 xlsx 略過

    if not main_frames:
        raise ValueError("找不到有效的主表（需含「metadata維護」工作表）。請確認上傳了正確的檔案。")

    master = pd.concat(main_frames, ignore_index=True)
    return _finalize(master, class_frames)


def load_all(base_dir: str = ".") -> pd.DataFrame:
    """從本機目錄讀取（本機開發用）。"""
    base = Path(base_dir)
    main_frames = []
    class_frames = []

    for fpath in sorted(base.glob("*.xlsx")):
        file_bytes = fpath.read_bytes()
        if _is_class_file(fpath.name):
            try:
                class_frames.append(_parse_class_bytes(file_bytes))
            except Exception:
                pass
        else:
            try:
                xl = pd.ExcelFile(BytesIO(file_bytes))
                if "metadata維護" in xl.sheet_names:
                    main_frames.append(_parse_main_bytes(file_bytes, fpath.name))
            except Exception:
                pass

    if not main_frames:
        raise FileNotFoundError(f"在 {base_dir} 找不到任何主表檔案。")

    master = pd.concat(main_frames, ignore_index=True)
    return _finalize(master, class_frames)
