# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 執行方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

部署平台：Streamlit Community Cloud（`taipower-metadata.streamlit.app`），監聽 `main` 分支，push 後自動重新部署。

## 架構概覽

兩個模組，職責明確分離：

- **`data_loader.py`**：純資料層。解析 Excel、合併分類表、標準化欄位值。不依賴 Streamlit。
- **`app.py`**：純呈現層。從 `st.session_state["df_all"]` 取得已整合的 DataFrame，負責篩選、圖表、匯出。

### 資料流

```
使用者上傳 xlsx（多檔）
  → load_from_uploads()         # 識別主表 vs 分類表
  → _parse_main_bytes()         # 跳過前3列標頭，col55為key，計算照片數
  → _parse_class_bytes()        # 取登錄號→建議分類、分級
  → _finalize()                 # merge、型別轉換、值標準化
  → st.session_state["df_all"] # 整個 session 共用
```

### 欄位索引（`COL` 字典，0-based）

主表 Excel 結構：第 0 列為序號，第 1 列起為欄位（`row.iloc[col_index]`）。資料從第 4 列（`raw.iloc[3:]`）開始。

關鍵索引：`文物登錄號=55`、`典藏單位=81`、`資料開放狀態=82`、`長=47`、`寬=48`、`高厚=49`。

### 檔案識別規則

- 檔名含 `建議文資分類` 或 `圖片資訊及建議` → 分類表
- 含 `metadata維護` 工作表 → 主表
- 其餘略過

### 值標準化（`_finalize` 中）

| 原始值 | 標準化後 |
|--------|----------|
| 列冊追踪 | 列冊追蹤 |
| 公司級 | Company.公司級文物 |
| 單位級 | Unit.單位級文物 |

## 側欄篩選結構

16 個 multiselect 分成 4 個 `st.sidebar.expander` 群組，套用篩選的迴圈在 `app.py` 的「套用篩選」區段，欄位名稱須與 `COL` 字典的 key 完全一致。

新增篩選欄位時需同步修改兩處：
1. expander 群組內新增 `ms_ex()`
2. 套用篩選的 `for col, chosen in [...]` 清單

## 注意事項

- 資料檔（xlsx）不進 repo，`.gitignore` 已排除。
- `load_all(base_dir)` 僅供本機開發，部署版只用 `load_from_uploads()`。
- `ms_ex()` 函式的 `help` tooltip 最多預覽 20 個選項值。
