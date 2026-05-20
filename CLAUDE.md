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

- 檔名含 `CLASS_KEYWORDS`（`建議文資分類`、`圖片資訊及建議`）→ 分類表
- 含 `metadata維護` 工作表 → 主表
- 其餘略過（靜默，不報錯）

`_parse_class_bytes` 會優先選取名稱含「文資」或「分類」的工作表；若無符合者，退回第一個工作表（`sheet_names[0]`）。因此分類表的工作表名稱不限定，但須確認第一個工作表即為資料表。

### 值標準化（`_finalize` 中）

| 原始值 | 標準化後 |
|--------|----------|
| 列冊追踪 | 列冊追蹤 |
| 公司級 | Company.公司級文物 |
| 單位級 | Unit.單位級文物 |

## 側欄篩選結構

17 個 multiselect 分成 4 個 `st.sidebar.expander` 群組：

| 群組 | 欄位 |
|------|------|
| 📦 批次與單位 | 審查批次、典藏單位、系統別、來源對象名稱 |
| 🗂️ 典藏分類 | 典藏類型、典藏次類型、原件與否、藏品層次 |
| 🔖 主題與屬性 | 電業主題、文物屬性、主要材質、事件名稱 |
| ✅ 狀態與管理 | 保存狀況、建議分類、分級、使用限制、資料開放狀態 |

新增篩選欄位時需同步修改兩處：
1. 對應 expander 群組內新增 `ms_ex()`
2. 套用篩選的 `for col, chosen in [...]` 清單

`ms_ex()` 函式標籤顯示選項數（N 種），`help` tooltip 預覽前 20 個選項值。

## 已知資料陷阱

**分類表 merge 後建議分類全為空值**：原因是分類表的「文物登錄號」格式與主表不一致。主表格式為 `03015202302-00001`（無分隔符），若分類表寫成 `03-015-2023-02-00001`（有破折號），left join 找不到對應，欄位全為 NaN。解法是修正分類表的登錄號格式，去除多餘的 `-`。

## 注意事項

- 資料檔（xlsx）不進 repo，`.gitignore` 已排除。
- `load_all(base_dir)` 僅供本機開發，部署版只用 `load_from_uploads()`。
- 直接在 GitHub 上修改 `app.py` 後，需確認側欄篩選的兩處同步修改未被覆蓋。
