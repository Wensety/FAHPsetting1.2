# FAHP (模糊層次分析法) 分析程式

> Version 1.0
> Latest: Version 1.1

這是一個使用 Python 實作的模糊層次分析法（Fuzzy Analytic Hierarchy Process, FAHP）分析程式，支援從 Excel 檔案讀取資料進行分析。

## 功能特點

- ✅ 使用 pandas 進行資料處理
- ✅ 使用 NumPy 進行數值計算
- ✅ 使用 scikit-fuzzy 進行模糊邏輯運算
- ✅ 支援從 xlsx 檔案讀取分析資料
- ✅ 自動計算模糊權重和去模糊化
- ✅ 支援多準則多方案決策分析
- ✅ 自動匯出分析結果到 Excel

## 安裝需求

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 準備輸入資料

建立 Excel 檔案（例如 `input_data.xlsx`），包含以下工作表：

- **Criteria**: 準則列表
- **Alternatives**: 方案列表
- **Comparisons**: 準則比較矩陣

### 2. 執行分析

```python
from fahp_analysis import FAHPAnalyzer

# 建立分析器
analyzer = FAHPAnalyzer()

# 從Excel讀取資料
analyzer.read_excel('input_data.xlsx')

# 定義比較矩陣（1-9尺度）
criteria_comparison = [
    [1, 3, 5],
    [1/3, 1, 3],
    [1/5, 1/3, 1]
]

alternative_comparisons = {
    '價格': [[1, 2, 3], [1/2, 1, 2], [1/3, 1/2, 1]],
    '品質': [[1, 1/2, 1/3], [2, 1, 1/2], [3, 2, 1]],
    '服務': [[1, 3, 2], [1/3, 1, 1/2], [1/2, 2, 1]]
}

# 執行分析
results = analyzer.perform_fahp_analysis(
    criteria_comparison=criteria_comparison,
    alternative_comparisons=alternative_comparisons
)

# 匯出結果
analyzer.export_results(results, 'fahp_results.xlsx')
```

### 3. 直接執行主程式（會顯示版本）

```bash
python fahp_analysis.py
```

### 4. 產生範例輸入檔案

```bash
python example_input.py
```

## 比較矩陣尺度說明

使用標準的 1-9 尺度：

- 1: 同等重要
- 3: 稍微重要
- 5: 明顯重要
- 7: 強烈重要
- 9: 極端重要

倒數值（1/3, 1/5 等）表示相反的重要性關係。

### 5. 批次匯入多筆 xlsx 進行 FAHP 比較（Version 1.1）

1) 先準備多個符合問卷模板的 xlsx（建議使用 `generate_template.py` 生成 `questionnaire_template.xlsx` 後複製填寫）

2) 在資料夾中以檔名樣式篩選（例如 `*input*.xlsx`）

3) 執行批次：

```bash
py batch_fahp.py --input "D:\Path\To\Folder" --pattern "*.xlsx" --output batch_fahp_results.xlsx
```

輸出 `batch_fahp_results.xlsx` 內容：
- `Summary_Indicators`: 各檔案的指標全域權重對照與平均
- `Summary_Criteria`: 各檔案的構面權重對照與平均
- `File_{each}.xlsx`: 每個檔案的詳細結果

## 輸出結果

分析結果會匯出到 Excel 檔案，包含：

- **Criteria_Weights**: 準則權重（包含模糊數）
- **Final_Ranking**: 最終排名
- **Alt_Weights_***: 各準則下的方案權重

## 版本

- Version 1.1
  - 新增批次處理：可一次匯入多個問卷 xlsx，輸出指標/構面權重的跨檔案比較與平均。
  - 新增 `batch_fahp.py` 腳本（`--input`、`--pattern`、`--output`）。
  - 新增 `FAHPAnalyzer.analyze_from_questionnaire_file()` 與 `analyze_batch()` API。

- Version 1.0
  - 新增問卷結構（構面與指標）支援與自動解析。
  - 新增 `questionnaire_input.xlsx` 模板產生器。
  - 在執行檔與範例中輸出版本資訊。

## 類別說明

### FAHPAnalyzer

主要的分析器類別，提供以下方法：

- `read_excel()`: 從Excel讀取資料
- `create_fuzzy_comparison_matrix()`: 建立模糊比較矩陣
- `calculate_fuzzy_weights()`: 計算模糊權重
- `defuzzify()`: 去模糊化
- `perform_fahp_analysis()`: 執行完整分析
- `export_results()`: 匯出結果

## 注意事項

1. 確保比較矩陣符合一致性要求
2. 輸入的比較值應在 1-9 尺度範圍內
3. 對角線元素應為 1
4. 矩陣應為互反矩陣（a[i][j] = 1/a[j][i]）

## 授權

本程式碼供學習和研究使用。

