"""
FAHP (Fuzzy Analytic Hierarchy Process) 分析程式
使用 pandas, NumPy 和 scikit-fuzzy 進行模糊層次分析法分析
支援從 xlsx 檔案讀取資料
"""

import pandas as pd
import numpy as np
import skfuzzy as fuzz
from skfuzzy import control as ctrl
import openpyxl
from typing import List, Tuple, Dict, Optional
import warnings
warnings.filterwarnings('ignore')

# 版本資訊
__version__ = "1.2"


# 問卷預設結構：主軸構面與其下的策略指標
QUESTIONNAIRE_STRUCTURE: Dict[str, List[str]] = {
    '環境品質與水資源': [
        '雨水滯留再生系統',
        '降低空污與改善',
        '地表水質維護',
        '園區碳排控管調節',
        '綠色基礎設施建置'
    ],
    '生態保育': [
        '提升物種多樣性',
        '植被覆蓋率',
        '生態廊道保護規劃',
        '保育稀有物種'
    ],
    '減災與氣候調適': [
        '極端降雨洪災緩解',
        '熱島效應調節改善',
        '土地河岸侵蝕防制',
        '地勢水文規劃'
    ],
    '社會共融與參與性': [
        '居民參與及治理',
        '社區凝聚力',
        '公共環境教育',
        '園區居民滿意度'
    ],
    '經濟可行性': [
        '創造綠色產業鏈',
        '綠色產業投資誘因',
        '生態環境經濟效益',
        '綠色技術採用度'
    ]
}


def _generate_sample_matrix(size: int, step: float = 1.5) -> List[List[float]]:
    """建立對稱的範例成對比較矩陣"""
    matrix: List[List[float]] = []
    for i in range(size):
        row: List[float] = []
        for j in range(size):
            if i == j:
                value = 1.0
            elif i < j:
                value = min(9.0, 1.0 + step * (j - i))
            else:
                # 取先前已建立之元素倒數，確保持成對倒數關係
                value = 1.0 / matrix[j][i]
            row.append(round(value, 4))
        matrix.append(row)
    return matrix


DEFAULT_CRITERIA_COMPARISON: List[List[float]] = _generate_sample_matrix(
    len(QUESTIONNAIRE_STRUCTURE), step=1.8
)


DEFAULT_INDICATOR_COMPARISONS: Dict[str, List[List[float]]] = {}
for idx, (criterion, indicators) in enumerate(QUESTIONNAIRE_STRUCTURE.items()):
    step_value = 1.4 + 0.2 * idx
    DEFAULT_INDICATOR_COMPARISONS[criterion] = _generate_sample_matrix(
        len(indicators), step=step_value
    )


class FAHPAnalyzer:
    """模糊層次分析法分析器"""
    
    def __init__(self):
        """初始化FAHP分析器"""
        self.criteria_names = []
        self.alternatives_names = []
        self.indicator_names: Dict[str, List[str]] = {}
        self.priority: Dict[str, Dict[str, float]] = {}  # 優先度：{構面: {指標: 優先度值}}
        self.fuzzy_matrices = {}
        self.weights = {}
        
    def read_excel(self, file_path: str, criteria_sheet: str = 'Criteria',
                   alternatives_sheet: Optional[str] = 'Alternatives',
                   comparison_sheet: str = 'Comparisons',
                   indicators_sheet: Optional[str] = None):
        """
        從 xlsx 檔案讀取資料
        
        Parameters:
        -----------
        file_path : str
            Excel檔案路徑
        criteria_sheet : str
            準則工作表名稱
        alternatives_sheet : str
            方案工作表名稱
        comparison_sheet : str
            比較矩陣工作表名稱（僅保留參考，實際讀取於 read_comparison_matrix_from_excel 中進行）
        indicators_sheet : str, optional
            策略指標工作表名稱，需包含構面與指標對應欄位
        """
        try:
            # 讀取準則
            try:
                criteria_df = pd.read_excel(file_path, sheet_name=criteria_sheet)
                if 'Criteria' in criteria_df.columns:
                    self.criteria_names = criteria_df['Criteria'].dropna().tolist()
                else:
                    self.criteria_names = criteria_df.iloc[:, 0].dropna().tolist()
                print(f"讀取到 {len(self.criteria_names)} 個準則: {self.criteria_names}")
            except Exception as e:
                print(f"警告: 無法讀取 {criteria_sheet} 工作表: {e}")
                self.criteria_names = []
            
            # 讀取方案（若有提供）
            self.alternatives_names = []
            if alternatives_sheet:
                try:
                    alternatives_df = pd.read_excel(file_path, sheet_name=alternatives_sheet)
                    if 'Alternatives' in alternatives_df.columns:
                        self.alternatives_names = alternatives_df['Alternatives'].dropna().astype(str).tolist()
                    else:
                        self.alternatives_names = alternatives_df.iloc[:, 0].dropna().astype(str).tolist()
                    print(f"讀取到 {len(self.alternatives_names)} 個方案: {self.alternatives_names}")
                except Exception as e:
                    print(f"警告: 無法讀取 {alternatives_sheet} 工作表: {e}")
                    self.alternatives_names = []

            # 讀取策略指標（若有提供）
            self.indicator_names = {}
            if indicators_sheet:
                try:
                    indicators_df = pd.read_excel(file_path, sheet_name=indicators_sheet)
                    if indicators_df.empty:
                        print(f"警告: {indicators_sheet} 工作表沒有資料")
                    else:
                        column_map = {str(col).lower(): col for col in indicators_df.columns}
                        candidate_criterion_cols = ['criterion', 'criteria', '構面', '主軸構面', '主構面']
                        candidate_indicator_cols = ['indicator', 'indicators', '指標', '策略項目', '項目']

                        criterion_col = next((column_map[name] for name in candidate_criterion_cols if name in column_map), None)
                        indicator_col = next((column_map[name] for name in candidate_indicator_cols if name in column_map), None)

                        if criterion_col is None and len(indicators_df.columns) >= 1:
                            criterion_col = indicators_df.columns[0]
                        if indicator_col is None:
                            if len(indicators_df.columns) >= 2:
                                indicator_col = indicators_df.columns[1]
                            else:
                                indicator_col = indicators_df.columns[0]

                        # 檢查是否有優先度欄位
                        candidate_priority_cols = ['priority', '優先度', '優先級', '重要性']
                        priority_col = next((column_map[name] for name in candidate_priority_cols if name in column_map), None)
                        
                        grouped = indicators_df[[criterion_col, indicator_col]].dropna(how='all')
                        for criterion, subset in grouped.groupby(criterion_col):
                            indicators = subset[indicator_col].dropna().astype(str).tolist()
                            if indicators:
                                self.indicator_names[str(criterion)] = indicators
                                
                                # 讀取優先度（如果有）
                                if priority_col and priority_col in subset.columns:
                                    priority_dict = {}
                                    for idx, row in subset.iterrows():
                                        indicator = str(row[indicator_col])
                                        priority_val = row[priority_col]
                                        if pd.notna(priority_val):
                                            try:
                                                priority_dict[indicator] = float(priority_val)
                                            except (ValueError, TypeError):
                                                priority_dict[indicator] = None
                                    if priority_dict:
                                        if str(criterion) not in self.priority:
                                            self.priority[str(criterion)] = {}
                                        self.priority[str(criterion)].update(priority_dict)

                        # 若尚未讀取到構面名稱，使用指標表的唯一值
                        if not self.criteria_names and criterion_col:
                            ordered_criteria = indicators_df[criterion_col].dropna().astype(str).unique().tolist()
                            self.criteria_names = ordered_criteria

                        if self.indicator_names:
                            summary = {k: len(v) for k, v in self.indicator_names.items()}
                            print(f"讀取到 {len(self.indicator_names)} 個構面的指標: {summary}")
                            if self.priority:
                                priority_summary = {k: len(v) for k, v in self.priority.items()}
                                print(f"讀取到優先度資料: {priority_summary}")
                        else:
                            print(f"警告: 無法自 {indicators_sheet} 解析指標資料")
                except Exception as e:
                    print(f"警告: 無法讀取 {indicators_sheet} 工作表: {e}")
                    self.indicator_names = {}
                
        except Exception as e:
            print(f"讀取Excel檔案時發生錯誤: {e}")
            raise
    
    def read_comparison_matrix_from_excel(self, file_path: str, sheet_name: str, 
                                         header_row: int = 0, index_col: int = 0) -> List[List[float]]:
        """
        從Excel檔案讀取比較矩陣
        
        Parameters:
        -----------
        file_path : str
            Excel檔案路徑
        sheet_name : str
            工作表名稱
        header_row : int
            標題行索引（預設0）
        index_col : int
            索引列索引（預設0）
        
        Returns:
        --------
        List[List[float]]
            比較矩陣（二維列表）
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, 
                             header=header_row, index_col=index_col)
            
            # 轉換為數值矩陣
            matrix = []
            for i in range(len(df)):
                row = []
                for j in range(len(df.columns)):
                    value = df.iloc[i, j]
                    # 處理字串形式的分數（如 "1/3"）
                    if isinstance(value, str) and '/' in value:
                        parts = value.split('/')
                        if len(parts) == 2:
                            value = float(parts[0]) / float(parts[1])
                        else:
                            value = float(value)
                    else:
                        value = float(value)
                    row.append(value)
                matrix.append(row)
            
            print(f"成功從 {sheet_name} 讀取 {len(matrix)}x{len(matrix[0])} 比較矩陣")
            return matrix
            
        except Exception as e:
            print(f"讀取比較矩陣時發生錯誤: {e}")
            raise
    
    
    def create_fuzzy_comparison_matrix(self, matrix_name: str, 
                                      comparison_values: List[List[float]]):
        """
        建立模糊比較矩陣
        
        Parameters:
        -----------
        matrix_name : str
            矩陣名稱（例如：'criteria' 或 'alternative_criterion_1'）
        comparison_values : List[List[float]]
            比較值矩陣，每個值代表重要性比例（1-9尺度）
        """
        n = len(comparison_values)
        fuzzy_matrix = np.zeros((n, n, 3))  # 每個元素是三角模糊數 (l, m, u)
        
        # 模糊化函數：將清晰值轉換為三角模糊數
        for i in range(n):
            for j in range(n):
                if i == j:
                    # 對角線元素為 (1, 1, 1)
                    fuzzy_matrix[i, j] = [1.0, 1.0, 1.0]
                else:
                    crisp_value = comparison_values[i][j]
                    # 將清晰值轉換為三角模糊數
                    # 使用標準的模糊化方法
                    l, m, u = self._fuzzify_value(crisp_value)
                    fuzzy_matrix[i, j] = [l, m, u]
                    # 對稱元素為倒數
                    fuzzy_matrix[j, i] = [1/u, 1/m, 1/l]
        
        self.fuzzy_matrices[matrix_name] = fuzzy_matrix
        return fuzzy_matrix
    
    def _fuzzify_value(self, value: float) -> Tuple[float, float, float]:
        """
        將清晰值模糊化為三角模糊數
        
        Parameters:
        -----------
        value : float
            清晰值（1-9尺度）
        
        Returns:
        --------
        Tuple[float, float, float]
            三角模糊數 (l, m, u)
        """
        # 標準的模糊化方法
        # 對於1-9尺度，使用以下轉換
        if value == 1:
            return (1.0, 1.0, 1.0)
        elif value == 2:
            return (1.0, 2.0, 3.0)
        elif value == 3:
            return (2.0, 3.0, 4.0)
        elif value == 4:
            return (3.0, 4.0, 5.0)
        elif value == 5:
            return (4.0, 5.0, 6.0)
        elif value == 6:
            return (5.0, 6.0, 7.0)
        elif value == 7:
            return (6.0, 7.0, 8.0)
        elif value == 8:
            return (7.0, 8.0, 9.0)
        elif value == 9:
            return (8.0, 9.0, 9.0)
        else:
            # 對於其他值，使用線性插值
            if value < 1:
                return (max(0.1, value - 0.5), value, min(1.0, value + 0.5))
            elif value > 9:
                return (value - 0.5, value, min(9.0, value + 0.5))
            else:
                # 線性插值
                lower = max(1, int(value) - 1)
                upper = min(9, int(value) + 1)
                return (lower, value, upper)
    
    def calculate_fuzzy_weights(self, matrix_name: str) -> np.ndarray:
        """
        計算模糊權重
        
        Parameters:
        -----------
        matrix_name : str
            矩陣名稱
        
        Returns:
        --------
        np.ndarray
            模糊權重矩陣 (n, 3)，每行是一個三角模糊數
        """
        if matrix_name not in self.fuzzy_matrices:
            raise ValueError(f"矩陣 {matrix_name} 不存在")
        
        fuzzy_matrix = self.fuzzy_matrices[matrix_name]
        n = fuzzy_matrix.shape[0]
        
        # 計算每行的幾何平均數（模糊數運算）
        fuzzy_weights = np.zeros((n, 3))
        
        for i in range(n):
            # 計算第i行的幾何平均數
            row_product = np.ones(3)
            for j in range(n):
                fuzzy_value = fuzzy_matrix[i, j]
                # 模糊數乘法：(l1, m1, u1) * (l2, m2, u2) = (l1*l2, m1*m2, u1*u2)
                row_product = row_product * fuzzy_value
            
            # 開n次方根
            fuzzy_weights[i] = np.power(row_product, 1.0/n)
        
        # 計算總和
        total_sum = np.sum(fuzzy_weights, axis=0)
        
        # 歸一化（模糊數除法）
        normalized_weights = np.zeros((n, 3))
        for i in range(n):
            # (l, m, u) / (L, M, U) = (l/U, m/M, u/L)
            normalized_weights[i] = [
                fuzzy_weights[i, 0] / total_sum[2],  # l / U
                fuzzy_weights[i, 1] / total_sum[1],  # m / M
                fuzzy_weights[i, 2] / total_sum[0]   # u / L
            ]
        
        self.weights[matrix_name] = normalized_weights
        return normalized_weights
    
    def defuzzify(self, fuzzy_weights: np.ndarray, method: str = 'centroid') -> np.ndarray:
        """
        去模糊化：將模糊權重轉換為清晰值
        
        Parameters:
        -----------
        fuzzy_weights : np.ndarray
            模糊權重矩陣 (n, 3)
        method : str
            去模糊化方法 ('centroid', 'mean', 'max')
        
        Returns:
        --------
        np.ndarray
            清晰權重向量
        """
        n = fuzzy_weights.shape[0]
        crisp_weights = np.zeros(n)
        
        if method == 'centroid':
            # 重心法：(l + m + u) / 3
            for i in range(n):
                crisp_weights[i] = np.mean(fuzzy_weights[i])
        elif method == 'mean':
            # 平均值法：(l + 2*m + u) / 4
            for i in range(n):
                crisp_weights[i] = (fuzzy_weights[i, 0] + 2*fuzzy_weights[i, 1] + 
                                   fuzzy_weights[i, 2]) / 4.0
        elif method == 'max':
            # 最大值法：取中值
            for i in range(n):
                crisp_weights[i] = fuzzy_weights[i, 1]
        else:
            raise ValueError(f"未知的去模糊化方法: {method}")
        
        # 歸一化
        total = np.sum(crisp_weights)
        if total > 0:
            crisp_weights = crisp_weights / total
        
        return crisp_weights
    
    def perform_fahp_analysis(self, criteria_comparison: List[List[float]],
                            alternative_comparisons: Dict[str, List[List[float]]] = None,
                            defuzzify_method: str = 'centroid') -> Dict:
        """
        執行完整的FAHP分析
        
        Parameters:
        -----------
        criteria_comparison : List[List[float]]
            準則比較矩陣
        alternative_comparisons : Dict[str, List[List[float]]]
            各準則下的方案比較矩陣，鍵為準則名稱
        defuzzify_method : str
            去模糊化方法
        
        Returns:
        --------
        Dict
            分析結果，包含權重和排名
        """
        results = {}
        
        # 1. 計算準則權重
        print("步驟 1: 計算準則模糊權重...")
        self.create_fuzzy_comparison_matrix('criteria', criteria_comparison)
        criteria_fuzzy_weights = self.calculate_fuzzy_weights('criteria')
        criteria_crisp_weights = self.defuzzify(criteria_fuzzy_weights, defuzzify_method)
        results['criteria_weights'] = criteria_crisp_weights
        results['criteria_fuzzy_weights'] = criteria_fuzzy_weights
        
        print(f"準則權重: {dict(zip(self.criteria_names, criteria_crisp_weights))}")
        
        # 2. 計算各準則下的方案權重
        if alternative_comparisons:
            if self.alternatives_names:
                print("\n步驟 2: 計算方案模糊權重...")
                alternative_weights = {}
                alternative_fuzzy_weights = {}

                alt_names = self.alternatives_names or []
                for criterion, comparison_matrix in alternative_comparisons.items():
                    matrix_name = f'alternative_{criterion}'
                    self.create_fuzzy_comparison_matrix(matrix_name, comparison_matrix)
                    fuzzy_weights = self.calculate_fuzzy_weights(matrix_name)
                    crisp_weights = self.defuzzify(fuzzy_weights, defuzzify_method)
                    alternative_weights[criterion] = crisp_weights
                    alternative_fuzzy_weights[criterion] = fuzzy_weights
                    if alt_names:
                        print(f"{criterion} 下的方案權重: {dict(zip(alt_names, np.round(crisp_weights, 4)))}")
                    else:
                        print(f"{criterion} 下的方案權重: {np.round(crisp_weights, 4).tolist()}")

                results['alternative_weights'] = alternative_weights
                results['alternative_fuzzy_weights'] = alternative_fuzzy_weights

                # 3. 計算最終綜合權重
                print("\n步驟 3: 計算最終綜合權重...")
                n_alternatives = len(alt_names)
                final_weights = np.zeros(n_alternatives)

                for i, alt in enumerate(alt_names):
                    for j, criterion in enumerate(self.criteria_names):
                        if criterion in alternative_weights:
                            final_weights[i] += criteria_crisp_weights[j] * alternative_weights[criterion][i]

                results['final_weights'] = final_weights
                results['ranking'] = sorted(zip(alt_names, final_weights),
                                           key=lambda x: x[1], reverse=True)

                print("\n最終綜合權重與排名:")
                for rank, (alt, weight) in enumerate(results['ranking'], 1):
                    print(f"{rank}. {alt}: {weight:.4f}")

            elif self.indicator_names:
                print("\n步驟 2: 計算構面下指標的模糊權重...")
                indicator_weights: Dict[str, np.ndarray] = {}
                indicator_fuzzy_weights: Dict[str, np.ndarray] = {}

                for criterion, comparison_matrix in alternative_comparisons.items():
                    matrix_name = f'indicator_{criterion}'
                    self.create_fuzzy_comparison_matrix(matrix_name, comparison_matrix)
                    fuzzy_weights = self.calculate_fuzzy_weights(matrix_name)
                    crisp_weights = self.defuzzify(fuzzy_weights, defuzzify_method)
                    indicator_weights[criterion] = crisp_weights
                    indicator_fuzzy_weights[criterion] = fuzzy_weights

                    indicator_labels = self.indicator_names.get(
                        criterion,
                        [f'指標{i+1}' for i in range(len(crisp_weights))]
                    )
                    print(f"{criterion} 下的指標權重: {dict(zip(indicator_labels, np.round(crisp_weights, 4)))}")

                results['indicator_weights'] = indicator_weights
                results['indicator_fuzzy_weights'] = indicator_fuzzy_weights

                # 3. 計算指標全域權重
                print("\n步驟 3: 計算指標全域權重...")
                indicator_global_weights = []

                for criterion_index, criterion in enumerate(self.criteria_names):
                    if criterion not in indicator_weights:
                        continue

                    indicator_labels = self.indicator_names.get(
                        criterion,
                        [f'指標{i+1}' for i in range(len(indicator_weights[criterion]))]
                    )
                    fuzzy_matrix = indicator_fuzzy_weights[criterion]

                    for label_index, indicator_label in enumerate(indicator_labels):
                        local_weight = float(indicator_weights[criterion][label_index])
                        global_weight = float(criteria_crisp_weights[criterion_index] * local_weight)
                        
                        # 讀取優先度（如果有）
                        priority_value = None
                        if criterion in self.priority and indicator_label in self.priority[criterion]:
                            priority_value = self.priority[criterion][indicator_label]
                        
                        item = {
                            'Criterion': criterion,
                            'Indicator': indicator_label,
                            'Local_Weight': local_weight,
                            'Criterion_Weight': float(criteria_crisp_weights[criterion_index]),
                            'Global_Weight': global_weight,
                            'Fuzzy_Lower': float(fuzzy_matrix[label_index, 0]),
                            'Fuzzy_Middle': float(fuzzy_matrix[label_index, 1]),
                            'Fuzzy_Upper': float(fuzzy_matrix[label_index, 2])
                        }
                        
                        # 如果有優先度，添加到結果中
                        if priority_value is not None:
                            item['Priority'] = priority_value
                        
                        indicator_global_weights.append(item)

                indicator_ranking = sorted(
                    [(item['Indicator'], item['Criterion'], item['Global_Weight']) for item in indicator_global_weights],
                    key=lambda x: x[2], reverse=True
                )

                results['indicator_global_weights'] = indicator_global_weights
                results['indicator_ranking'] = indicator_ranking

                print("\n指標整體權重排名:")
                for rank, (indicator_label, criterion, weight) in enumerate(indicator_ranking, 1):
                    print(f"{rank}. {indicator_label} ({criterion}): {weight:.4f}")

            else:
                print("警告: 未設定方案或指標名稱，無法計算後續權重。")
        
        return results

    # --------------- Version 1.1: 單檔與批次處理 API -----------------
    def analyze_from_questionnaire_file(self, excel_file: str) -> Dict:
        """依問卷格式讀取單一 xlsx 並完成 FAHP 分析，回傳結果。

        需求：包含 'Criteria'、'Comparisons'、'Indicators' 與各 'Pairwise_{構面}' 工作表。
        若部分欠缺，將落回 QUESTIONNAIRE_STRUCTURE 的預設並給提示。
        """
        # 清理名稱以避免殘留
        self.criteria_names = []
        self.alternatives_names = []
        self.indicator_names = {}
        self.priority = {}

        # 讀基本表
        self.read_excel(
            excel_file,
            criteria_sheet='Criteria',
            alternatives_sheet=None,
            indicators_sheet='Indicators'
        )

        # 構面比較矩陣
        try:
            criteria_comparison = self.read_comparison_matrix_from_excel(
                excel_file, 'Comparisons', header_row=0, index_col=0
            )
        except Exception as exc:
            print(f"警告: 無法讀取 Comparisons，改用預設；原因: {exc}")
            if not self.criteria_names:
                self.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
            criteria_comparison = DEFAULT_CRITERIA_COMPARISON

        # 指標映射與比較矩陣
        if not self.criteria_names:
            self.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
        if not self.indicator_names:
            self.indicator_names = QUESTIONNAIRE_STRUCTURE

        indicator_comparisons: Dict[str, List[List[float]]] = {}
        for criterion in self.criteria_names:
            sheet_name = f'Pairwise_{criterion}'
            try:
                indicator_comparisons[criterion] = self.read_comparison_matrix_from_excel(
                    excel_file, sheet_name, header_row=0, index_col=0
                )
            except Exception as exc:
                print(f"警告: 無法讀取 {sheet_name}，改用預設；原因: {exc}")
                indicator_comparisons[criterion] = DEFAULT_INDICATOR_COMPARISONS.get(
                    criterion,
                    _generate_sample_matrix(len(self.indicator_names.get(criterion, [])) or 3)
                )

        # 執行分析
        return self.perform_fahp_analysis(
            criteria_comparison=criteria_comparison,
            alternative_comparisons=indicator_comparisons,
            defuzzify_method='centroid'
        )

    def analyze_batch(self, excel_files: List[str], output_file: str = 'batch_fahp_results.xlsx') -> Dict:
        """批次分析多個問卷 xlsx，輸出彙總比較結果。

        產出：
        - Summary_Indicators：各指標於不同檔案的 Global_Weight 對照與平均
        - Summary_Criteria：各構面於不同檔案的權重對照與平均
        - 每個檔案一個 Sheet：紀錄其指標全域權重與排名
        """
        import pandas as pd
        import numpy as np
        summary_indicator_rows = []
        summary_criteria_rows = []
        per_file_results: Dict[str, Dict] = {}

        for file_path in excel_files:
            print(f"批次處理: {file_path}")
            # 每次重新建立一個分析器，避免狀態殘留
            analyzer = FAHPAnalyzer()
            result = analyzer.analyze_from_questionnaire_file(file_path)
            per_file_results[file_path] = {
                'criteria_names': analyzer.criteria_names,
                'indicator_names': analyzer.indicator_names,
                'results': result
            }

            # 收集構面（準則）權重
            for i, crit in enumerate(analyzer.criteria_names):
                summary_criteria_rows.append({
                    'File': file_path,
                    'Criterion': crit,
                    'Weight': float(result['criteria_weights'][i])
                })

            # 收集指標全域權重
            for item in result.get('indicator_global_weights', []):
                summary_indicator_rows.append({
                    'File': file_path,
                    'Criterion': item['Criterion'],
                    'Indicator': item['Indicator'],
                    'Global_Weight': float(item['Global_Weight'])
                })

        # 彙整輸出
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 指標彙整：每列一個 (Criterion, Indicator)，欄為各檔案 + Average
            if summary_indicator_rows:
                ind_df = pd.DataFrame(summary_indicator_rows)
                pivot = ind_df.pivot_table(
                    index=['Criterion', 'Indicator'],
                    columns='File',
                    values='Global_Weight',
                    aggfunc='mean'
                )
                pivot['Average'] = pivot.mean(axis=1, skipna=True)
                pivot = pivot.sort_values(by='Average', ascending=False)
                pivot.to_excel(writer, sheet_name='Summary_Indicators')

            # 構面彙整：每列一個 Criterion，欄為各檔案 + Average
            if summary_criteria_rows:
                crit_df = pd.DataFrame(summary_criteria_rows)
                pivot_c = crit_df.pivot_table(
                    index=['Criterion'],
                    columns='File',
                    values='Weight',
                    aggfunc='mean'
                )
                pivot_c['Average'] = pivot_c.mean(axis=1, skipna=True)
                pivot_c = pivot_c.sort_values(by='Average', ascending=False)
                pivot_c.to_excel(writer, sheet_name='Summary_Criteria')

            # 個別檔案結果
            for file_path, content in per_file_results.items():
                res = content['results']
                if 'indicator_global_weights' in res:
                    df = pd.DataFrame(res['indicator_global_weights'])
                    safe_name = file_path.split('/')[-1].split('\\')[-1]
                    sheet = f"File_{safe_name}"
                    if len(sheet) > 31:
                        sheet = sheet[:31]
                    df.to_excel(writer, sheet_name=sheet, index=False)

        return {
            'per_file_results': per_file_results,
            'output_file': output_file
        }
    
    def export_results(self, results: Dict, output_file: str = 'fahp_results.xlsx'):
        """
        匯出分析結果到Excel檔案
        
        Parameters:
        -----------
        results : Dict
            分析結果
        output_file : str
            輸出檔案名稱
        """
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 準則權重
            criteria_df = pd.DataFrame({
                'Criteria': self.criteria_names,
                'Weight': results['criteria_weights'],
                'Fuzzy_Lower': results['criteria_fuzzy_weights'][:, 0],
                'Fuzzy_Middle': results['criteria_fuzzy_weights'][:, 1],
                'Fuzzy_Upper': results['criteria_fuzzy_weights'][:, 2]
            })
            criteria_df.to_excel(writer, sheet_name='Criteria_Weights', index=False)
            
            # 指標權重結果（問卷結構）
            if 'indicator_global_weights' in results:
                indicator_global_df = pd.DataFrame(results['indicator_global_weights'])
                indicator_global_df.to_excel(writer, sheet_name='Indicator_Global_Weights', index=False)

            if 'indicator_weights' in results:
                indicator_fuzzy = results.get('indicator_fuzzy_weights', {})
                for criterion, weights in results['indicator_weights'].items():
                    indicator_labels = self.indicator_names.get(
                        criterion,
                        [f'指標{i+1}' for i in range(len(weights))]
                    )
                    fuzzy_matrix = indicator_fuzzy.get(criterion)
                    data = {
                        'Indicator': indicator_labels,
                        'Local_Weight': weights
                    }
                    if isinstance(fuzzy_matrix, np.ndarray):
                        data.update({
                            'Fuzzy_Lower': fuzzy_matrix[:, 0],
                            'Fuzzy_Middle': fuzzy_matrix[:, 1],
                            'Fuzzy_Upper': fuzzy_matrix[:, 2]
                        })
                    
                    # 添加優先度（如果有）
                    if criterion in self.priority:
                        priority_values = [
                            self.priority[criterion].get(ind, None) 
                            for ind in indicator_labels
                        ]
                        if any(p is not None for p in priority_values):
                            data['Priority'] = priority_values
                    
                    indicator_df = pd.DataFrame(data)
                    sheet_name = f'Indicators_{criterion}'
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    indicator_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 方案權重和排名
            if 'final_weights' in results:
                ranking_df = pd.DataFrame({
                    'Alternative': [alt for alt, _ in results['ranking']],
                    'Final_Weight': [weight for _, weight in results['ranking']],
                    'Rank': range(1, len(results['ranking']) + 1)
                })
                ranking_df.to_excel(writer, sheet_name='Final_Ranking', index=False)
                
                # 各準則下的方案權重
                for criterion, weights in results.get('alternative_weights', {}).items():
                    alt_labels = self.alternatives_names or [f'方案{i+1}' for i in range(len(weights))]
                    fuzzy_matrix = results.get('alternative_fuzzy_weights', {}).get(criterion)
                    data = {
                        'Alternative': alt_labels,
                        'Weight': weights
                    }
                    if isinstance(fuzzy_matrix, np.ndarray):
                        data.update({
                            'Fuzzy_Lower': fuzzy_matrix[:, 0],
                            'Fuzzy_Middle': fuzzy_matrix[:, 1],
                            'Fuzzy_Upper': fuzzy_matrix[:, 2]
                        })
                    alt_df = pd.DataFrame(data)
                    sheet_name = f'Alt_Weights_{criterion}'
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    alt_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n結果已匯出至: {output_file}")


def main():
    """主程式範例：配合問卷結構讀取並分析Excel資料"""

    analyzer = FAHPAnalyzer()
    print(f"FAHP Analyzer Version {__version__}")
    excel_file = 'questionnaire_input.xlsx'
    indicator_comparisons: Dict[str, List[List[float]]] = {}

    try:
        analyzer.read_excel(
            excel_file,
            criteria_sheet='Criteria',
            alternatives_sheet=None,
            indicators_sheet='Indicators'
        )
        print("\n成功讀取Excel檔案")

        try:
            criteria_comparison = analyzer.read_comparison_matrix_from_excel(
                excel_file, 'Comparisons', header_row=0, index_col=0
            )
            print("成功從Excel讀取準則比較矩陣")
        except Exception as exc:
            raise RuntimeError(f"無法從Excel讀取準則比較矩陣: {exc}") from exc

        if not analyzer.criteria_names:
            analyzer.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
        if not analyzer.indicator_names:
            analyzer.indicator_names = QUESTIONNAIRE_STRUCTURE

        for criterion in analyzer.criteria_names:
            sheet_name = f'Pairwise_{criterion}'
            try:
                matrix = analyzer.read_comparison_matrix_from_excel(
                    excel_file, sheet_name, header_row=0, index_col=0
                )
                indicator_comparisons[criterion] = matrix
                print(f"成功從Excel讀取 {criterion} 的指標比較矩陣")
            except Exception as exc:
                print(f"警告: 無法讀取 {criterion} 的指標比較矩陣 ({sheet_name}): {exc}")

        if not indicator_comparisons:
            print("警告: 沒有讀取到任何指標比較矩陣，使用預設範例矩陣")
            indicator_comparisons = {
                criterion: DEFAULT_INDICATOR_COMPARISONS.get(criterion)
                for criterion in analyzer.indicator_names
                if criterion in DEFAULT_INDICATOR_COMPARISONS
            }

    except FileNotFoundError:
        print("未找到 questionnaire_input.xlsx，使用問卷預設範例資料")
        analyzer.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
        analyzer.indicator_names = QUESTIONNAIRE_STRUCTURE
        criteria_comparison = DEFAULT_CRITERIA_COMPARISON
        indicator_comparisons = {
            criterion: DEFAULT_INDICATOR_COMPARISONS[criterion]
            for criterion in analyzer.criteria_names
        }
    except RuntimeError as exc:
        print(exc)
        print("改用問卷預設範例資料進行示範")
        analyzer.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
        analyzer.indicator_names = QUESTIONNAIRE_STRUCTURE
        criteria_comparison = DEFAULT_CRITERIA_COMPARISON
        indicator_comparisons = {
            criterion: DEFAULT_INDICATOR_COMPARISONS[criterion]
            for criterion in analyzer.criteria_names
        }

    # 執行FAHP分析
    print("\n開始執行FAHP分析...")
    results = analyzer.perform_fahp_analysis(
        criteria_comparison=criteria_comparison,
        alternative_comparisons=indicator_comparisons,
        defuzzify_method='centroid'
    )

    # 匯出結果
    analyzer.export_results(results, 'fahp_results.xlsx')

    return analyzer, results


if __name__ == '__main__':
    analyzer, results = main()

