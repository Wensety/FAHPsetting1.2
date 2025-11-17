"""
FAHP分析程式使用範例
展示如何使用FAHPAnalyzer進行分析，包含問卷匯入情境
"""

from fahp_analysis import (
    FAHPAnalyzer,
    QUESTIONNAIRE_STRUCTURE,
    DEFAULT_CRITERIA_COMPARISON,
    DEFAULT_INDICATOR_COMPARISONS,
    __version__
)

def example_1_from_excel():
    """範例1: 從問卷Excel檔案讀取資料並進行分析"""
    print("=" * 60)
    print(f"範例1: 從問卷Excel檔案讀取資料 (Version {__version__})")
    print("=" * 60)
    
    analyzer = FAHPAnalyzer()
    
    try:
        analyzer.read_excel(
            'questionnaire_input.xlsx',
            criteria_sheet='Criteria',
            alternatives_sheet=None,
            indicators_sheet='Indicators'
        )
        criteria_comparison = analyzer.read_comparison_matrix_from_excel(
            'questionnaire_input.xlsx', 'Comparisons', header_row=0, index_col=0
        )

        indicator_comparisons = {}
        for criterion in analyzer.criteria_names:
            sheet_name = f'Pairwise_{criterion}'
            try:
                indicator_comparisons[criterion] = analyzer.read_comparison_matrix_from_excel(
                    'questionnaire_input.xlsx', sheet_name, header_row=0, index_col=0
                )
            except Exception:
                print(f"警告: 無法讀取 {criterion} 的指標比較矩陣，改用預設矩陣")
                indicator_comparisons[criterion] = DEFAULT_INDICATOR_COMPARISONS[criterion]

    except FileNotFoundError:
        print("找不到 questionnaire_input.xlsx，改用預設範例資料")
        analyzer.criteria_names = list(QUESTIONNAIRE_STRUCTURE.keys())
        analyzer.indicator_names = QUESTIONNAIRE_STRUCTURE
        criteria_comparison = DEFAULT_CRITERIA_COMPARISON
        indicator_comparisons = DEFAULT_INDICATOR_COMPARISONS

    results = analyzer.perform_fahp_analysis(
        criteria_comparison=criteria_comparison,
        alternative_comparisons=indicator_comparisons,
        defuzzify_method='centroid'
    )

    analyzer.export_results(results, 'example1_results.xlsx')

    if 'indicator_ranking' in results:
        print("\n問卷指標整體排名：")
        for rank, (indicator, criterion, weight) in enumerate(results['indicator_ranking'], 1):
            print(f"{rank}. {indicator} ({criterion}) -> {weight:.4f}")

    return analyzer, results


def example_2_direct_input():
    """範例2: 直接輸入資料進行分析"""
    print("\n" + "=" * 60)
    print("範例2: 直接輸入資料")
    print("=" * 60)
    
    analyzer = FAHPAnalyzer()
    
    # 設定準則和方案
    analyzer.criteria_names = ['成本', '效益', '風險']
    analyzer.alternatives_names = ['選項1', '選項2', '選項3']
    
    # 準則比較矩陣
    criteria_comparison = [
        [1, 2, 4],      # 成本 vs 成本、效益、風險
        [1/2, 1, 2],    # 效益 vs 成本、效益、風險
        [1/4, 1/2, 1]   # 風險 vs 成本、效益、風險
    ]
    
    # 各準則下的方案比較矩陣
    alternative_comparisons = {
        '成本': [
            [1, 1/2, 1/3],
            [2, 1, 1/2],
            [3, 2, 1]
        ],
        '效益': [
            [1, 3, 2],
            [1/3, 1, 1/2],
            [1/2, 2, 1]
        ],
        '風險': [
            [1, 2, 3],
            [1/2, 1, 2],
            [1/3, 1/2, 1]
        ]
    }
    
    # 執行分析
    results = analyzer.perform_fahp_analysis(
        criteria_comparison=criteria_comparison,
        alternative_comparisons=alternative_comparisons,
        defuzzify_method='centroid'
    )
    
    # 匯出結果
    analyzer.export_results(results, 'example2_results.xlsx')
    
    return analyzer, results


def example_3_custom_defuzzify():
    """範例3: 使用不同的去模糊化方法"""
    print("\n" + "=" * 60)
    print("範例3: 比較不同的去模糊化方法")
    print("=" * 60)
    
    analyzer = FAHPAnalyzer()
    analyzer.criteria_names = ['準則1', '準則2', '準則3']
    
    criteria_comparison = [
        [1, 3, 5],
        [1/3, 1, 3],
        [1/5, 1/3, 1]
    ]
    
    # 建立模糊比較矩陣
    analyzer.create_fuzzy_comparison_matrix('criteria', criteria_comparison)
    
    # 計算模糊權重
    fuzzy_weights = analyzer.calculate_fuzzy_weights('criteria')
    
    # 比較不同的去模糊化方法
    methods = ['centroid', 'mean', 'max']
    print("\n不同去模糊化方法的結果比較:")
    print("-" * 60)
    
    for method in methods:
        crisp_weights = analyzer.defuzzify(fuzzy_weights, method=method)
        print(f"\n{method} 方法:")
        for i, criterion in enumerate(analyzer.criteria_names):
            print(f"  {criterion}: {crisp_weights[i]:.4f}")
    
    return analyzer


if __name__ == '__main__':
    # 執行範例1（需要先執行 example_input.py 產生 questionnaire_input.xlsx）
    try:
        example_1_from_excel()
    except FileNotFoundError:
        print("請先執行 example_input.py 產生 questionnaire_input.xlsx 檔案")
    
    # 執行範例2
    example_2_direct_input()
    
    # 執行範例3
    example_3_custom_defuzzify()
    
    print("\n" + "=" * 60)
    print("所有範例執行完成！")
    print("=" * 60)

