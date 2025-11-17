"""
建立問卷範例輸入Excel檔案的腳本
執行此腳本可產生 questionnaire_input.xlsx 檔案
"""

import pandas as pd

from fahp_analysis import (
    QUESTIONNAIRE_STRUCTURE,
    DEFAULT_CRITERIA_COMPARISON,
    DEFAULT_INDICATOR_COMPARISONS
)


def _build_pairwise_df(labels, matrix):
    df = pd.DataFrame(matrix, columns=labels, index=labels)
    df.index.name = 'Label'
    return df


criteria_list = list(QUESTIONNAIRE_STRUCTURE.keys())
criteria_df = pd.DataFrame({
    'Order': range(1, len(criteria_list) + 1),
    'Criteria': criteria_list
})

indicator_rows = []
for criterion, indicators in QUESTIONNAIRE_STRUCTURE.items():
    for order, indicator in enumerate(indicators, start=1):
        indicator_rows.append({
            'Criterion': criterion,
            'Indicator_Order': order,
            'Indicator': indicator
        })

indicators_df = pd.DataFrame(indicator_rows)


with pd.ExcelWriter('questionnaire_input.xlsx', engine='openpyxl') as writer:
    criteria_df.to_excel(writer, sheet_name='Criteria', index=False)
    indicators_df.to_excel(writer, sheet_name='Indicators', index=False)

    criteria_pairwise_df = _build_pairwise_df(criteria_list, DEFAULT_CRITERIA_COMPARISON)
    criteria_pairwise_df.to_excel(writer, sheet_name='Comparisons')

    for criterion, indicators in QUESTIONNAIRE_STRUCTURE.items():
        pairwise_matrix = DEFAULT_INDICATOR_COMPARISONS[criterion]
        pairwise_df = _build_pairwise_df(indicators, pairwise_matrix)
        sheet_name = f'Pairwise_{criterion}'
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        pairwise_df.to_excel(writer, sheet_name=sheet_name)

print("問卷範例輸入檔案 questionnaire_input.xlsx 已建立完成！")

