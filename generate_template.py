"""
產生問卷用的可填寫樣板：questionnaire_template.xlsx
包含：
- Criteria（構面清單，已填好）
- Indicators（各構面指標清單，已填好）
- Comparisons（構面成對比較矩陣，僅對角線為1，其餘留空待填）
- Pairwise_{構面}（該構面下指標成對比較矩陣，僅對角線為1，其餘留空待填）
- GUIDE（填寫說明與1-9尺度提示）
"""

import pandas as pd

from fahp_analysis import QUESTIONNAIRE_STRUCTURE


def _empty_pairwise_df(labels):
    df = pd.DataFrame(index=labels, columns=labels, dtype=object)
    for i, row_label in enumerate(labels):
        for j, col_label in enumerate(labels):
            df.loc[row_label, col_label] = 1 if i == j else None
    df.index.name = 'Label'
    return df


def main():
    criteria_list = list(QUESTIONNAIRE_STRUCTURE.keys())

    # Criteria sheet
    criteria_df = pd.DataFrame({
        'Order': range(1, len(criteria_list) + 1),
        'Criteria': criteria_list
    })

    # Indicators sheet
    indicator_rows = []
    for criterion, indicators in QUESTIONNAIRE_STRUCTURE.items():
        for order, indicator in enumerate(indicators, start=1):
            indicator_rows.append({
                'Criterion': criterion,
                'Indicator_Order': order,
                'Indicator': indicator
            })
    indicators_df = pd.DataFrame(indicator_rows)

    # Guide sheet
    guide_lines = [
        ["填寫說明"],
        ["1. Comparisons：為構面間的兩兩比較矩陣，請於非對角線的儲存格填入1~9或其倒數(如1/3)。"],
        ["2. Pairwise_{構面}：為該構面指標間的兩兩比較矩陣，請於非對角線儲存格填入1~9或其倒數。"],
        ["3. 對角線已填為1；其餘空白請依相對重要性填寫。矩陣需滿足互反關係 a[i][j] = 1 / a[j][i]。"],
        ["4. 標準尺度：1(同等) 3(稍重要) 5(明顯重要) 7(強烈重要) 9(極端重要)；2/4/6/8為中間值。"],
        ["5. 填寫完成後，可執行 fahp_analysis.py 進行分析並輸出權重與排名。"],
    ]
    guide_df = pd.DataFrame(guide_lines, columns=["Guide"])

    with pd.ExcelWriter('questionnaire_template.xlsx', engine='openpyxl') as writer:
        criteria_df.to_excel(writer, sheet_name='Criteria', index=False)
        indicators_df.to_excel(writer, sheet_name='Indicators', index=False)
        guide_df.to_excel(writer, sheet_name='GUIDE', index=False)

        # Empty pairwise for Criteria
        criteria_pairwise_df = _empty_pairwise_df(criteria_list)
        criteria_pairwise_df.to_excel(writer, sheet_name='Comparisons')

        # Empty pairwise per criterion for indicators
        for criterion, indicators in QUESTIONNAIRE_STRUCTURE.items():
            pairwise_df = _empty_pairwise_df(indicators)
            sheet_name = f'Pairwise_{criterion}'
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            pairwise_df.to_excel(writer, sheet_name=sheet_name)

    print('questionnaire_template.xlsx 已建立，請於空白儲存格填入1~9或分數值(如1/3)。')


if __name__ == '__main__':
    main()




