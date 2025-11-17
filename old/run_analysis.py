#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""執行 runtest 資料夾的批次分析並產生統計分析報表"""

import os
import sys
import glob
import pandas as pd
import numpy as np

try:
    from fahp_analysis import FAHPAnalyzer, __version__
except ImportError as e:
    print(f"導入錯誤: {e}")
    sys.exit(1)

# 取得 runtest 資料夾中的所有 xlsx 文件
runtest_dir = 'runtest'
if not os.path.exists(runtest_dir):
    print(f"錯誤: 資料夾 {runtest_dir} 不存在")
    sys.exit(1)

pattern = os.path.join(runtest_dir, '*.xlsx')
files = glob.glob(pattern)

if not files:
    print(f"錯誤: 在 {runtest_dir} 中找不到 xlsx 文件")
    sys.exit(1)

files.sort()  # 排序以便一致處理

print(f"FAHP 統計分析報表生成器 Version {__version__}")
print(f"找到 {len(files)} 個文件:")
for f in files:
    print(f"  - {os.path.basename(f)}")
print()

# 執行批次分析
analyzer = FAHPAnalyzer()
output_file = 'runtest_statistical_report.xlsx'
batch_result = analyzer.analyze_batch(files, output_file=output_file)

# 收集所有數據用於統計分析
all_indicator_data = []
all_criteria_data = []

for file_path, content in batch_result['per_file_results'].items():
    file_name = os.path.basename(file_path)
    results = content['results']
    
    # 收集指標全域權重
    if 'indicator_global_weights' in results:
        for item in results['indicator_global_weights']:
            row = {
                'File': file_name,
                'Criterion': item['Criterion'],
                'Indicator': item['Indicator'],
                'Global_Weight': item['Global_Weight'],
                'Local_Weight': item['Local_Weight'],
                'Criterion_Weight': item['Criterion_Weight']
            }
            if 'Priority' in item:
                row['Priority'] = item['Priority']
            all_indicator_data.append(row)
    
    # 收集構面權重
    if 'criteria_weights' in results:
        criteria_names = content['criteria_names']
        for i, criterion in enumerate(criteria_names):
            all_criteria_data.append({
                'File': file_name,
                'Criterion': criterion,
                'Weight': results['criteria_weights'][i]
            })

# 建立統計分析報表
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    
    # 1. 指標統計分析
    if all_indicator_data:
        ind_df = pd.DataFrame(all_indicator_data)
        
        # 基本統計
        stats_ind = ind_df.groupby(['Criterion', 'Indicator'])['Global_Weight'].agg([
            ('平均', 'mean'),
            ('標準差', 'std'),
            ('最小值', 'min'),
            ('最大值', 'max'),
            ('中位數', 'median'),
            ('樣本數', 'count')
        ]).reset_index()
        
        # 計算變異係數
        stats_ind['變異係數'] = stats_ind['標準差'] / stats_ind['平均']
        stats_ind = stats_ind.sort_values(by='平均', ascending=False)
        
        stats_ind.to_excel(writer, sheet_name='指標統計分析', index=False)
        
        # 詳細數據透視表
        pivot_ind = ind_df.pivot_table(
            index=['Criterion', 'Indicator'],
            columns='File',
            values='Global_Weight',
            aggfunc='mean'
        )
        pivot_ind['平均'] = pivot_ind.mean(axis=1, skipna=True)
        pivot_ind['標準差'] = pivot_ind.iloc[:, :-1].std(axis=1, skipna=True)
        pivot_ind = pivot_ind.sort_values(by='平均', ascending=False)
        pivot_ind.to_excel(writer, sheet_name='指標權重對照表')
    
    # 2. 構面統計分析
    if all_criteria_data:
        crit_df = pd.DataFrame(all_criteria_data)
        
        # 基本統計
        stats_crit = crit_df.groupby('Criterion')['Weight'].agg([
            ('平均', 'mean'),
            ('標準差', 'std'),
            ('最小值', 'min'),
            ('最大值', 'max'),
            ('中位數', 'median'),
            ('樣本數', 'count')
        ]).reset_index()
        
        # 計算變異係數
        stats_crit['變異係數'] = stats_crit['標準差'] / stats_crit['平均']
        stats_crit = stats_crit.sort_values(by='平均', ascending=False)
        
        stats_crit.to_excel(writer, sheet_name='構面統計分析', index=False)
        
        # 詳細數據透視表
        pivot_crit = crit_df.pivot_table(
            index='Criterion',
            columns='File',
            values='Weight',
            aggfunc='mean'
        )
        pivot_crit['平均'] = pivot_crit.mean(axis=1, skipna=True)
        pivot_crit['標準差'] = pivot_crit.iloc[:, :-1].std(axis=1, skipna=True)
        pivot_crit = pivot_crit.sort_values(by='平均', ascending=False)
        pivot_crit.to_excel(writer, sheet_name='構面權重對照表')
    
    # 3. 優先度統計（如果有）
    if all_indicator_data:
        ind_df = pd.DataFrame(all_indicator_data)
        if 'Priority' in ind_df.columns:
            priority_df = ind_df[ind_df['Priority'].notna()].copy()
            if not priority_df.empty:
                priority_stats = priority_df.groupby(['Criterion', 'Indicator'])['Priority'].agg([
                    ('平均', 'mean'),
                    ('標準差', 'std'),
                    ('最小值', 'min'),
                    ('最大值', 'max'),
                    ('樣本數', 'count')
                ]).reset_index()
                priority_stats = priority_stats.sort_values(by='平均', ascending=False)
                priority_stats.to_excel(writer, sheet_name='優先度統計', index=False)
    
    # 4. 排名分析
    if all_indicator_data:
        ranking_data = []
        ind_df = pd.DataFrame(all_indicator_data)
        for file_name in ind_df['File'].unique():
            file_data = ind_df[ind_df['File'] == file_name]
            sorted_data = file_data.sort_values('Global_Weight', ascending=False)
            for rank, (idx, row) in enumerate(sorted_data.iterrows(), 1):
                ranking_data.append({
                    'File': file_name,
                    'Rank': rank,
                    'Criterion': row['Criterion'],
                    'Indicator': row['Indicator'],
                    'Global_Weight': row['Global_Weight']
                })
        
        ranking_df = pd.DataFrame(ranking_data)
        
        # 計算每個指標的平均排名
        avg_ranking = ranking_df.groupby(['Criterion', 'Indicator']).agg({
            'Rank': 'mean',
            'Global_Weight': 'mean'
        }).reset_index()
        avg_ranking.columns = ['Criterion', 'Indicator', '平均排名', '平均權重']
        avg_ranking = avg_ranking.sort_values('平均排名')
        avg_ranking.to_excel(writer, sheet_name='排名分析', index=False)
    
    # 5. 個別檔案詳細結果
    for file_path, content in batch_result['per_file_results'].items():
        res = content['results']
        if 'indicator_global_weights' in res:
            df = pd.DataFrame(res['indicator_global_weights'])
            safe_name = os.path.basename(file_path)
            sheet = f"File_{safe_name}"
            if len(sheet) > 31:
                sheet = sheet[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)

print(f"\n統計分析報表已生成: {output_file}")
print(f"包含以下工作表:")
print(f"  - 指標統計分析: 各指標的統計資訊（平均、標準差、變異係數等）")
print(f"  - 構面統計分析: 各構面的統計資訊")
print(f"  - 指標權重對照表: 各檔案指標權重對照")
print(f"  - 構面權重對照表: 各檔案構面權重對照")
if all_indicator_data and 'Priority' in pd.DataFrame(all_indicator_data).columns:
    print(f"  - 優先度統計: 優先度統計資訊")
print(f"  - 排名分析: 各指標的平均排名")
print(f"  - File_*: 各檔案的詳細結果")
print(f"\n完成！報表已保存至: {output_file}")

