#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
FAHP 分析主程式
執行批次分析並產生統計分析報表

用法:
  python main.py [--input 資料夾路徑] [--pattern 檔案樣式] [--output 輸出檔名]

範例:
  python main.py --input runtest --pattern "*.xlsx" --output statistical_report.xlsx
  python main.py  # 使用預設值：處理 runtest 資料夾中的所有 xlsx 檔案
"""

import os
import sys
import glob
import argparse
import pandas as pd
import numpy as np

try:
    from fahp_analysis import FAHPAnalyzer, __version__
except ImportError as e:
    print(f"錯誤: 無法導入 fahp_analysis 模組: {e}")
    print(f"當前工作目錄: {os.getcwd()}")
    print(f"Python 路徑: {sys.path}")
    print("請確認 fahp_analysis.py 檔案存在於同一目錄")
    # 在調試模式下不調用 sys.exit()，讓錯誤自然顯示
    if 'pydevd' not in sys.modules and 'debugpy' not in sys.modules:
        sys.exit(1)
    else:
        raise  # 在調試模式下重新拋出異常以便調試


def generate_statistical_report(excel_files, output_file='statistical_report.xlsx'):
    """
    執行批次分析並產生統計分析報表
    
    Parameters:
    -----------
    excel_files : list
        要分析的 Excel 檔案路徑列表
    output_file : str
        輸出報表檔名
    
    Returns:
    --------
    tuple
        (輸出檔案路徑, 是否有優先度資料)
    """
    if not excel_files:
        print("錯誤: 沒有找到要分析的檔案")
        return None, False
    
    print(f"FAHP 統計分析報表生成器 Version {__version__}")
    print(f"找到 {len(excel_files)} 個文件:")
    for f in excel_files:
        print(f"  - {os.path.basename(f)}")
    print()
    
    # 執行批次分析
    analyzer = FAHPAnalyzer()
    batch_result = analyzer.analyze_batch(excel_files, output_file=output_file)
    
    # 收集所有數據用於統計分析
    all_indicator_data = []
    all_criteria_data = []
    
    for file_path, content in batch_result['per_file_results'].items():
        file_name = os.path.basename(file_path)
        results = content['results']
        
        # 收集指標全域權重
        if 'indicator_global_weights' in results and results['indicator_global_weights']:
            for item in results['indicator_global_weights']:
                try:
                    row = {
                        'File': file_name,
                        'Criterion': item.get('Criterion', ''),
                        'Indicator': item.get('Indicator', ''),
                        'Global_Weight': float(item.get('Global_Weight', 0.0)),
                        'Local_Weight': float(item.get('Local_Weight', 0.0)),
                        'Criterion_Weight': float(item.get('Criterion_Weight', 0.0))
                    }
                    if 'Priority' in item and item['Priority'] is not None:
                        row['Priority'] = float(item['Priority'])
                    all_indicator_data.append(row)
                except (KeyError, ValueError, TypeError) as e:
                    print(f"警告: 處理 {file_name} 的指標數據時發生錯誤: {e}")
                    continue
        
        # 收集構面權重
        if 'criteria_weights' in results and results['criteria_weights'] is not None:
            criteria_names = content.get('criteria_names', [])
            criteria_weights = results['criteria_weights']
            # 確保索引不會超出範圍
            min_len = min(len(criteria_names), len(criteria_weights))
            for i in range(min_len):
                try:
                    all_criteria_data.append({
                        'File': file_name,
                        'Criterion': criteria_names[i],
                        'Weight': float(criteria_weights[i])
                    })
                except (IndexError, ValueError, TypeError) as e:
                    print(f"警告: 處理 {file_name} 的構面數據時發生錯誤: {e}")
                    continue
    
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
            
            # 計算變異係數（避免除以零）
            stats_ind['變異係數'] = stats_ind.apply(
                lambda row: row['標準差'] / row['平均'] if row['平均'] != 0 else 0.0,
                axis=1
            )
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
            # 計算標準差時，需要確保有多列數據
            if len(pivot_ind.columns) > 1:
                # 排除最後一列（平均）來計算標準差
                pivot_ind['標準差'] = pivot_ind.iloc[:, :-1].std(axis=1, skipna=True)
            else:
                # 如果只有一列（平均），標準差為 0 或 NaN
                pivot_ind['標準差'] = 0.0
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
            
            # 計算變異係數（避免除以零）
            stats_crit['變異係數'] = stats_crit.apply(
                lambda row: row['標準差'] / row['平均'] if row['平均'] != 0 else 0.0,
                axis=1
            )
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
            # 計算標準差時，需要確保有多列數據
            if len(pivot_crit.columns) > 1:
                # 排除最後一列（平均）來計算標準差
                pivot_crit['標準差'] = pivot_crit.iloc[:, :-1].std(axis=1, skipna=True)
            else:
                # 如果只有一列（平均），標準差為 0 或 NaN
                pivot_crit['標準差'] = 0.0
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
    
    # 檢查是否有優先度資料
    has_priority = False
    if all_indicator_data:
        ind_df = pd.DataFrame(all_indicator_data)
        if 'Priority' in ind_df.columns and ind_df['Priority'].notna().any():
            has_priority = True
    
    return output_file, has_priority


def main():
    """主程式
    
    Returns:
    --------
    int
        退出碼：0 表示成功，非 0 表示失敗
    """
    parser = argparse.ArgumentParser(
        description='FAHP 批次分析與統計報表生成器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
  python main.py --input runtest --pattern "*.xlsx" --output report.xlsx
  python main.py --input runtest
  python main.py  # 使用預設值處理 runtest 資料夾
        """
    )
    parser.add_argument(
        '--input',
        default='runtest',
        help='輸入資料夾路徑（預設: runtest）'
    )
    parser.add_argument(
        '--pattern',
        default='*.xlsx',
        help='檔案比對樣式（預設: *.xlsx）'
    )
    parser.add_argument(
        '--output',
        default='statistical_report.xlsx',
        help='輸出報表檔名（預設: statistical_report.xlsx）'
    )
    
    args = parser.parse_args()
    
    # 檢查輸入資料夾
    input_dir = args.input
    if not os.path.exists(input_dir):
        print(f"錯誤: 資料夾 '{input_dir}' 不存在")
        return 1
    
    if not os.path.isdir(input_dir):
        print(f"錯誤: '{input_dir}' 不是一個資料夾")
        return 1
    
    # 搜尋檔案
    pattern = os.path.join(input_dir, args.pattern)
    files = glob.glob(pattern)
    files = [f for f in files if os.path.isfile(f)]
    files.sort()  # 排序以便一致處理
    
    if not files:
        print(f"錯誤: 在 '{input_dir}' 中找不到符合樣式 '{args.pattern}' 的檔案")
        return 1
    
    # 產生統計分析報表
    try:
        output_file, has_priority = generate_statistical_report(files, args.output)
        
        if output_file:
            print(f"\n{'='*60}")
            print(f"統計分析報表已生成: {output_file}")
            print(f"{'='*60}")
            print(f"包含以下工作表:")
            print(f"  ✓ 指標統計分析: 各指標的統計資訊（平均、標準差、變異係數等）")
            print(f"  ✓ 構面統計分析: 各構面的統計資訊")
            print(f"  ✓ 指標權重對照表: 各檔案指標權重對照")
            print(f"  ✓ 構面權重對照表: 各檔案構面權重對照")
            if has_priority:
                print(f"  ✓ 優先度統計: 優先度統計資訊")
            print(f"  ✓ 排名分析: 各指標的平均排名")
            print(f"  ✓ File_*: 各檔案的詳細結果")
            print(f"\n完成！報表已保存至: {output_file}")
            return 0
        else:
            print("錯誤: 無法產生報表")
            return 1
            
    except KeyboardInterrupt:
        print("\n\n程式被用戶中斷")
        return 130
    except Exception as e:
        print(f"錯誤: 執行過程中發生異常: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    try:
        exit_code = main()
        
        # 檢查是否在調試模式（VSCode 或 PyCharm）
        is_debugging = (
            'pydevd' in sys.modules or 
            'debugpy' in sys.modules or
            hasattr(sys, 'gettrace') and sys.gettrace() is not None
        )
        
        if not is_debugging:
            # 非調試模式下正常退出
            sys.exit(exit_code)
        else:
            # 調試模式下，只打印退出碼，不調用 sys.exit()
            # 這樣可以避免在調試器中顯示 SystemExit 異常
            if exit_code != 0:
                print(f"\n程式執行失敗，退出碼: {exit_code}")
            # 在調試模式下，讓程式自然結束，不調用 sys.exit()
            
    except KeyboardInterrupt:
        print("\n\n程式被用戶中斷")
        sys.exit(130)
    except Exception as e:
        print(f"\n發生未預期的錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

