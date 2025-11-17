#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""測試 runtest 資料夾中的問卷文件，驗證優先度讀取功能"""

import os
import glob
from fahp_analysis import FAHPAnalyzer

def test_runtest_files():
    """測試 runtest 資料夾中的文件"""
    runtest_dir = 'runtest'
    
    if not os.path.exists(runtest_dir):
        print(f"資料夾 {runtest_dir} 不存在")
        return
    
    # 取得所有 xlsx 文件
    pattern = os.path.join(runtest_dir, '*.xlsx')
    files = glob.glob(pattern)
    
    if not files:
        print(f"在 {runtest_dir} 中找不到 xlsx 文件")
        return
    
    print(f"找到 {len(files)} 個文件，開始測試...\n")
    
    for file_path in files:
        print(f"=" * 60)
        print(f"處理文件: {os.path.basename(file_path)}")
        print(f"=" * 60)
        
        analyzer = FAHPAnalyzer()
        try:
            result = analyzer.analyze_from_questionnaire_file(file_path)
            
            # 顯示優先度資訊
            if analyzer.priority:
                print("\n✓ 讀取到優先度資料:")
                for criterion, priorities in analyzer.priority.items():
                    print(f"  構面: {criterion}")
                    for indicator, priority_val in priorities.items():
                        print(f"    - {indicator}: {priority_val}")
            else:
                print("\n⚠ 未讀取到優先度資料（可能文件中沒有優先度欄位）")
            
            # 檢查結果中的優先度
            if 'indicator_global_weights' in result:
                items_with_priority = [item for item in result['indicator_global_weights'] if 'Priority' in item]
                if items_with_priority:
                    print(f"\n✓ 結果中包含優先度的指標: {len(items_with_priority)} 個")
                else:
                    print("\n⚠ 結果中未包含優先度欄位")
            
            print("\n")
            
        except Exception as e:
            print(f"✗ 處理文件時發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            print("\n")

if __name__ == '__main__':
    test_runtest_files()

