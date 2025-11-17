#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""批次執行 FAHP 分析

用法（PowerShell / CMD）:
  py batch_fahp.py --input "D:\\FAHPset" --pattern "*input*.xlsx" --output batch_fahp_results.xlsx

參數:
  --input   指定資料夾路徑
  --pattern 檔名比對樣式（預設 *.xlsx）
  --output  輸出彙總檔名（預設 batch_fahp_results.xlsx）
"""

import argparse
import glob
import os
from fahp_analysis import FAHPAnalyzer, __version__


def main():
    parser = argparse.ArgumentParser(description='批次執行 FAHP 分析')
    parser.add_argument('--input', required=True, help='輸入資料夾路徑')
    parser.add_argument('--pattern', default='*.xlsx', help='比對樣式，預設 *.xlsx')
    parser.add_argument('--output', default='batch_fahp_results.xlsx', help='輸出Excel檔名')
    args = parser.parse_args()

    folder = args.input
    pattern = args.pattern
    output = args.output

    search_path = os.path.join(folder, pattern)
    files = [f for f in glob.glob(search_path) if os.path.isfile(f)]
    if not files:
        print(f'找不到符合的檔案: {search_path}')
        return 1

    print(f'FAHP Batch Runner Version {__version__}')
    print('將處理以下檔案:')
    for f in files:
        print(' -', f)

    analyzer = FAHPAnalyzer()
    result = analyzer.analyze_batch(files, output_file=output)
    print(f"彙總輸出完成: {result['output_file']}")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())


