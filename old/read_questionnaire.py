#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""讀取問卷PDF文件內容"""

try:
    import pdfplumber
    
    print("使用 pdfplumber 讀取PDF...")
    with pdfplumber.open('Questionnaire.pdf') as pdf:
        all_text = []
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                all_text.append(f"\n{'='*60}\n第 {i+1} 頁\n{'='*60}\n{text}")
        
        full_text = "\n".join(all_text)
        print(full_text)
        
        with open('questionnaire_content.txt', 'w', encoding='utf-8') as f:
            f.write(full_text)
        print(f"\n\n內容已保存到 questionnaire_content.txt")
        
except ImportError:
    print("pdfplumber 未安裝，嘗試使用 PyPDF2...")
    try:
        import PyPDF2
        
        with open('Questionnaire.pdf', 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            all_text = []
            for i, page in enumerate(pdf_reader.pages):
                text = page.extract_text()
                if text:
                    all_text.append(f"\n{'='*60}\n第 {i+1} 頁\n{'='*60}\n{text}")
            
            full_text = "\n".join(all_text)
            print(full_text)
            
            with open('questionnaire_content.txt', 'w', encoding='utf-8') as f:
                f.write(full_text)
            print(f"\n\n內容已保存到 questionnaire_content.txt")
            
    except Exception as e:
        print(f"讀取PDF時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
except Exception as e:
    print(f"讀取PDF時發生錯誤: {e}")
    import traceback
    traceback.print_exc()

