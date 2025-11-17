"""
讀取PDF文件內容
"""
import pdfplumber

def read_pdf(file_path):
    """讀取PDF文件內容"""
    text_content = []
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                text_content.append(f"=== 第 {i+1} 頁 ===\n{text}\n")
    return "\n".join(text_content)

if __name__ == '__main__':
    content = read_pdf('Questionnaire.pdf')
    print(content)
    # 同時保存到文件
    with open('questionnaire_content.txt', 'w', encoding='utf-8') as f:
        f.write(content)
    print("\n內容已保存到 questionnaire_content.txt")

