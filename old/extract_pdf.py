import pdfplumber
import sys

try:
    with pdfplumber.open('Questionnaire.pdf') as pdf:
        all_text = []
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                all_text.append(f"=== Page {i+1} ===\n{text}\n")
        
        full_text = "\n".join(all_text)
        print(full_text)
        
        with open('questionnaire_content.txt', 'w', encoding='utf-8') as f:
            f.write(full_text)
        print("\n\nContent saved to questionnaire_content.txt")
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)

