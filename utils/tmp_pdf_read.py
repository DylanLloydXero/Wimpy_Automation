import pdfplumber

with pdfplumber.open("test_docs/WIMPY STAFF LIST 2026.pdf") as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"--- Page {i+1} ---")
        text = page.extract_text()
        print(text)
        tables = page.extract_tables()
        if tables:
            print("Tables:")
            for t in tables:
                for row in t:
                    print(row)
