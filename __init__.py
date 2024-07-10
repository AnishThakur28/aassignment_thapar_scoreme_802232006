from .converter import extract_tables_from_pdf, write_to_excel
import sys

def main():
    if len(sys.argv) < 3:
        print("Usage: python -m pdf_to_excel <pdf_file1> <pdf_file2> ... <output_excel>")
        sys.exit(1)
    
    pdf_paths = sys.argv[1:-1]
    output_excel = sys.argv[-1]
    
    tables = []
    for pdf_path in pdf_paths:
        tables.extend(extract_tables_from_pdf(pdf_path))
    
    write_to_excel(tables, output_excel)

if __name__ == "__main__":
    main()
