"""
Command-line interface for PDF to DOCX conversion
Uses PDFToDocxConverter from pipeline.py
"""

import sys
import os
from pathlib import Path
from generalised_converter import PDFToDocxConverter


def main():
    print("=" * 60)
    print("🎯 PDF to DOCX Converter - Command Line")
    print("=" * 60)

    if len(sys.argv) < 2:
        print("\n❌ Error: No PDF file specified")
        print("\n📋 Usage:")
        print("   python cli.py input.pdf [output.docx]")
        print("\n💡 Examples:")
        print('   python cli.py "django_assignment.pdf"')
        print('   python cli.py "document.pdf" "output.docx"')
        print("\n📦 Batch mode:")
        print('   python pipeline.py --batch ./pdfs')
        return

    pdf_path = sys.argv[1]
    output_path = (
        sys.argv[2] if len(sys.argv) > 2 else str(Path(pdf_path).with_suffix(".docx"))
    )

    if not os.path.exists(pdf_path):
        print(f"\n❌ Error: PDF file not found: {pdf_path}")
        print(f"💡 Make sure the file exists and the path is correct.")
        return

    print(f"\n📄 Input:  {pdf_path}")
    print(f"📄 Output: {output_path}\n")

    # Initialize converter
    converter = PDFToDocxConverter()

    # Convert
    success, result = converter.convert_single(pdf_path, output_path, verbose=True)

    if success:
        print(f"\n✅ SUCCESS! Converted document saved.")
        print(f"📂 Open: {os.path.abspath(output_path)}")
    else:
        print(f"\n❌ Conversion failed: {result}")


if __name__ == "__main__":
    main()



 