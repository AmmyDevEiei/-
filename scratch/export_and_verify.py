import os
import sys
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

docx_path = os.path.abspath(r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.docx')
pdf_path = os.path.abspath(r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.pdf')

print("Converting docx to pdf via Word COM...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False

try:
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdFormatPDF
    doc.Close(False)
    print("Word PDF conversion completed successfully!")
except Exception as e:
    print("Error during conversion:", e)
finally:
    word.Quit()

# Verify PDF with fitz
if os.path.exists(pdf_path):
    pdf = fitz.open(pdf_path)
    print(f"PDF verified: {len(pdf)} pages, size {os.path.getsize(pdf_path)} bytes")
    for i in range(len(pdf)):
        text = pdf[i].get_text()
        print(f"Page {i+1}: length {len(text)} chars, sample: {repr(text[:60])}")
else:
    print("PDF not found!")
