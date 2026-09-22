import os
import win32com.client

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

docx_path = os.path.abspath(r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.docx')
test_pdf_path = os.path.abspath(r'scratch\test_export.pdf')

doc = word.Documents.Open(docx_path)
doc.SaveAs(test_pdf_path, FileFormat=17) # 17 = wdFormatPDF
doc.Close(False)
word.Quit()

print('Exported successfully! Size:', os.path.getsize(test_pdf_path))
