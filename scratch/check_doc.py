import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'แผนที่สมบูรณ์สำหรับแทนนี่/025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.docx')

for i in range(25):
    p = doc.paragraphs[i]
    xml = p._element.xml
    has_break = 'w:type="page"' in xml or 'w:br' in xml
    print(f"P{i:02d} [break={has_break}]: repr={repr(p.text)}")
