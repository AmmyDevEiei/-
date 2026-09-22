import os
import sys
import shutil
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

sys.stdout.reconfigure(encoding='utf-8')

# Helper: format runs with TH SarabunPSK
def add_run(p, text, font_name="TH SarabunPSK", font_size=16, bold=False, italic=False, color=None):
    r = p.add_run(text)
    r.font.name = font_name
    r.font.size = Pt(font_size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = RGBColor(*color)
    rPr = r._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    return r

# Helper: add paragraph with spacing and font
def add_para(doc, text="", font_size=16, bold=False, italic=False, color=None, 
             align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=4, 
             line_spacing=1.15, keep_with_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.keep_with_next = keep_with_next
    if text:
        add_run(p, text, font_size=font_size, bold=bold, italic=italic, color=color)
    return p

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D0D5DD", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build():
    src_docx = r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.docx'
    src_pdf = r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.pdf'
    
    # Backups
    shutil.copy2(src_docx, src_docx + '.bak2')
    shutil.copy2(src_pdf, src_pdf + '.bak2')
    print("Backups created.")

    doc = docx.Document(src_docx)
    
    # Keep paragraphs 0 to 37 (Cover, คำชี้แจง, 1. ชื่อสื่อ, 2. วัตถุประสงค์, 3. การประยุกต์ใช้)
    # Remove paragraphs 38 onwards
    body = doc._body._element
    for p in list(doc.paragraphs[38:]):
        body.remove(p._element)

    print("Existing section 4 & 5 removed. Adding updated rich content...")

    # Update P26 if needed (make sure duration matches 3.35 นาที)
    # P26 is: รูปแบบสื่อ: วิดีโอแอนิเมชันขนาดความละเอียด Full HD (1080p) สัดส่วน 16:9 แนวนอน ความยาวประมาณ 3.30 นาที
    p26 = doc.paragraphs[26]
    p26.text = ""
    add_run(p26, "รูปแบบสื่อ: วิดีโอแอนิเมชันขนาดความละเอียด Full HD (1080p) สัดส่วน 16:9 แนวนอน ความยาวประมาณ 3.35 นาที", font_size=16)

    # =========================================================================
    # 4. ขั้นตอนการผลิตสื่อ
    # =========================================================================
    p_sec4 = add_para(doc, "4. ขั้นตอนการผลิตสื่อ", font_size=18, bold=True, space_before=12, space_after=6, keep_with_next=True)

    p_intro = add_para(doc, 
        "การผลิตสื่อวิดีโอการเรียนรู้ เรื่อง \"ไขความลับพีทาโกรัส: ตอน พิชิตการหาด้านตรงข้ามมุมฉาก\" เพื่อใช้จัดการเรียนรู้ล่วงหน้านอกห้องเรียน (Out-of-Class Learning) ตามแนวคิดห้องเรียนกลับด้าน (Flipped Classroom) ความยาวรวมประมาณ 3.35 นาที ได้บูรณาการเครื่องมือปัญญาประดิษฐ์ (Generative AI) ร่วมกับซอฟต์แวร์ตัดต่อสมัยใหม่ โดยดำเนินงานตามขั้นตอนการผลิตอย่างเป็นระบบ ดังนี้:",
        font_size=16, space_before=0, space_after=6)

    # -------------------------------------------------------------------------
    # 4.1 ภาพรวมขั้นตอนและกระบวนการผลิตวิดีโอ (Production Workflow)
    # -------------------------------------------------------------------------
    add_para(doc, "4.1 ภาพรวมขั้นตอนและกระบวนการผลิตวิดีโอ (Production Workflow)", font_size=16, bold=True, space_before=8, space_after=4, keep_with_next=True)
    
    add_para(doc, "การสร้างวิดีโอการสอนความยาวประมาณ 3-4 นาทีด้วย AI ให้ได้คุณภาพสูง ชัดเจน และตรงตามหลักสูตร จำเป็นต้องแยกทำทีละฉาก (Scene-by-Scene) และใช้คำสั่งที่จำเพาะเจาะจง โดยมี 4 ขั้นตอนหลักในการดำเนินงาน:", font_size=16, space_after=4)

    # Workflow 1
    p_wf1 = add_para(doc, "1) การเตรียมเสียงพากย์และบทพูด (Voiceover & Audio):", font_size=16, bold=True, space_before=4, space_after=2, keep_with_next=True)
    p_wf1_sub = add_para(doc, 
        "• บันทึกและสังเคราะห์เสียงพากย์ของ \"ครูณัฐ\" ทีละฉาก โดยใช้เทคโนโลยี AI Text-to-Speech (เช่น CapCut AI TTS / ElevenLabs) ซึ่งให้เสียงภาษาไทยที่เป็นธรรมชาติ คุมโทนเสียงให้สดใส อบอุ่น เป็นกันเอง และตื่นเต้นชวนติดตามตามบทบาทครูผู้สอน\n"
        "• จัดเตรียมเสียงประกอบ (Sound Effects: SFX) เช่น เสียงไซเรนสั้นๆ, เสียงแมวร้อง \"เหมียว~\", เสียง \"ปิ๊ง!\", เสียงกระดิ่ง \"กริ๊ง!\" และเพลงประกอบแนว Edutainment จังหวะสดใสคลอเบาๆ",
        font_size=15, space_before=0, space_after=4)

    # Workflow 2
    p_wf2 = add_para(doc, "2) การสร้างคลิปวิดีโอทีละฉากด้วย AI (Scene Generation):", font_size=16, bold=True, space_before=4, space_after=2, keep_with_next=True)
    p_wf2_sub = add_para(doc, 
        "• นำ Video Prompt ภาษาอังกฤษที่ผ่านการปรับแต่งเฉพาะ (Optimized AI Video Prompts) ไปใส่ในเครื่องมือสร้างวิดีโอ AI (เช่น Runway Gen-3 Alpha, Kling AI, Luma Dream Machine) เพื่อสร้างคลิปภาพเคลื่อนไหวแอนิเมชัน 3 มิติ (3D cute cartoon educational animation) ความยาวฉากละ 15 - 45 วินาที\n"
        "• ข้อแนะนำเชิงเทคนิค: การใช้คำสั่งภาษาอังกฤษในการเจนวิดีโอ AI ส่วนใหญ่จะให้ผลลัพธ์ภาพ 3 มิติที่คมชัด ท่าทางการเคลื่อนไหว และมุมกล้องที่แม่นยำกว่าภาษาอื่น จึงได้จัดทำชุด Prompt ภาษาอังกฤษสำหรับแต่ละฉากไว้โดยเฉพาะ",
        font_size=15, space_before=0, space_after=4)

    # Workflow 3
    p_wf3 = add_para(doc, "3) การซ้อนข้อความ กราฟิก และการแสดงลิขสิทธิ์ (Text & Graphics Overlay):", font_size=16, bold=True, space_before=4, space_after=2, keep_with_next=True)
    p_wf3_sub = add_para(doc, 
        "• นำคลิปวิดีโอที่ได้จากการเจน AI มารวบรวมและจัดวางในโปรแกรมตัดต่อ (CapCut PC / Canva Pro)\n"
        "• ซ้อนข้อความอธิบายความรู้บนหน้าจอ (Text & Math Overlay) เช่น สูตรคำนวณพีทาโกรัส สเต็ป 1-4 และป้ายเตือนข้อควรระวังตามที่ระบุในแต่ละฉาก\n"
        "• ข้อกำหนดสำคัญด้านลิขสิทธิ์: วางสัญลักษณ์ข้อความและรหัสนิสิต \"ณัฐณิชา 025\" ไว้ที่มุมขวาบนของหน้าจอตลอดทั้งคลิปวิดีโอ เพื่อยืนยันลิขสิทธิ์ความเป็นเจ้าของผลงานของนิสิต",
        font_size=15, space_before=0, space_after=4)

    # Workflow 4
    p_wf4 = add_para(doc, "4) การจัดจังหวะ ผสมเสียง และควบคุมความยาว (Editing & Audio Mixing):", font_size=16, bold=True, space_before=4, space_after=2, keep_with_next=True)
    p_wf4_sub = add_para(doc, 
        "• ตรวจสอบความต่อเนื่องและความลื่นไหลของภาพ แอนิเมชัน เสียงพากย์ ดนตรีคลอ และซับไตเติลภาษาไทย\n"
        "• ควบคุมและเกลี่ยความยาวรวมของสื่อวิดีโอให้อยู่ภายใน 3 นาที 35 วินาที (อยู่ในเกณฑ์ 3 - 4 นาที และไม่เกิน 5 นาทีตามกำหนดอย่างเคร่งครัด)",
        font_size=15, space_before=0, space_after=6)

    # -------------------------------------------------------------------------
    # 4.2 การประยุกต์ใช้เครื่องมือเทคโนโลยีและปัญญาประดิษฐ์ (AI Tools & Technologies)
    # -------------------------------------------------------------------------
    add_para(doc, "4.2 การประยุกต์ใช้เครื่องมือเทคโนโลยีและปัญญาประดิษฐ์ (AI Tools & Applications)", font_size=16, bold=True, space_before=8, space_after=4, keep_with_next=True)
    
    # Table of AI tools
    tbl_tools = doc.add_table(rows=6, cols=3)
    set_table_borders(tbl_tools)
    tbl_tools.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers = ["ขั้นตอนการทำงาน", "เครื่องมือ / ซอฟต์แวร์ที่ใช้", "วัตถุประสงค์และการประยุกต์ใช้"]
    for c_idx, h in enumerate(headers):
        cell = tbl_tools.cell(0, c_idx)
        set_cell_shading(cell, "F2F4F7")
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, h, font_size=15, bold=True)

    tools_data = [
        ("1. การวางแผนและเขียนบทพูด\n(Script Writing)", "Google Gemini &\nOpenAI ChatGPT (GPT-4o)", "กำหนดโครงเรื่อง 3 ช่วง (Intro -> Body -> Outro) ยกร่างบทพากย์คำสอนครูณัฐ และคำนวณเวลาให้กระชับเหมาะสมกับเด็ก ม.2"),
        ("2. การสร้างภาพนิ่งและคีย์เฟรม\n(Visual Concept)", "Canva Magic Media &\nCanva Pro", "สร้างภาพร่างต้นแบบตัวละครคุณครู ภาพจำลองรถดับเพลิง และไดอะแกรมเรขาคณิตสามเหลี่ยมมุมฉาก"),
        ("3. การสร้างวิดีโอแอนิเมชัน AI\n(Scene Generation)", "Runway Gen-3 Alpha /\nKling AI / Luma Dream Machine", "นำชุดคำสั่ง Video Prompt ภาษาอังกฤษไปสร้างคลิปแอนิเมชัน 3D การ์ตูนเพื่อการศึกษาทีละฉาก (ความละเอียด Full HD 1080p)"),
        ("4. การสังเคราะห์เสียงพากย์ AI\n(Voiceover Generation)", "CapCut AI Text-to-Speech\n(เสียงครูใจดี) / ElevenLabs", "แปลงบทพูดภาษาไทยเป็นเสียงบรรยายที่มีความเป็นธรรมชาติ ชัดถ้อยชัดคำ เว้นจังหวะเน้นย้ำสูตรและจุดสำคัญได้อย่างถูกต้อง"),
        ("5. การตัดต่อ กราฟิก และเรนเดอร์\n(Editing & Compositing)", "CapCut PC &\nCanva Pro", "รวมคลิปวิดีโอ ซ้อนภาพ ซ้อนสูตรคณิตศาสตร์ ใส่ซับไตเติล ผสมเสียงดนตรี/SFX ฝังลายน้ำ 'ณัฐณิชา 025' และเรนเดอร์เป็นไฟล์ MP4")
    ]

    for r_idx, row_data in enumerate(tools_data, 1):
        for c_idx, text in enumerate(row_data):
            cell = tbl_tools.cell(r_idx, c_idx)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F9FAFB")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, text, font_size=14, bold=(c_idx==0))

    # Set column widths
    col_widths = [Inches(1.8), Inches(1.8), Inches(2.9)]
    for row in tbl_tools.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    add_para(doc, "", font_size=10, space_before=0, space_after=4)

    # -------------------------------------------------------------------------
    # 4.3 ตารางสรุปภาพรวมโครงสร้างสื่อวิดีโอ (Video Storyboard Overview)
    # -------------------------------------------------------------------------
    add_para(doc, "4.3 ตารางสรุปภาพรวมโครงสร้างสื่อวิดีโอ (Video Storyboard Overview)", font_size=16, bold=True, space_before=8, space_after=4, keep_with_next=True)
    
    tbl_sb = doc.add_table(rows=7, cols=5)
    set_table_borders(tbl_sb)
    tbl_sb.alignment = WD_TABLE_ALIGNMENT.CENTER

    sb_headers = ["ฉากที่", "ช่วงเวลา", "ชื่อฉาก / สาระการเรียนรู้", "รูปแบบภาพและแอนิเมชัน AI", "เสียงบรรยายและเสียงประกอบ"]
    for c_idx, h in enumerate(sb_headers):
        cell = tbl_sb.cell(0, c_idx)
        set_cell_shading(cell, "EBF3FB")
        set_cell_margins(cell, top=120, bottom=120, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, h, font_size=15, bold=True)

    sb_rows = [
        ("ฉากที่ 1", "0:00 – 0:35\n(35 วิ)", "เปิดภารกิจกู้ภัยตึกสูง\n(กระตุ้นความสนใจ/ตั้งปัญหา)", "3D Animation รถดับเพลิงยืดบันไดช่วยลูกแมวส้มบนระเบียงตึกชั้น 3", "เสียงครูสดใสทักทาย ชวนคิดหาความยาวบันได / เสียงไซเรน และเสียงแมวเหมียว"),
        ("ฉากที่ 2", "0:35 – 1:15\n(40 วิ)", "วาดรูปจำลอง & หาด้าน c\n(ทักษะแปลงโจทย์)", "Motion Graphics ตึกแปลงเป็นสามเหลี่ยมมุมฉาก ลูกศรชี้พุ่งไปหาด้าน c", "เสียงอบอุ่นสอนเทคนิคสังเกตมุมฉาก ด้าน c อยู่ตรงข้ามและยาวที่สุด / เสียงปิ๊ง!"),
        ("ฉากที่ 3", "1:15 – 1:55\n(40 วิ)", "เปิดกล่องสูตรลับ & ชุดตัวเลข\n(ที่มาสูตรสำเร็จ)", "บล็อกสี่เหลี่ยมบนด้าน a และ b รวมตัวเป็นบล็อกสีทองบนด้าน c, สูตรแปลงรูป", "เสียงสอนสูตรสำเร็จรูป c = √(a² + b²) และแนะนำชุดตัวเลข 3-4-5 และ 9-12-15"),
        ("ฉากที่ 4", "1:55 – 2:45\n(50 วิ)", "สาธิต 4 สเต็ปทองคำ\n(ขั้นตอนการคำนวณ)", "Split Screen ด้านซ้ายใบงาน ภารกิจที่ 1 ด้านขวากระดานเขียนวิธีทำทีละสเต็ป", "เสียงครูอธิบาย 4 สเต็ปทองคำคำนวณละเอียด a=9, b=12 ได้ c=15 เมตร / เสียงปากกาเขียน"),
        ("ฉากที่ 5", "2:45 – 3:10\n(25 วิ)", "กับดักที่เด็ก ม.2 ชอบพลาด!\n(ข้อควรระวัง)", "การ์ตูนปั๊มตรากากบาทสีแดงบนสูตรผิด a+b=c พร้อมป้ายเตือนระวังสีเหลือง", "เสียงตื่นเต้นเตือน 2 จุดพลาด: ห้ามบวกตรงๆ ต้องยกกำลังสอง และอย่าลืมถอดรูท"),
        ("ฉากที่ 6", "3:10 – 3:35\n(25 วิ)", "ภารกิจก่อนเข้าห้องเรียน\n(สรุปและมอบหมายงาน)", "นักเรียน 5 บทบาทนั่งทำงานกลุ่มร่วมกัน มีมือถือเปิด Padlet และครูโบกมือลา", "เสียงอบอุ่นมอบหมาย 2 ภารกิจ: สรุป Mind Map 1 หน้า และตั้งคำถามลง Padlet")
    ]

    for r_idx, r_data in enumerate(sb_rows, 1):
        for c_idx, text in enumerate(r_data):
            cell = tbl_sb.cell(r_idx, c_idx)
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F9FAFB")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 1) else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, text, font_size=14, bold=(c_idx==0))

    sb_widths = [Inches(0.9), Inches(1.1), Inches(1.7), Inches(1.9), Inches(1.9)]
    for row in tbl_sb.rows:
        for idx, width in enumerate(sb_widths):
            row.cells[idx].width = width

    add_para(doc, "", font_size=10, space_before=0, space_after=4)

    # -------------------------------------------------------------------------
    # 4.4 รายละเอียดชุดคำสั่ง AI และการผลิตสื่อแยกรายฉาก (Scene-by-Scene Prompts)
    # -------------------------------------------------------------------------
    add_para(doc, "4.4 รายละเอียดชุดคำสั่ง AI และการผลิตสื่อแยกรายฉาก (Scene-by-Scene Production & AI Video Prompts)", font_size=16, bold=True, space_before=8, space_after=4, keep_with_next=True)
    
    add_para(doc, 
        "รายละเอียดชุดคำสั่ง (Prompt) ที่ใช้ในการสร้างสรรค์ภาพเคลื่อนไหว บทพากย์เสียง องค์ประกอบกราฟิก และเอฟเฟกต์เสียงในแต่ละฉาก มีรายละเอียดดังนี้:",
        font_size=16, space_after=6)

    # Function to add a styled scene container
    def add_scene_detail(scene_no, title, duration, video_prompt, voiceover, graphics, sfx, sample_img=None, img_caption=""):
        p_hdr = add_para(doc, f"ฉากที่ {scene_no}: {title} (ช่วงเวลา {duration})", font_size=16, bold=True, color=(15, 23, 42), space_before=8, space_after=3, keep_with_next=True)
        
        # Single cell callout table for scene details
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        set_cell_shading(cell, "F8FAFC")
        set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
        set_table_borders(tbl, color="CBD5E1", sz="6")
        cell.width = Inches(6.5)

        # Video Prompt
        p1 = cell.paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.line_spacing = 1.15
        add_run(p1, "• Video Prompt (สำหรับเครื่องมือสร้างวิดีโอ AI):\n", font_size=15, bold=True, color=(30, 41, 59))
        add_run(p1, f'"{video_prompt}"', font_size=14, italic=True, color=(51, 65, 85))

        # Voiceover
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(4)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.line_spacing = 1.15
        add_run(p2, "• เสียงพากย์ (Voiceover ของครูณัฐ):\n", font_size=15, bold=True, color=(30, 41, 59))
        add_run(p2, voiceover, font_size=15, color=(15, 23, 42))

        # Graphics
        p3 = cell.add_paragraph()
        p3.paragraph_format.space_before = Pt(4)
        p3.paragraph_format.space_after = Pt(2)
        p3.paragraph_format.line_spacing = 1.15
        add_run(p3, "• กราฟิกและข้อความบนจอ (On-Screen Graphics): ", font_size=15, bold=True, color=(30, 41, 59))
        add_run(p3, graphics, font_size=15, color=(51, 65, 85))

        # SFX
        p4 = cell.add_paragraph()
        p4.paragraph_format.space_before = Pt(4)
        p4.paragraph_format.space_after = Pt(2)
        p4.paragraph_format.line_spacing = 1.15
        add_run(p4, "• เสียงเอฟเฟกต์และดนตรี (SFX & Music): ", font_size=15, bold=True, color=(30, 41, 59))
        add_run(p4, sfx, font_size=15, color=(51, 65, 85))

        # Optional image
        if sample_img and os.path.exists(sample_img):
            p_img = cell.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run_img = p_img.add_run()
            run_img.add_picture(sample_img, width=Inches(5.0))
            
            p_cap = cell.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(2)
            add_run(p_cap, img_caption, font_size=13, italic=True, color=(100, 116, 139))

        add_para(doc, "", font_size=6, space_before=0, space_after=2)

    # --- Scene 1 ---
    add_scene_detail(
        1, "เปิดภารกิจกู้ภัยตึกสูง", "0:00 – 0:35",
        "3D cute cartoon educational animation, bright morning atmosphere, a red fire truck parked in front of a commercial building. On the 3rd floor balcony, a cute orange kitten is crying for help. A firefighter is extending an aluminum ladder from the truck up to the balcony. Distance marker lines on the street. Smooth camera pan down from the kitten to the fire truck. 16:9 widescreen, high quality.",
        "\"[สดใส] สวัสดีค่ะนักเรียน ม.2 ทุกคน! ดูที่หน้าจอนะคะ น้องแมวส้มตัวน้อยกำลังติดอยู่บนระเบียงตึกสูง [สงสัย] ถ้าเรารู้ความสูงของตึกและระยะห่างที่รถดับเพลิงจอด พี่ๆ จะต้องยืดบันไดยาวกี่เมตรถึงจะพาดช่วยน้องแมวได้พอดี? [ยิ้ม] และนี่คือสถานการณ์จริงในภารกิจ Pythagoras Mission ที่เราจะทำร่วมกันในคาบเรียนพรุ่งนี้ค่ะ!\"",
        "เส้นประสีเขียว (แสดงความสูงตึก) / เส้นประสีฟ้า (แสดงระยะห่างจากตึก) / เส้นบันไดพาดเฉียงพร้อมเครื่องหมายคำถาม (?)",
        "เสียงไซเรนสั้นๆ -> เสียงแมวร้อง \"เหมียว~\" -> ดนตรีประกอบสดใสน่ารักคลอเบาๆ",
        sample_img="scratch/scene_sample_1.jpg",
        img_caption="ภาพจำลองแอนิเมชัน 3D จาก AI Video Prompt ฉากที่ 1 (พร้อมลายน้ำลิขสิทธิ์ ณัฐณิชา 025 ที่มุมบนขวา)"
    )

    # --- Scene 2 ---
    add_scene_detail(
        2, "วาดรูปจำลอง & หาด้าน c", "0:35 – 1:15",
        "3D modern clean motion graphics, a commercial building and fire truck transforming into a right-angled triangle on a math blueprint background. A 90-degree square corner symbol appears. A golden neon arrow points from the right-angle to the hypotenuse. The hypotenuse glows in gold with clear Thai text overlay \"ด้านตรงข้ามมุมฉาก (c)\". Clean and minimalist style.",
        "\"[อบอุ่น] ทักษะแรกที่สำคัญคือการแปลงโจทย์ให้เป็นรูปเรขาคณิต กำแพงตึกกับพื้นทำมุม 90 องศาเสมอ [ภาคภูมิใจ] เทคนิคจำง่ายๆ แค่มองหาสัญลักษณ์มุมฉาก แล้วมองพุ่งตรงข้ามไปทันที! ด้านที่บันไดพาดคือ 'ด้านตรงข้ามมุมฉาก' หรือด้าน c เสมอ และจะเป็นด้านที่ยาวที่สุดในรูปเสมอค่ะ!\"",
        "ป้ายข้อความ: \"ด้าน c = ด้านตรงข้ามมุมฉาก และยาวที่สุด\" / สัญลักษณ์มุมฉาก 90° ⟶ ลูกศรสีทองชี้พุ่งไปที่ด้าน c",
        "เสียง \"ปิ๊ง!\" ตอนลูกศรชี้โดนด้านตรงข้ามมุมฉาก c"
    )

    # --- Scene 3 ---
    add_scene_detail(
        3, "เปิดกล่องสูตรลับ & ชุดตัวเลข", "1:15 – 1:55",
        "3D math educational motion graphics. A right-angled triangle in the center with a blue square block on side a and green square block on side b. Both blocks merge together onto the hypotenuse c, forming a large gold block. Glowing math formula transforms from c-squared = a-squared + b-squared into c = square root of (a-squared + b-squared). A banner showing number set 3-4-5 expanding to 9-12-15.",
        "\"[อบอุ่น] ทฤษฎีบทพีทาโกรัสบอกว่า c² = a² + b² แต่ในใบกิจกรรมเราต้องการความยาวด้าน c ตัวเดียว จึงถอดรากได้สูตรสำเร็จรูป c = √(a² + b²) ค่ะ! [ตื่นเต้น] และจำชุดตัวเลขยอดฮิตอย่าง 3-4-5 และ 9-12-15 ไว้ให้ดี จะช่วยให้ทำใบงานได้เร็วขึ้นมากค่ะ!\"",
        "สูตรคำนวณ: c² = a² + b² ⟶ c = √(a² + b²) / ป้ายคู่หูชุดตัวเลขยอดนิยม: (3, 4, 5) ⟶ (9, 12, 15)",
        "เสียงเวทมนตร์กล่องรวมตัวกัน -> เสียงปิ๊งตอนสูตรสำเร็จรูปปรากฏ"
    )

    # --- Scene 4 ---
    add_scene_detail(
        4, "สาธิต 4 สเต็ปทองคำ", "1:55 – 2:45",
        "Split screen online teaching style. Left side shows a cartoon worksheet paper with Thai title \"ภารกิจพิชิตพีทาโกรัส ภารกิจที่ 1\". Right side shows a clean white board with a cartoon pen writing step-by-step math solutions in Thai clearly: formula, substitution, calculation, and final square root answer with meters unit.",
        "\"[ยิ้ม] มาดูวิธีเขียนแสดงวิธีทำ 4 สเต็ปทองคำเพื่อให้ได้เต็ม 10 คะแนนค่ะ:\n"
        "1. สเต็ปที่ 1: เขียนสูตรตั้งต้น c = √(a² + b²)\n"
        "2. สเต็ปที่ 2: แทนค่าความยาวด้าน\n"
        "3. สเต็ปที่ 3: ยกกำลังสองแล้วบวกกัน (c = √(9² + 12²) = √(81 + 144) = √225)\n"
        "4. สเต็ปที่ 4: ถอดรากที่สองเป็นคำตอบสุดท้ายพร้อมใส่หน่วย (c = 15 เมตร)\"",
        "แสดง 4 สเต็ปทองคำบนกระดานตามโครงสร้างใบงานอย่างเป็นระเบียบชัดเจน",
        "เสียงปากกาเขียนบนกระดาน -> เสียงกระดิ่ง \"กริ๊ง!\" เมื่อคำนวณได้คำตอบสุดท้าย",
        sample_img="scratch/scene_sample_2.jpg",
        img_caption="ภาพจำลองหน้าจอการสอนสาธิต 4 สเต็ปทองคำและการจัดวางกราฟิก (พร้อมลายน้ำลิขสิทธิ์ ณัฐณิชา 025)"
    )

    # --- Scene 5 ---
    add_scene_detail(
        5, "กับดักที่เด็ก ม.2 ชอบพลาด!", "2:45 – 3:10",
        "Playful cartoon warning visual. A big red cross stamp over a wrong formula example \"a + b = c\". Background shaking slightly with a flashing yellow warning sign. Zoom in on square root symbol with a green check stamp saying \"Do not forget square root\". Cartoon teacher shaking finger with a friendly smile.",
        "\"[ตกใจ] จุดที่นักเรียนชอบพลาดจนโดนหักคะแนนมี 2 ข้อค่ะ!\n"
        "1. ห้ามนำความยาวด้านมาบวกกันดื้อๆ เด็ดขาด ต้องยกกำลังสองก่อนเสมอ!\n"
        "2. เมื่อบวกเลขในรูทแล้ว อย่าลืมถอดรากที่สองนะคะ และที่สำคัญ... อย่าลืมใส่หน่วยกำกับทุกครั้งค่ะ!\"",
        "❌ a + b = c (ห้ามบวกตรงๆ) / ⚠️ ป้ายเตือนสีเหลือง: 1. ต้องยกกำลังสอง 2. ต้องถอดรูท 3. ต้องใส่หน่วยกำกับ",
        "เสียงสัญญาณเตือนตลกๆ -> เสียงปิ๊งเตือนสติ"
    )

    # --- Scene 6 ---
    add_scene_detail(
        6, "ภารกิจก่อนเข้าห้องเรียน", "3:10 – 3:35",
        "Heartwarming ending scene. Cartoon M.2 students in groups of 5 sitting around a table working on mission worksheets happily. Top shows 5 role tags: Leader, Drawer, Calculator, Verifier, Presenter. Right side shows a smartphone screen displaying classroom Padlet board. Top-right corner shows clear copyright text \"ณัฐณิชา 025\". Cartoon teacher waving goodbye warmly. Fade out.",
        "\"[อบอุ่น] ก่อนมาลุยใบงานจริงวันพรุ่งนี้ มี 2 ภารกิจง่ายๆ ค่ะ:\n"
        "1. สรุป 4 สเต็ปทองคำเป็นแผนผังความคิด 1 หน้า\n"
        "2. โพสต์คำถามที่สงสัย 1 ข้อลงใน Padlet ของห้องเราค่ะ\n"
        "พรุ่งนี้กลุ่มละ 5 คนมารวมพลังพิชิต Pythagoras Mission คว้าเต็ม 10 ไปด้วยกัน แล้วพบกันในห้องเรียนนะคะ สวัสดีค่ะ!\"",
        "ป้ายสรุปภารกิจก่อนเรียน 2 ข้อ / โลโก้แสดงลิขสิทธิ์มุมขวาบน: ณัฐณิชา 025",
        "ดนตรีท่อนจบสนุกสนาน บรรเลงส่งท้ายและค่อยๆ เฟดเบาลงอย่างนุ่มนวล"
    )

    # -------------------------------------------------------------------------
    # 4.5 การวางสัญลักษณ์แสดงลิขสิทธิ์ประจำตัวนิสิต (Copyright Identification)
    # -------------------------------------------------------------------------
    add_para(doc, "4.5 การวางสัญลักษณ์แสดงลิขสิทธิ์ประจำตัวนิสิต (Copyright Identification)", font_size=16, bold=True, space_before=8, space_after=4, keep_with_next=True)
    
    add_para(doc, 
        "ตามข้อกำหนดของรายวิชาในการผลิตสื่อนวัตกรรม นิสิตได้ดำเนินการฝังข้อความและสัญลักษณ์แสดงความเป็นเจ้าของลิขสิทธิ์ผลงานอย่างชัดเจน โดยกำหนดให้มีข้อความชื่อ-สกุล และรหัสนิสิต 3 ตัวท้าย \"ณัฐณิชา 025\" แสดงอยู่ที่บริเวณมุมบนด้านขวา (Top-Right Corner) ของจอภาพวิดีโอตลอดความยาวคลิปตั้งแต่ต้นจนจบ (0:00 – 3:35 นาที) เพื่อยืนยันว่าเป็นผลงานที่นิสิตได้จัดทำขึ้นด้วยตนเองอย่างถูกต้อง",
        font_size=16, space_after=8)

    # =========================================================================
    # 5. ลิงก์สื่อ Video และ QR Code สำหรับเข้าถึง
    # =========================================================================
    add_para(doc, "5. ลิงก์สื่อ Video และ QR Code สำหรับเข้าถึง", font_size=18, bold=True, space_before=12, space_after=4, keep_with_next=True)

    add_para(doc, 
        "สื่อวิดีโอได้รับการอัปโหลดขึ้นสู่ Google Drive ภายใต้บัญชีของมหาวิทยาลัยมหาสารคาม (MAHASARAKHAM UNIVERSITY) โดยตั้งค่าสิทธิ์การเข้าถึงให้เฉพาะผู้มีอีเมลมหาวิทยาลัย (@msu.ac.th) สามารถรับชมได้ตามข้อกำหนดของรายวิชา:",
        font_size=16, space_after=4)

    p_link = add_para(doc, "ลิงก์ Google Drive: ", font_size=16, bold=True, space_after=6)
    add_run(p_link, "https://drive.google.com/drive/u/0/my-drive", font_size=16, bold=True, color=(37, 99, 235))

    # QR Code image
    qr_path = "scratch/qr_code.png"
    if os.path.exists(qr_path):
        p_qr = doc.add_paragraph()
        p_qr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_qr.paragraph_format.space_before = Pt(8)
        p_qr.paragraph_format.space_after = Pt(4)
        run_qr = p_qr.add_run()
        run_qr.add_picture(qr_path, width=Inches(1.8))

    p_qrcap = add_para(doc, "(สแกน QR Code เพื่อรับชมสื่อ Video บน Google Drive บัญชี มมส.)", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    # Save docx
    output_docx = r'แผนที่สมบูรณ์สำหรับแทนนี่\025 ณัฐณิชา งานเล่มคู่มือผลิตสื่อ Video.docx'
    doc.save(output_docx)
    print(f"Updated docx saved successfully: {output_docx}")

if __name__ == "__main__":
    build()
