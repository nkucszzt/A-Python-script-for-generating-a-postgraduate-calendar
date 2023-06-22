from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, timedelta
from docx.oxml.ns import qn
from docx.shared import Pt 


def calculate_date(seed):
    start_date = datetime(2023, 6, 22)
    delta = timedelta(days=seed)
    target_date = start_date + delta
    return target_date.strftime('%Y-%m-%d')
def calculate_thedate(seed):
    mydate=['一','二','三','四','五','六','天']
    a=mydate[seed]
    return a
        
document = Document()
document.styles['Normal'].font.name = 'Times New Roman'
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
document.styles['Normal'].font.size = Pt(20)
i=0
for i in range(185):
    p=document.add_paragraph('\n考研历\n\n星期' + calculate_thedate((i+3)%7)+'\n\n'+ calculate_date(i)+'\n\n倒计时：'+str(184-i))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_page_break()
    i=i+1

document.save('demo.docx')