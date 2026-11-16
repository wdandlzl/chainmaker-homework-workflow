import argparse,json,subprocess,re
from datetime import date
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from template_cache import sha
FIELDS=['course','team','name','school','contact','team_en','date','declaration_date','partners1','partners2','project','equipment','division']


def write_sections(document, anchor, sections):
    page = document.sections[0]
    content_width = page.page_width - page.left_margin - page.right_margin

    def compact(paragraph):
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        formatting = paragraph.paragraph_format
        formatting.first_line_indent = Pt(0)
        paragraph._p.get_or_add_pPr().get_or_add_ind().set(qn('w:firstLineChars'), '0')
        formatting.line_spacing = 1.15
        formatting.space_after = Pt(3)

    figure_number = 0
    for number, section in enumerate(sections, 1):
        title = document.add_paragraph()
        title.add_run(f"{number}. {section['title']}").bold = True
        title.paragraph_format.keep_with_next = True
        anchor.addnext(title._p)
        anchor = title._p
        for body in section['paragraphs']:
            paragraph = document.add_paragraph(body)
            paragraph.paragraph_format.widow_control = True
            anchor.addnext(paragraph._p)
            anchor = paragraph._p
        for command in section.get('commands', []):
            paragraph = document.add_paragraph()
            compact(paragraph)
            paragraph.paragraph_format.keep_together = True
            font = paragraph.add_run(command).font
            font.name = 'Consolas'
            font.size = Pt(10)
            anchor.addnext(paragraph._p)
            anchor = paragraph._p
        if 'results' in section:
            results = section['results']
            headers = results['headers']
            table = document.add_table(rows=0, cols=len(headers))
            table.autofit = False
            if len(headers) == 2:
                table.columns[0].width = int(content_width * 0.24)
                table.columns[1].width = content_width - table.columns[0].width
            borders = OxmlElement('w:tblBorders')
            for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                border = OxmlElement(f'w:{edge}')
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '4')
                border.set(qn('w:color'), 'BFBFBF')
                borders.append(border)
            table._tbl.tblPr.append(borders)
            for row_number, values in enumerate([headers, *results['rows']]):
                if len(values) != len(headers):
                    raise ValueError('结果表列数不一致')
                row = table.add_row()
                properties = row._tr.get_or_add_trPr()
                properties.append(OxmlElement('w:cantSplit'))
                if row_number == 0:
                    properties.append(OxmlElement('w:tblHeader'))
                for cell, value in zip(row.cells, values):
                    paragraph = cell.paragraphs[0]
                    compact(paragraph)
                    run = paragraph.add_run(value)
                    run.bold = row_number == 0
                    run.font.size = Pt(11)
            anchor.addnext(table._tbl)
            anchor = table._tbl
        for figure in section.get('figures', []):
            figure_number += 1
            picture = document.add_paragraph()
            picture.alignment = 1
            picture.paragraph_format.keep_with_next = True
            picture.add_run().add_picture(figure['path'], width=min(content_width, Inches(6)))
            caption = document.add_paragraph()
            caption.alignment = 1
            caption.add_run(f"图 {figure_number} {figure['caption']}").font.size = Pt(10)
            anchor.addnext(picture._p)
            picture._p.addnext(caption._p)
            anchor = caption._p


def generate(src,rules,data,out):
 r=json.loads(Path(rules).read_text(encoding='utf-8')); f=data['fields']; out=Path(out)
 if r['sha256']!=sha(src) or not r['requirements']: raise ValueError('缓存无效')
 if any(not str(f.get(k,'')).strip() for k in FIELDS): raise ValueError('必填字段缺失')
 if out.exists(): raise FileExistsError(out)
 d=Document(src); ps=list(d.paragraphs)
 req=r['requirements']
 declaration=req['declaration']
 if [p.text.strip() for p in ps if p.text.strip() in declaration]!=declaration: raise ValueError('模板声明与规则不一致')
 match=re.fullmatch(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日',f['declaration_date'])
 if not match: raise ValueError('声明日期格式错误')
 date(*map(int,match.groups()))
 def fill(p,text):
  if not p.runs:p.add_run()
  p.runs[0].text=text
  p.runs[0].font.color.rgb=RGBColor(0,0,0)
  for run in p.runs[1:]:run.text=''
 steps = f['division'] if 'sections' in data else f['division']+'\n'+data['steps']
 for i,v in {3:'《'+f['project']+'》实验报告',7:f['course'],10:'学校：'+f['school'],11:'队伍：'+f['team'],12:'姓名：'+f['name'],15:f['date'],19:steps,25:data['summary']}.items():fill(ps[i],v)
 # Only remove reviewed instruction/example paragraphs; never delete the tail wholesale.
 for i in req['remove_paragraph_indices']:
  p=ps[i];p._element.getparent().remove(p._element)
 declaration_start=next(i for i,p in enumerate(ps) if p.text.strip()==declaration[0])
 date_slots=[p for p in ps[declaration_start+len(declaration):] if re.fullmatch(r'\s*\d{4}\s*年\s*月\s*日\s*',p.text)]
 if len(date_slots)!=1: raise ValueError('声明日期槽位不唯一')
 fill(date_slots[0],f['declaration_date'])
 anchor=ps[19]._p
 if 'sections' in data:
  write_sections(d, anchor, data['sections'])
  declaration_start = ps[declaration_start]
  for neighbor in (declaration_start._p, date_slots[0]._p):
   previous = neighbor.getprevious()
   while previous is not None and previous.tag == qn('w:p') and not previous.xpath(
       './/w:t | .//w:drawing | .//w:pict | .//w:sectPr | .//w:br'):
    earlier = previous.getprevious()
    previous.getparent().remove(previous)
    previous = earlier
 else:
  for img in data['images']:
   p=d.add_paragraph();p.alignment=1;p.add_run().add_picture(str(img),width=Inches(5));anchor.addnext(p._p);anchor=p._p
 ps[24].paragraph_format.keep_with_next=True
 t=d.tables[0]
 for row,col,key in [(0,1,'school'),(0,3,'team'),(0,5,'name'),(0,7,'contact'),(1,1,'course'),(1,3,'team_en'),(1,5,'date'),(1,7,'partners1'),(2,7,'partners2'),(3,1,'project'),(4,1,'equipment')]:
  cell=t.cell(row,col);fill(cell.paragraphs[0],f[key])
  for extra in list(cell.paragraphs)[1:]:extra._p.getparent().remove(extra._p)
 out.parent.mkdir(parents=True,exist_ok=True);d.save(out)
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--template',required=True);p.add_argument('--rules',default='.workflow/template_rules.json');p.add_argument('--data',required=True);p.add_argument('--out',required=True);p.add_argument('--pdf',action='store_true');a=p.parse_args();generate(a.template,a.rules,json.loads(Path(a.data).read_text(encoding='utf-8')),a.out)
 if a.pdf:subprocess.run(['powershell','-NoProfile','-File',str(Path(__file__).with_name('export_pdf.ps1')),'-Docx',str(Path(a.out).resolve())],check=True)
