import argparse,json,re
from datetime import date
from hashlib import sha256
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from template_cache import sha

STEPS_HEADING = '一、实验步骤、截图记录及遇到问题的处理过程'
SUMMARY_HEADING = '二、实验总结'


def check_sections(document, data):
    expected = [('paragraph', data['fields']['division'])]
    figure_number = 0
    for number, section in enumerate(data['sections'], 1):
        expected.append(('paragraph', f"{number}. {section['title']}"))
        expected.extend(('paragraph', text) for text in section['paragraphs'])
        expected.extend(('paragraph', text) for text in section.get('commands', []))
        if 'results' in section:
            table = section['results']
            expected.append(('table', [table['headers'], *table['rows']]))
        for figure in section.get('figures', []):
            figure_number += 1
            expected.append(('picture', sha(figure['path'])))
            expected.append(('paragraph', f"图 {figure_number} {figure['caption']}"))

    actual = []
    in_steps = False
    for element in document.element.body:
        if element.tag == qn('w:tbl') and in_steps:
            table = Table(element, document)
            actual.append(('table', [[cell.text for cell in row.cells] for row in table.rows]))
        if element.tag != qn('w:p'):
            continue
        paragraph = Paragraph(element, document)
        if paragraph.text.strip() == STEPS_HEADING:
            in_steps = True
            continue
        if paragraph.text.strip() == SUMMARY_HEADING:
            return in_steps and actual == expected
        if in_steps:
            for image in paragraph._p.xpath('.//wp:inline//a:blip'):
                part = paragraph.part.related_parts.get(image.get(qn('r:embed')))
                actual.append(('picture', sha256(part.blob).hexdigest() if part else None))
            if paragraph.text.strip():
                actual.append(('paragraph', paragraph.text))
    return False


def validate(report,rules,data):
 d=Document(report);r=json.loads(Path(rules).read_text(encoding='utf-8'));f=data['fields'];text='\n'.join(p.text for p in d.paragraphs);errors=[]
 def ck(ok,msg):
  if not ok:errors.append(msg)
 req=r['requirements'];ck(sha(r['template'])==r['sha256'],'模板缓存失效')
 try:
  steps=text.split(STEPS_HEADING,1)[1].split(SUMMARY_HEADING,1)[0];summary=text.split(SUMMARY_HEADING,1)[1]
 except IndexError:steps=summary='';errors.append('章节缺失')
 declaration=req['declaration']
 # Exclude the entire declaration block, including its date, from summary count.
 stops=[summary.find(s) for s in declaration if s in summary]
 if stops:summary=summary[:min(stops)]
 positions=[text.find(s) for s in declaration]
 ck(all(text.count(s)==1 for s in declaration) and positions==sorted(positions),'承诺声明缺失、重复或顺序错误')
 expected_date=str(f.get('declaration_date',''))
 m=re.fullmatch(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日',expected_date)
 valid_date=False
 if m:
  try:date(*map(int,m.groups()));valid_date=True
  except ValueError:pass
 ck(valid_date,'声明日期无效')
 tail=text.split(declaration[-1],1)[-1] if declaration[-1] in text else ''
 ck(valid_date and tail.strip()==expected_date,'声明日期未正确填写或存在残留内容')
 count=lambda x:len(re.findall(r'[\u4e00-\u9fffA-Za-z0-9]',x))
 body = d.paragraphs[next((i for i,p in enumerate(d.paragraphs) if p.text.strip()==STEPS_HEADING),len(d.paragraphs)):next((i for i,p in enumerate(d.paragraphs) if p.text.strip()==SUMMARY_HEADING),len(d.paragraphs))]
 ck(count(steps)>=req['steps_min'],'步骤字数不足');ck(count(summary)>=req['summary_min'],'总结字数不足');ck(sum(len(p._p.xpath('.//wp:inline')) for p in body)>=req['images_min'],'截图不足')
 if 'sections' in data:
  ck(check_sections(d,data),'实验小节正文、命令、结果表或截图与图注的内容或顺序不一致')
 full=text+'\n'+'\n'.join(c.text for t in d.tables for row in t.rows for c in row.cells)
 for k in req['required_fields']:ck(bool(str(f.get(k,'')).strip()) and str(f.get(k,'')) in full,'必填内容缺失:'+k)
 for token in req['forbidden']:ck(token not in full,'残留说明:'+token)
 expected='+'.join(f.get(k,'') for k in ['course','team','name','school'])
 ck(Path(report).stem==expected,'文件命名错误');pdf=Path(report).with_suffix('.pdf');ck(pdf.is_file() and pdf.read_bytes().startswith(b'%PDF-'),'PDF 缺失或格式错误')
 return errors
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--report',required=True);p.add_argument('--rules',default='.workflow/template_rules.json');p.add_argument('--data',required=True);p.add_argument('--json-out',required=True);a=p.parse_args();errors=validate(a.report,a.rules,json.loads(Path(a.data).read_text(encoding='utf-8')));Path(a.json_out).write_text(json.dumps(dict(errors=errors,program_pass=not errors,visual_status='待验证',note='程序通过不代表真实性、任务覆盖及视觉通过'),ensure_ascii=False,indent=2),encoding='utf-8');raise SystemExit(bool(errors))
