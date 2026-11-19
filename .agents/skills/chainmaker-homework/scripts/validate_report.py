import argparse,json,re
from datetime import date
from hashlib import sha256
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph
from template_cache import sha
from report_generator import TEXT_BLOCKS, section_blocks

STEPS_HEADING = '一、实验步骤、截图记录及遇到问题的处理过程'
SUMMARY_HEADING = '二、实验总结'


def check_sections(document, data):
    expected = [('paragraph', data['fields']['division'])]
    figure_number = 0
    for number, section in enumerate(data['sections'], 1):
        expected.append(('paragraph', f"{number}. {section['title']}"))
        for block in section_blocks(section):
            kind = block['type']
            if kind in TEXT_BLOCKS or kind == 'command':
                expected.append(('paragraph', block['text']))
            elif kind == 'table':
                expected.append(('table', [block['headers'], *block['rows']]))
            elif kind == 'figure':
                figure_number += 1
                expected.append(('picture', sha(block['path'])))
                expected.append(('paragraph', f"图 {figure_number} {block['caption']}"))

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


def check_process(document, data, errors, warnings):
    """Check recorded evidence and layout signals for the ordered input format."""
    figures = set()
    facts_by_run = {}
    if not data['sections']:
        errors.append('报告缺少实验阶段')
    for section in data['sections']:
        if 'blocks' not in section:
            continue
        label = section['title']
        blocks = section['blocks']
        kinds = [block['type'] for block in blocks]
        run_id = section.get('run_id')
        if not label.strip() or not section.get('task_ids') or not run_id:
            errors.append(f'{label}缺少阶段标题、任务编号或运行标识')
        if not {'purpose', 'expected', 'observation'} <= set(kinds) or not (
            {'command', 'operation'} & set(kinds)
        ):
            errors.append(f'{label}缺少目的、操作、预期或实测结果')
        first_operation = next((i for i, kind in enumerate(kinds)
                                if kind in {'command', 'operation'}), len(kinds))
        evidence_found = False
        for index, block in enumerate(blocks):
            kind = block['type']
            if kind == 'table' and any(
                '\n' in cell or len(cell) > 250 for row in block['rows'] for cell in row
            ):
                warnings.append(f'{label}的表格含长文本，请检查是否应引用证据文件')
            if kind in TEXT_BLOCKS or kind == 'command':
                value = block['text']
                if not value.strip():
                    errors.append(f'{label}存在空的{kind}内容')
                if (len(value.splitlines()) > 12 or re.search(r'[0-9a-fA-F]{64,}', value)):
                    warnings.append(f'{label}有长输出或标识，请检查正文是否保留了必要片段')
            if kind != 'figure' and not (kind == 'observation' and block.get('path')):
                continue
            evidence_found = True
            if block.get('run_id') != run_id:
                errors.append(f'{label}证据运行标识与阶段不一致')
            if not Path(block['path']).is_file():
                errors.append(f"{label}证据文件缺失 {block['path']}")
                continue
            if kind == 'figure':
                if not block['caption'].strip():
                    errors.append(f'{label}截图缺少图注')
                next_figure = kinds.index('figure', index + 1) if 'figure' in kinds[index + 1:] else len(kinds)
                if index <= first_operation or not (
                    {'observation', 'analysis', 'issue'} & set(kinds[index + 1:next_figure])
                ):
                    errors.append(f'{label}截图缺少前置操作或图后结果解释')
                digest = sha(block['path'])
                if digest in figures:
                    warnings.append(f'{label}存在内容相同的截图，请确认是否需要复用')
                figures.add(digest)
        if not evidence_found:
            errors.append(f'{label}缺少截图或带证据路径的实测结果')
        facts = facts_by_run.setdefault(run_id, {})
        for source in [section, *blocks]:
            for name, value in source.get('facts', {}).items():
                if facts.setdefault(name, value) != value:
                    errors.append(f'{label}同一运行的{name}记录不一致')

    in_steps = False
    page = document.sections[0]
    page_width = page.page_width - page.left_margin - page.right_margin
    page_height = page.page_height - page.top_margin - page.bottom_margin
    for index, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip() == STEPS_HEADING:
            in_steps = True
        if paragraph.text.strip() == SUMMARY_HEADING:
            break
        if not in_steps:
            continue
        for inline in paragraph._p.xpath('.//wp:inline'):
            caption = paragraph._p.getnext()
            if (
                paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER
                or not paragraph.paragraph_format.keep_with_next
                or caption is None or caption.tag != qn('w:p')
                or not re.match(r'^图 \d+ \S', Paragraph(caption, document).text)
            ):
                errors.append(f'实验正文第 {index + 1} 段截图未居中或未与图注相邻')
            extent = inline.find(qn('wp:extent'))
            blips = inline.xpath('.//a:blip')
            if extent is None or not blips:
                continue
            width, height = (int(extent.get(axis)) for axis in ('cx', 'cy'))
            if not (0 < width <= page_width and 0 < height <= page_height):
                errors.append(f'实验正文第 {index + 1} 段截图显示尺寸超出正文区域')
            if any(
                crop.get(axis, '0') != '0'
                for crop in inline.xpath('.//a:srcRect') for axis in ('l', 't', 'r', 'b')
            ):
                warnings.append(f'实验正文第 {index + 1} 段存在 Word 内裁切，请核对关键输出')
            part = paragraph.part.related_parts.get(blips[0].get(qn('r:embed')))
            if part is not None and width > 0 and height > 0:
                image = DocxImage.from_blob(part.blob)
                if abs(width / height / (image.px_width / image.px_height) - 1) > 0.01:
                    errors.append(f'实验正文第 {index + 1} 段截图宽高比例异常')
                if width < Inches(2) or image.px_width / (width / Inches(1)) < 96:
                    warnings.append(f'实验正文第 {index + 1} 段截图偏小或分辨率偏低，请查看渲染页')


def check_pdf_order(pdf, sections):
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError

    try:
        pages = PdfReader(pdf).pages
        text = re.sub(r'\s+', '', '\n'.join(page.extract_text() for page in pages))
    except (OSError, PdfReadError):
        return False
    cursor = 0
    figure_number = 0
    for number, section in enumerate(sections, 1):
        markers = [f"{number}. {section['title']}"]
        for block in section_blocks(section):
            if block['type'] == 'figure':
                figure_number += 1
                markers.append(f"图 {figure_number} {block['caption']}")
        for marker in markers:
            normalized = re.sub(r'\s+', '', marker)
            position = text.find(normalized, cursor)
            if position < 0:
                return False
            cursor = position + len(normalized)
    return True


def validate(report,rules,data,warnings=None):
 d=Document(report);r=json.loads(Path(rules).read_text(encoding='utf-8'));f=data['fields'];text='\n'.join(p.text for p in d.paragraphs);errors=[]
 if warnings is None:warnings=[]
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
  try:ck(check_sections(d,data),'实验小节正文、命令、结果表或截图与图注的内容或顺序不一致')
  except FileNotFoundError as err:errors.append(f'截图文件缺失 {err.filename}')
  if not data['sections'] or any('blocks' in section for section in data['sections']):
   check_process(d,data,errors,warnings)
 full=text+'\n'+'\n'.join(c.text for t in d.tables for row in t.rows for c in row.cells)
 for k in req['required_fields']:ck(bool(str(f.get(k,'')).strip()) and str(f.get(k,'')) in full,'必填内容缺失:'+k)
 for token in req['forbidden']:ck(token not in full,'残留说明:'+token)
 expected='+'.join(f.get(k,'') for k in ['course','team','name','school'])
 ck(Path(report).stem==expected,'文件命名错误')
 pdf=Path(report).with_suffix('.pdf')
 pdf_exists=pdf.is_file() and pdf.read_bytes().startswith(b'%PDF-')
 ck(pdf_exists,'PDF 缺失或格式错误')
 if pdf_exists and any('blocks' in section for section in data.get('sections', [])):
  ck(check_pdf_order(pdf,data['sections']),'PDF 无法读取或小节与图注的顺序不一致')
 return errors
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--report',required=True);p.add_argument('--rules',default='.workflow/template_rules.json');p.add_argument('--data',required=True);p.add_argument('--json-out',required=True);a=p.parse_args()
 warnings=[]
 errors=validate(a.report,a.rules,json.loads(Path(a.data).read_text(encoding='utf-8')),warnings)
 Path(a.json_out).write_text(json.dumps(dict(errors=errors,warnings=warnings,program_pass=not errors,visual_status='待验证',note='程序通过不代表真实性、任务覆盖及视觉通过'),ensure_ascii=False,indent=2),encoding='utf-8')
 raise SystemExit(bool(errors))
