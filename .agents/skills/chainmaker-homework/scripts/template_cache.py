import argparse,hashlib,json,zipfile
from pathlib import Path
from docx import Document
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def cache(src,out):
 src,out=Path(src),Path(out)
 try:
  r=json.loads(out.read_text(encoding='utf-8'))
  if r['sha256']==sha(src) and r.get('schema')==1 and all(k in r for k in ['parts','paragraphs','tables','requirements']): return r,False
 except (OSError,ValueError,KeyError): pass
 d=Document(src)
 with zipfile.ZipFile(src) as z: parts={n:z.read(n).decode('utf-8') for n in z.namelist() if n.startswith('word/') and n.endswith('.xml')}
 r=dict(schema=1,template=str(src.resolve()),sha256=sha(src),review_status='待核实',requirements={},parts=parts,paragraphs=[p.text for p in d.paragraphs],tables=[[[c.text for c in row.cells] for row in t.rows] for t in d.tables])
 out.parent.mkdir(parents=True,exist_ok=True); out.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8'); return r,True
if __name__=='__main__':
 p=argparse.ArgumentParser()
 p.add_argument('--docx',required=True)
 p.add_argument('--cache',default='.workflow/template_rules.json')
 a=p.parse_args()
 r,reparsed=cache(a.docx,a.cache)
 summary=dict(reparsed=reparsed,review_status=r.get('review_status','待核实'),
              requirements=r['requirements'])
 print(json.dumps(summary,ensure_ascii=False))
