# -*- coding: utf-8 -*-
import openpyxl, json, re, collections
SRC='/sessions/amazing-affectionate-davinci/mnt/uploads/PROFESSORA MARIA CASIMIRO SOARES - REGISTRO DE NOTAS DOS ESTUDANTES - 2026.xlsx'
wb=openpyxl.load_workbook(SRC, data_only=True, read_only=True)
POS=list(range(38,227,7))
PER=['1º Período','2º Período','3º Período','4º Período']
def num(v): return round(float(v),2) if isinstance(v,(int,float)) else None
def txt(v): return re.sub(r'\s+',' ',str(v)).strip() if v is not None else ''
def label(code):
    m=re.match(r'.*?-([A-Z]+)-(\d+)$',code)
    if not m: return code
    tp,nn=m.group(1),m.group(2)
    return nn if tp=='EPT' else nn+' '+tp
def serie(code):
    m=re.search(r'-(\d)\d\d$',code)
    return {'1':'1º Ano','2':'2º Ano','3':'3º Ano'}.get(m.group(1),'—') if m else '—'

recs={}; compset=[]
def cidx(c):
    if c not in compset: compset.append(c)
    return compset.index(c)
for pi,pname in enumerate(PER,1):
    ws=wb[pname]; cur=None
    for r in ws.iter_rows(min_row=1,max_row=2100,max_col=263,values_only=True):
        r=list(r)+[None]*(263-len(r))
        if txt(r[4])=='Turma': cur=[txt(r[p]) for p in POS]; continue
        if isinstance(r[1],(int,float)) and isinstance(r[2],str) and isinstance(r[4],str) and txt(r[4]) not in ('NOME','Turma'):
            turma,nome=txt(r[2]),txt(r[4])
            if not turma or not nome: continue
            a=recs.setdefault((turma,nome),{'t':label(turma),'s':serie(turma),'n':nome,'st':{},'p':{}})
            if txt(r[36]): a['st'][str(pi)]=txt(r[36])
            d={}
            for ci,p in enumerate(POS):
                cn=re.sub(r'\s*-\s*$','',(cur[ci] if cur else '') or '').strip()
                if not cn or cn.upper() in ('FTP','APF','-'): continue
                vals=[num(r[p]),num(r[p+2]),num(r[p+4]),num(r[p+1]),num(r[p+3]),num(r[p+5]),num(r[p+6])]
                if all(v is None for v in vals): continue
                d[str(cidx(cn))]=vals
            if d: a['p'][str(pi)]=d
alunos=[v for v in recs.values() if v['p']]
alunos.sort(key=lambda a:(a['s'],a['t'],a['n']))
turmas=sorted({a['t'] for a in alunos}, key=lambda t:(t[0],t))
out={'comp':compset,'turmas':turmas,'series':sorted({a['s'] for a in alunos}),
     'per':PER,'perComNotas':sorted({k for a in alunos for k in a['p']}),'alunos':alunos}
json.dump(out,open('notas.json','w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
import os
print('alunos',len(alunos),'turmas',turmas)
print('comps',compset)
print('periodos com notas',out['perComNotas'])
print('json KB %.0f'%(os.path.getsize('notas.json')/1024))
# checagem: média geral de um aluno
a=alunos[0]
for p,d in a['p'].items():
    ms=[v[6] for v in d.values() if v[6] is not None]
    print(a['n'],a['t'],'P'+p,'comps',len(d),'media %.2f'%(sum(ms)/len(ms)))
