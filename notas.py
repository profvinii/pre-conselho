# -*- coding: utf-8 -*-
import openpyxl, json, re, collections
SRC='/sessions/amazing-affectionate-davinci/mnt/uploads/PROFESSORA MARIA CASIMIRO SOARES - REGISTRO DE NOTAS DOS ESTUDANTES - 2026.xlsx'
wb=openpyxl.load_workbook(SRC, data_only=True, read_only=True)
POS=list(range(38,227,7))   # inicio de cada componente
PER=['1º Período','2º Período','3º Período','4º Período']

def num(v):
    if isinstance(v,(int,float)): return round(float(v),2)
    return None
def txt(v):
    if v is None: return ''
    return re.sub(r'\s+',' ',str(v)).strip()

alunos={}   # (turma,nome) -> dados
comps_by_turma=collections.defaultdict(set)
for pi,pname in enumerate(PER):
    ws=wb[pname]
    rows=[list(r) for r in ws.iter_rows(min_row=1,max_row=2100,max_col=263,values_only=True)]
    cur=None
    for r in rows:
        if len(r)<227: r=r+[None]*(227-len(r))
        if txt(r[4])=='Turma':
            cur=[txt(r[p]) for p in POS]; continue
        if isinstance(r[1],(int,float)) and isinstance(r[2],str) and isinstance(r[4],str) and txt(r[4]) not in ('NOME','Turma'):
            turma=txt(r[2]); nome=txt(r[4])
            if not turma or not nome: continue
            k=(turma,nome)
            a=alunos.setdefault(k,{'turma':turma,'nome':nome,'num':int(r[1]),'periodos':{}})
            a['status_'+str(pi+1)]=txt(r[36]); a['ecd']=txt(r[20])
            notas={}
            for ci,p in enumerate(POS):
                cn=(cur[ci] if cur else '') or ''
                cn=re.sub(r'\s*-\s*$','',cn).strip()
                if not cn or cn.upper() in ('FTP','APF','-'): continue
                av=[num(r[p]),num(r[p+2]),num(r[p+4])]
                rp=[num(r[p+1]),num(r[p+3]),num(r[p+5])]
                med=num(r[p+6])
                if med is None and not any(x is not None for x in av): continue
                comps_by_turma[turma].add(cn)
                notas[cn]={'av':av,'rp':rp,'m':med}
            if notas: a['periodos'][pi+1]=notas
alunos=[v for v in alunos.values() if v['periodos']]
for a in alunos: a['periodos']={str(k):v for k,v in a['periodos'].items()}
turmas=sorted({a['turma'] for a in alunos})
comps=sorted({c for a in alunos for p in a['periodos'].values() for c in p})
out={'turmas':turmas,'componentes':comps,'alunos':alunos,'periodos':PER}
json.dump(out,open('/sessions/amazing-affectionate-davinci/mnt/outputs/notas.json','w',encoding='utf-8'),ensure_ascii=False)
print('alunos:',len(alunos),'turmas:',len(turmas))
print(turmas)
print('componentes:',len(comps)); print(comps)
c=collections.Counter()
for a in alunos:
    for p in a['periodos']: c[p]+=1
print('alunos com notas por periodo:',dict(c))
ex=alunos[0]; print('EX:',ex['nome'],ex['turma'],{k:{kk:vv for kk,vv in list(v.items())[:2]} for k,v in ex['periodos'].items()})
