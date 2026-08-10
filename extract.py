# -*- coding: utf-8 -*-
import openpyxl, json, re
SRC='/sessions/amazing-affectionate-davinci/mnt/uploads/Pré Conselho (respostas).xlsx'
wb=openpyxl.load_workbook(SRC, data_only=True); ws=wb.active
rows=[list(r) for r in ws.iter_rows(min_row=2, values_only=True)]
rows=[r for r in rows if any(v is not None for v in r)]
CAT={
4:["Os professores são acessíveis e ajudam nas dúvidas dos estudantes.",
   "Os estudantes e professores dialogam para o melhor desenvolvimento das aulas.",
   "A turma se esforça para entregar as atividades nos prazos estabelecidos.",
   "A turma coopera com os professores mantendo um ambiente equilibrado para a aprendizagem.",
   "Os professores trabalham de forma dinâmica e criativa."],
5:["A turma não coopera com os professores para manter um ambiente equilibrado para a aprendizagem.",
   "A ausência ou atraso dos professores.","Há muita brincadeiras e conversas paralelas na turma.",
   "Falta de recursos didáticos (vídeos, músicas, jogos, dinâmicas, práticas).",
   "Os atrasos constantes ou faltas dos estudantes.","O uso inadequado de celular ou equipamento eletrônico.",
   "Os professores não são acessíveis e não ajudam nas dúvidas dos estudantes."],
6:["Aulas práticas","Vídeos ou filmes","Debates","Trabalho em pares ou grupo","Jogos (gamificação)",
   "Dinâmicas reflexivas","Desafios (perguntas, enigmas)","Músicas","Palestras","Júri-simulado","Seminários"],
9:["Sinto sono ou tenho dificuldade de concentração.","Tenho vergonha de tirar dúvidas.",
   "Não consigo me organizar para estudar para as avaliações.","Não compreendo os conteúdos.",
   "Não consigo acompanhar a explicação do professor.","Passo por conflitos familiares e/ou emocionais.",
   "Não presto atenção nas aulas.","Não compreendo as tarefas."],
14:["Brincalhona: Gosta de brincadeiras e conversas durante as aulas.","Participativa: Colabora com sugestões de melhoria.",
    "Interessada: Pergunta quando não entende.","Atenciosa: Presta atenção nas aulas.",
    "Esforçada: Empenha-se para aprender, apesar das dificuldades.","Desatenta: Não consegue se concentrar.",
    "Desinteressada: Não tem interesse pelos estudos.","Responsável: Faz as atividades e pesquisas solicitadas pelos professores.",
    "Respeitosa: Zela pelo respeito mútuo.","Dedicada: Estuda em casa, revisa o conteúdo."],
}
SUBJECTS=["Língua Portuguesa","Língua Inglesa","Língua Espanhola","Matemática","Física","Química","Biologia",
          "História","Geografia","Sociologia","Filosofia","Arte","Educação Física"]
def parse_multi(val,catalog):
    if val is None: return []
    low=re.sub(r'\s+',' ',str(val)).strip().lower(); found=[]
    for opt in catalog:
        o=re.sub(r'\s+',' ',opt).strip().lower()
        if o in low: found.append(opt); low=low.replace(o,' | ')
    rest=[p.strip(' ,.;|') for p in re.split(r'\|',low) if len(p.strip(' ,.;|'))>3]
    if rest: found.append("Outro (texto livre)")
    return found
def parse_subjects(val):
    if val is None: return []
    low=re.sub(r'\s+',' ',str(val)).strip().lower(); found=[]
    for opt in sorted(SUBJECTS,key=len,reverse=True):
        if opt.lower() in low: found.append(opt); low=low.replace(opt.lower(),' | ')
    rest=[p.strip(' ,.;|') for p in re.split(r'\|',low) if len(p.strip(' ,.;|'))>3]
    if rest: found.append("Outro (texto livre)")
    return found
OPEN={13:"Sugestões para melhorar a avaliação",15:"Fortalezas da turma",16:"Ameaças da turma",
      17:"O que EU posso fazer",18:"Sugestões para ambiente acolhedor",19:"O que a escola pode fazer",
      20:"Compromissos assumidos",12:"Dificuldades nas avaliações"}
data=[]
for r in rows:
    data.append({"serie":str(r[1]).strip() if r[1] else "Sem série","turma":str(r[2]).strip() if r[2] else "Sem turma",
      "convivencia":str(r[3]).strip() if r[3] else None,"facilitou":parse_multi(r[4],CAT[4]),
      "atrapalhou":parse_multi(r[5],CAT[5]),"metodos":parse_multi(r[6],CAT[6]),
      "maior":parse_subjects(r[7]),"menor":parse_subjects(r[8]),"dificuldades":parse_multi(r[9],CAT[9]),
      "informado":str(r[10]).strip() if r[10] else None,"dif_aval":str(r[11]).strip() if r[11] else None,
      "perfil":parse_multi(r[14],CAT[14]),
      "abertas":{OPEN[i]:(re.sub(r'\s+',' ',str(r[i])).strip() if r[i] else "") for i in OPEN},
      "data": r[0].strftime('%d/%m/%Y') if r[0] else ""})
meta={"catalogs":{"facilitou":CAT[4],"atrapalhou":CAT[5],"metodos":CAT[6],"dificuldades":CAT[9],
      "perfil":CAT[14],"subjects":SUBJECTS},"series":sorted({d["serie"] for d in data}),
      "turmas":sorted({d["turma"] for d in data}),"n":len(data)}
json.dump({"meta":meta,"rows":data},open('dados.json','w',encoding='utf-8'),ensure_ascii=False)
print("respostas",len(data),"turmas",meta["turmas"])
