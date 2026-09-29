/* =====================================================================
   Documentos da contratação — cadeia ETP → Anexo I (memória .xlsx) → TR → Proposta
   Bibliotecas carregadas sob demanda (jsDelivr): ExcelJS, docx, JSZip.
   ===================================================================== */
const XLSX_STATIC = __XLSX_STATIC__;
const LIBS = {
  xlsx: ['https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js', () => window.ExcelJS],
  docx: ['https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js', () => window.docx],
  zip:  ['https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js', () => window.JSZip]};
const _libP = {};
function carregar(k){
  if(LIBS[k][1]()) return Promise.resolve();
  return _libP[k] || (_libP[k] = new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = LIBS[k][0]; s.async = true;
    s.onload = () => LIBS[k][1]() ? res() : rej(new Error('Biblioteca '+k+' não inicializou'));
    s.onerror = () => { delete _libP[k]; rej(new Error('Sem acesso à biblioteca '+k+' (verifique a conexão com a internet)')); };
    document.head.appendChild(s); }));
}
function salvarBlob(blob, nome){
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

/* ---------- escolhas do usuário ---------- */
const MODOS = {
  srp:     {t:'Pregão eletrônico para Registro de Preços', s:'Ata de 1 ano e contratação por demanda (arts. 82 a 86 da Lei nº 14.133/2021).'},
  pregao:  {t:'Pregão eletrônico — contrato de serviço continuado', s:'Contrato de 12 meses, prorrogável até 5 anos (arts. 106 e 107).'},
  srpcons: {t:'Registro de Preços compartilhado por consórcio público', s:'Consórcio como órgão gerenciador e municípios como participantes (art. 86; Lei nº 11.107/2005).'},
  adesao:  {t:'Adesão a Ata de Registro de Preços vigente', s:'Uso de ata de outro órgão, com demonstração de vantagem (art. 86, §§ 2º a 5º).'},
  direta:  {t:'Contratação direta — piloto', s:'Hipótese legal a ser definida pela assessoria jurídica (arts. 72 a 75).'}};
const DEMS = {edu:'Educação', sau:'Saúde', conj:'Educação e Saúde (demanda conjunta)'};
S.doc = {dem:null, modo:null, lotes:'rubrica', g3:false, proc:'', pca:'', uo:''};
const CANAL_MODO = {direta:'direta', licit:'srp', cons:'srpcons', pmi:'srp'};
function docModo(){ return S.doc.modo || (S.tipo==='cons' ? 'srpcons' : CANAL_MODO[S.canal] || 'srp'); }
function docDem(){ return S.doc.dem || ((LAST && LAST.inc.some(l=>l.cam==='C')) ? 'conj' : 'edu'); }

function esferaTxt(){ return S.tipo==='cons' ? 'Intermunicipal — consórcio público' : S.tipo==='uf' ? (S.uf==='DF'?'Distrital':'Estadual') : 'Municipal'; }
function orgaoDemandante(){
  const dem = docDem();
  if(S.tipo==='mun'){ const m=M[S.mun]; const sec = dem==='edu'?'Secretaria Municipal de Educação':dem==='sau'?'Secretaria Municipal de Saúde':'Secretarias Municipais de Educação e de Saúde';
    return 'Prefeitura Municipal de '+m[2]+'/'+m[1]+' — '+sec; }
  if(S.tipo==='cons') return (S.consNome.trim()||'Consórcio público intermunicipal')+' — órgão gerenciador, em nome dos municípios consorciados';
  const sec = dem==='edu'?'Secretaria de Estado da Educação':dem==='sau'?'Secretaria de Estado da Saúde':'Secretarias de Estado da Educação e da Saúde';
  if(S.uf==='DF') return 'Governo do Distrito Federal — '+sec.replace('de Estado da','de Estado de');
  const u=UFN[S.uf]; return 'Governo do Estado '+u[1]+' '+u[0]+' — '+sec;
}
function redeTxt(){ return S.tipo==='uf' ? 'rede estadual de ensino' : S.tipo==='cons' ? 'redes municipais de ensino dos municípios consorciados' : 'rede municipal de ensino'; }
const MESES=['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
function dataExt(){ const t=new Date(); return MESES[t.getMonth()]+' de '+t.getFullYear(); }
const PRE = t => '[PREENCHER'+(t?' — '+t:'')+']';

/* ---------- itens para documentos ---------- */
const DESC = {
  esc:'Licença de uso — Módulo Escola (governança pedagógica e gestão de PEI)',
  edu:'Licença de uso — por educando com PEI ativo',
  sau:'Licença de uso — Módulo Saúde',
  fam:'Licença de uso — Módulo Família',
  pai:'Painel de governança e inteligência para a Secretaria',
  imp:'Implantação e parametrização por unidade escolar',
  int:'Migração de dados e integração com sistema legado',
  for:'Formação continuada — turma de até 20 participantes, 20 h',
  sup:'Supervisão e mentoria de casos por equipe multiprofissional',
  sus:'Suporte técnico, sustentação e evolução',
  ava:'Licença de plataforma de avaliação pedagógica de habilidades (gamificada), sem finalidade diagnóstica',
  lab:'Locação de laboratório itinerante de avaliação pedagógica'};
const MEDICAO = {
  esc:'Mensal (1/12 do valor anual), proporcional às licenças ativas; ativações no curso do mês pro rata die.', edu:'idem', sau:'idem', fam:'idem',
  pai:'Mensal, correspondente a 1/12 do valor anual.',
  imp:'Por unidade escolar, após o termo de aceite da respectiva unidade.',
  int:'Por ponto de integração, após validação do relatório de integração pelo fiscal técnico.',
  for:'Por turma concluída, com lista de presença, avaliações e certificados; frequência inferior a 75% é medida proporcionalmente.',
  sup:'Por hora técnica efetivamente prestada, mediante relatório de atendimento validado.',
  sus:'Mensal, condicionada ao relatório de serviços e à apuração dos níveis de serviço.',
  ava:'Mensal (1/12 do valor anual), proporcional aos educandos com avaliação habilitada.',
  lab:'Por serviço realizado, mediante relatório de aplicação validado.',
  cli:'Por atendimento realizado, comprovado por registro no módulo de saúde e confirmação do responsável legal.'};
function memo(l, d){
  const aj = l.q!==l.qa ? ' Quantidade ajustada manualmente para '+num(l.q)+' (cálculo-padrão: '+num(l.qa)+').' : '';
  const nEq = Math.ceil(d.acomp/P.alunosEquipe);
  const m = {
    esc:int(d.escolas)+' unidades escolares — estimativa de '+int(d.acomp)+' educandos ÷ '+num(P.alunosEsc,1)+' por unidade. Confirmar no Censo Escolar.',
    edu:int(d.acomp)+' educandos = '+int(d.recorte)+' no recorte etário × '+num(P.rede,1)+'% em rede pública × fator '+num(P.aee,2)+' (público do AEE) × '+num(P.adesao,1)+'% de adesão.',
    sau:int(d.profs)+' profissionais — '+num(P.profEsc,1)+' por unidade escolar. Confirmar com a Saúde (AEE, CER, RAPS, eMulti).',
    fam:num(P.familia,1)+'% dos educandos acompanhados (fator de irmandade).',
    pai:'1 licença por órgão demandante'+(d.entes>1?' ('+d.entes+' entes)':'')+'.',
    imp:'Igual ao nº de unidades escolares, executada de forma faseada no 1º ano.',
    int:num(P.pontos)+' pontos por ente (gestão escolar, Educacenso/INEP, sistemas de saúde).',
    for:int(d.escolas*P.partEsc)+' profissionais a capacitar ('+num(P.partEsc)+' por unidade) ÷ 20 por turma.',
    sup:num(P.horasEsc)+' h/ano por unidade escolar (≈ '+num(d.escolas*P.horasEsc/12,0)+' h/mês).',
    sus:'12 meses × '+nEq+' equipe'+(nEq>1?'s':'')+' de sustentação (1 a cada '+int(P.alunosEquipe)+' educandos).',
    ava:'Igual ao nº de educandos acompanhados.',
    lab:'1 laboratório para cada '+num(P.escLab)+' unidades escolares.'};
  return (m[l.id]||'')+aj;
}
function itensDoc(C){
  const out=[];
  C.inc.forEach(l=>{
    if(l.id==='cli'){
      const fat = pacoteTotal() ? l.pu/pacoteTotal() : 1;
      PAC.forEach(x=>{ if(!(x.n>0 && x.v>0)) return; const pu=r2(x.v*fat), q=l.q*x.n;
        out.push({id:'cli', cam:'C', rec:true, tipo:'srv', desc:(/Neuro|Psiq/.test(x.s)?'Consulta em ':'Sessão de ')+x.s.toLowerCase().replace('neurologista','neurologia').replace('psiquiatra','psiquiatria')+' ao educando do público-alvo',
          un:/Neuro|Psiq/.test(x.s)?'Consulta':'Sessão', q, pu, tot:r2(q*pu), memo:int(l.q)+' educandos com pacote clínico × '+x.n+' atendimentos/ano ('+x.o.toLowerCase()+').'}); });
    } else out.push({id:l.id, cam:l.cam, rec:l.rec, tipo:l.tipo, desc:DESC[l.id], un:l.un, q:l.q, pu:l.pu, tot:l.tot, memo:memo(l, C.d)});
  });
  out.forEach((x,i)=>x.n=i+1);
  return out;
}
function montarLotes(C){
  const it = itensDoc(C), vig=S.vig;
  const mk = (nome, itens, o) => { const ano=r2(itens.reduce((a,x)=>a+x.tot,0)), nr=r2(itens.filter(x=>!x.rec).reduce((a,x)=>a+x.tot,0));
    return Object.assign({nome, itens, ano, nr, vig:r2(nr+(ano-nr)*vig/12)}, o); };
  const TIC = '3.3.90.40 — Serviços de Tecnologia da Informação e Comunicação — Pessoa Jurídica';
  const A=it.filter(x=>x.cam==='A'), B=it.filter(x=>x.cam==='B'), Cc=it.filter(x=>x.cam==='C');
  const L=[];
  if(S.doc.lotes==='unico' || [A,B,Cc].filter(a=>a.length).length<2){
    L.push(mk('Solução integrada de inclusão escolar e coordenação do cuidado', it, {rub:Cc.length?(A.length||B.length?'Educação e Saúde':'Saúde'):'Educação', nat:Cc.length&&(A.length||B.length)?TIC+'; 3.3.90.39 — Outros Serviços de Terceiros — PJ (parcela clínica)':Cc.length?'3.3.90.39 — Outros Serviços de Terceiros — Pessoa Jurídica':TIC,
      natServ:'Serviço comum, de natureza continuada, sem dedicação exclusiva de mão de obra', dem:Cc.length&&!(A.length||B.length)?'Saúde':'Educação'}));
  } else {
    if(A.length) L.push(mk('Solução tecnológica de gestão do PEI e coordenação intersetorial, com serviços habilitadores', A, {rub:'Educação', nat:TIC, natServ:'Serviço comum de tecnologia da informação, de natureza continuada, sem dedicação exclusiva de mão de obra', dem:'Educação'}));
    if(B.length) L.push(mk('Avaliação pedagógica de habilidades, com laboratórios itinerantes', B, {rub:'Educação', nat:TIC, natServ:'Serviço comum, de natureza continuada, sem dedicação exclusiva de mão de obra, sem finalidade diagnóstica', dem:'Educação'}));
    if(Cc.length) L.push(mk('Atendimento multiprofissional em saúde ao público-alvo', Cc, {rub:'Saúde', nat:'3.3.90.39 — Outros Serviços de Terceiros — Pessoa Jurídica', natServ:'Serviço de saúde de natureza continuada, prestado por profissionais legalmente habilitados', dem:'Saúde'}));
  }
  L.forEach((x,i)=>x.n=i+1);
  return L;
}
function participantes(){
  if(S.tipo!=='cons') return [];
  return S.cons.map(i=>{ const d=dims([M[i][3],M[i][4],M[i][5],M[i][6]],1); const t=somar(linhas(d,false).filter(incluido));
    return {nome:M[i][2]+'/'+M[i][1], cod:M[i][0], total:d.total, recorte:d.recorte, acomp:d.acomp, clin:d.clin, ano:t.ano}; });
}
function objetoPartes(L){
  const cams = new Set(L.flatMap(x=>x.itens.map(i=>i.cam)));
  const p=[];
  if(cams.has('A')) p.push('plataforma tecnológica de gestão, coordenação e monitoramento do Plano Educacional Individualizado (PEI) e do Atendimento Educacional Especializado (AEE), com inteligência artificial e articulação intersetorial entre educação, saúde e família, com licenciamento em nuvem, implantação, integração com sistemas legados, formação continuada, supervisão técnica de casos e suporte');
  if(cams.has('B')) p.push('avaliação pedagógica de habilidades, com laboratórios itinerantes, sem finalidade diagnóstica');
  if(cams.has('C')) p.push('atendimento multiprofissional em saúde (neurologia, psiquiatria, fonoaudiologia, psicologia e psicopedagogia) ao público-alvo');
  return p;
}
function ctxDoc(){
  const C = calc(); LAST = C;
  const L = montarLotes(C);
  const pr = PRESETS.find(p=>p.id===S.preset);
  return {C, d:C.d, L, label:enteLabel(), orgao:orgaoDemandante(), esfera:esferaTxt(), modo:docModo(), dem:docDem(), vig:S.vig,
    ref:refCode(), data:dataExt(), recorteTxt: S.preset==='custom'?'recorte etário personalizado':pr.t, part:participantes(),
    total:r2(L.reduce((a,x)=>a+x.vig,0)), ano:r2(L.reduce((a,x)=>a+x.ano,0)), nr:r2(L.reduce((a,x)=>a+x.nr,0)),
    cams:new Set(C.inc.map(l=>l.cam)), partes:objetoPartes(L), slug:slug(enteLabel())};
}

/* =====================================================================
   ANEXO I — MEMÓRIA DE CÁLCULO (.xlsx, fórmulas vivas)
   ===================================================================== */
async function gerarXlsx(ctx){
  await carregar('xlsx');
  const E = window.ExcelJS, {C, d} = ctx;
  const wb = new E.Workbook(); wb.creator='G3 Health Service'; wb.created=new Date(); wb.calcProperties.fullCalcOnLoad = true;
  const F='Times New Roman';
  const FM = {int:'#,##0;\\(#,##0\\);\\-', money:'#,##0.00;\\(#,##0.00\\);\\-', pct:'0.0%', fat:'0.00', dec:'#,##0.0;\\(#,##0.0\\);\\-'};
  const thin = {style:'thin', color:{argb:'FFBFBFBF'}};
  const BORD = {top:thin,left:thin,bottom:thin,right:thin};
  const fill = a => ({type:'pattern', pattern:'solid', fgColor:{argb:a}});
  const put = (ws, a, v, o={}) => { const c=ws.getCell(a); c.value=v;
    c.font = {name:F, size:o.size||11, bold:!!o.bold, color:{argb:o.inp?'FF0000FF':'FF000000'}, italic:!!o.it};
    if(o.inp) c.fill=fill(o.ov?'FFFCE4D6':'FFFFFF00'); else if(o.hdr) c.fill=fill('FFD9D9D9'); else if(o.tot) c.fill=fill('FFE8E8E8');
    if(o.fmt) c.numFmt=FM[o.fmt]||o.fmt; if(o.wrap||o.hdr) c.alignment={wrapText:true, vertical:'top'};
    if(o.b) c.border=BORD; if(o.note) c.note=o.note; return c; };
  const fx = (formula, result) => ({formula, result});
  const title = (ws, a, t) => put(ws, a, t, {bold:true, size:12});
  const hdr = (ws, r, arr) => arr.forEach((t,i)=>put(ws, String.fromCharCode(65+i)+r, t, {hdr:true, bold:true, size:12, b:true}));
  const P1="'1. Parametros'!", P2="'2. Dimensionamento'!", P3="'3. Plataforma'!", P4="'4. Cuidado clinico'!";
  const L = C.ls; const byId = id => L.find(l=>l.id===id);
  const inc = id => incluido(byId(id)) ? 1 : 0;
  const idx = selIdx(), nM = idx.length, mR0 = 5, mR1 = 4+nM;
  const b = d.b;

  /* ---- 1. Parametros ---- */
  const w1 = wb.addWorksheet('1. Parametros', {properties:{tabColor:{argb:'FF14225A'}}});
  w1.columns = [{width:46},{width:22},{width:62},{width:16}];
  title(w1,'A1','MODELO DE VIABILIDADE — PLATAFORMA DE INCLUSÃO ESCOLAR E CUIDADO CLÍNICO');
  put(w1,'A2','Ente: '+ctx.label+' — esfera '+ctx.esfera.toLowerCase()+(S.tipo!=='mun'?' ('+nM+' municípios, ver aba 8)':' (IBGE '+M[S.mun][0]+')')+'. Anexo I do ETP — Ref. '+ctx.ref+'.');
  put(w1,'A3','LEGENDA: células AMARELAS com texto AZUL são os parâmetros editáveis. Células pretas são fórmulas — não editar. Células LARANJA foram ajustadas manualmente no painel.');
  put(w1,'A4','Elaboração: gerado pelo painel Inclusi.Via (G3 Health Service) em '+new Date().toLocaleDateString('pt-BR')+'. Data-base dos preços: setembro de 2026.');
  title(w1,'A6','A. POPULAÇÃO-ALVO — CENSO 2022 (dado oficial, não editar)');
  hdr(w1,7,['Faixa etária','Pessoas com TEA — '+ctx.label,'Fonte','% no recorte adotado']);
  FAIXAS.forEach((f,k)=>{ const r=8+k, col=String.fromCharCode(68+k);
    put(w1,'A'+r,f.k); put(w1,'B'+r,fx(`SUM('8. Municipios'!${col}${mR0}:${col}${mR1})`, b[k]),{fmt:'int'});
    put(w1,'C'+r,'IBGE, SIDRA 10145, var. 13267 — Censo 2022'); put(w1,'D'+r,S.f[k],{inp:true,fmt:'0%'}); });
  put(w1,'A12','Total 0 a 19 anos',{tot:true,bold:true}); put(w1,'B12',fx('SUM(B8:B11)',d.total),{tot:true,bold:true,fmt:'int'});
  title(w1,'A14','B. RECORTES ETÁRIOS ESTIMADOS (distribuição uniforme dentro da faixa)');
  hdr(w1,15,['Recorte','Fórmula aplicada','Pessoas com TEA']);
  put(w1,'A16','4 a 18 anos (idade escolar ampla)'); put(w1,'B16','1/5 da faixa 0-4 + 5-9 + 10-14 + 4/5 da faixa 15-19',{wrap:true}); put(w1,'C16',fx('ROUND(B8/5+B9+B10+B11*4/5,0)',Math.round(b[0]/5+b[1]+b[2]+b[3]*.8)),{fmt:'int'});
  put(w1,'A17','6 a 14 anos (Ensino Fundamental I e II)'); put(w1,'B17','4/5 da faixa 5-9 + 10-14'); put(w1,'C17',fx('ROUND(B9*4/5+B10,0)',Math.round(b[1]*.8+b[2])),{fmt:'int'});
  put(w1,'A18','RECORTE ADOTADO NESTA MEMÓRIA — '+ctx.recorteTxt,{bold:true,tot:true}); put(w1,'B18','Soma das faixas × % da coluna D',{tot:true}); put(w1,'C18',fx('ROUND(SUMPRODUCT(B8:B11,D8:D11),0)',d.recorte),{fmt:'int',bold:true,tot:true});
  title(w1,'A20','C. PARÂMETROS DE COBERTURA (editáveis)');
  hdr(w1,21,['Parâmetro','Valor','Observação']);
  const pc = [['% do público em rede pública '+(S.tipo==='uf'?'estadual':'municipal'),P.rede/100,'0.0%','A confirmar no Censo Escolar. Redes municipais costumam concentrar a educação infantil e o EF I; o EF II e o ensino médio são frequentemente estaduais.'],
    ['Fator de ampliação para demais públicos do AEE',P.aee,'fat','TEA é apenas parte do público-alvo da educação especial. Fator 1,0 = só TEA. Calibrar pelo Censo Escolar local.'],
    ['Taxa de adesão efetiva ao programa',P.adesao/100,'0.0%','Nem todo elegível adere. Considera recusa familiar e perda de seguimento.'],
    ['% do público que recebe o pacote clínico',P.clin/100,'0.0%','Parâmetro de política. Nem todo aluno com PEI demanda acompanhamento clínico intensivo.'],
    ['Fator de famílias no Módulo Família',P.familia/100,'0.0%','Fator de irmandade (educandos da mesma família).']];
  pc.forEach((x,i)=>{ const r=22+i; put(w1,'A'+r,x[0]); put(w1,'B'+r,x[1],{inp:true,fmt:x[2]}); put(w1,'C'+r,x[3],{wrap:true}); });
  title(w1,'A28','D. ESTRUTURA DA REDE (editáveis — confirmar no Censo Escolar)');
  hdr(w1,29,['Parâmetro','Valor','Observação']);
  const nEq = d.acomp? Math.ceil(d.acomp/P.alunosEquipe):0;
  const est = [['Unidades escolares com educandos atendidos',d.escolas,'int','A CONFIRMAR. Estimativa: '+int(d.acomp)+' educandos ÷ '+num(P.alunosEsc,1)+' por unidade (calibração Guarujá).'],
    ['Profissionais da educação a capacitar',d.escolas*P.partEsc,'int','A CONFIRMAR. '+num(P.partEsc)+' por unidade: regentes, AEE, coordenação e gestão.'],
    ['Profissionais de saúde com acesso ao módulo',d.profs,'int','A CONFIRMAR. Rede municipal e conveniada que atende o público-alvo.'],
    ['Participantes por turma de formação',20,'int','Limite pedagógico usual e adotado na ARP 045/2026 do CIMINAS.'],
    ['Horas de supervisão de casos por mês',d.escolas*P.horasEsc/12,'dec','Equipe multiprofissional. Dimensionar pelo número de casos complexos.'],
    ['Pontos de integração com sistemas legados',d.acomp?d.entes*P.pontos:0,'int','Gestão escolar, Educacenso/INEP, sistemas de saúde — por ente.'],
    ['Equipes-mês de suporte por ano',12*nEq,'int','12 meses × '+nEq+' equipe(s); 1 equipe a cada '+int(P.alunosEquipe)+' educandos.'],
    ['Unidades escolares por laboratório itinerante',P.escLab,'int','ARP 045/2026: um laboratório para cada 4 unidades.'],
    ['Órgãos com painel de governança',d.acomp?d.entes:0,'int','Uma licença por Secretaria/ente.']];
  est.forEach((x,i)=>{ const r=30+i; put(w1,'A'+r,x[0]); put(w1,'B'+r,x[1],{inp:true,fmt:x[2]}); put(w1,'C'+r,x[3],{wrap:true}); });
  title(w1,'A40','E. FONTES E ÂNCORAS DE PREÇO PÚBLICO');
  hdr(w1,41,['Referência','Valor','Origem']);
  [['Licença de plataforma de avaliação — aluno/ano',750.97,'ARP 045/2026, CIMINAS, Item 1. Concorrência Eletrônica 004/2026.'],['Implantação por unidade educacional',1162.49,'ARP 045/2026, CIMINAS, Item 2.'],['Formação — turma de até 20 profissionais',3520.84,'ARP 045/2026, CIMINAS, Item 3.'],['Locação de laboratório itinerante',2542.89,'ARP 045/2026, CIMINAS, Item 4.'],['Solução de AEE consorciada — valor global',116389239.08,'CIMPAR, Pregão Eletrônico 07/2026, homologado 16/07/2026. Vencedora: Cognvox.'],['Sistema de gestão escolar municipal — rede inteira/ano',360000,'Chorozinho/CE, processo 001-005/2026-AD-SEDUC, 12 meses.'],['Licença privada de gestão de PEI — aluno/ano (inferido)',189,'IncludED, plano intermediário no teto de 20 alunos. Cálculo aritmético, não preço publicado.']]
    .forEach((x,i)=>{ const r=42+i; put(w1,'A'+r,x[0]); put(w1,'B'+r,x[1],{fmt:'money'}); put(w1,'C'+r,x[2],{wrap:true}); });
  title(w1,'A50','F. CONDIÇÕES DA PROPOSTA (editáveis)');
  hdr(w1,51,['Parâmetro','Valor','Observação']);
  put(w1,'A52','Vigência do contrato (meses)'); put(w1,'B52',S.vig,{inp:true,fmt:'int'}); put(w1,'C52','Serviço contínuo: até 5 anos (arts. 106 e 107 da Lei nº 14.133/2021).');
  put(w1,'A53','Escopo adotado'); put(w1,'B53',S.cen==='P'?'Personalizado':'Cenário '+S.cen); put(w1,'C53','Itens da proposta marcados com 1 na coluna G da aba 3 e na célula F16 da aba 4.');
  put(w1,'A54','Tipo de contratação'); put(w1,'B54',MODOS[ctx.modo].t,{wrap:true}); put(w1,'C54','Órgão demandante: '+ctx.orgao,{wrap:true});

  /* ---- 2. Dimensionamento ---- */
  const w2 = wb.addWorksheet('2. Dimensionamento', {properties:{tabColor:{argb:'FF3A4FA8'}}});
  w2.columns=[{width:52},{width:16},{width:58}];
  title(w2,'A1','DIMENSIONAMENTO DO PÚBLICO — '+ctx.label.toUpperCase());
  put(w2,'A2','Todos os valores derivam da aba 1. Alterar os parâmetros lá, não aqui.');
  hdr(w2,4,['Etapa do dimensionamento','Alunos','Cálculo']);
  [['Pessoas com TEA no recorte adotado (Censo 2022)',P1+'C18',d.recorte,'Base populacional oficial — '+ctx.recorteTxt],
   ['Em rede pública',`ROUND(B5*${P1}$B$22,0)`,d.rede,'Aplicado o percentual de rede'],
   ['Ampliado para o público-alvo do AEE',`ROUND(B6*${P1}$B$23,0)`,d.aee,'Inclui demais deficiências, TGD e altas habilidades'],
   ['Com adesão efetiva ao programa',`ROUND(B7*${P1}$B$24,0)`,d.acomp,'PÚBLICO ATENDIDO PELA PLATAFORMA'],
   ['Público que recebe o pacote clínico',`ROUND(B8*${P1}$B$25,0)`,d.clin,'Aplica o percentual clínico da aba 1']]
   .forEach((x,i)=>{ const r=5+i, key=(r===8); put(w2,'A'+r,x[0],{bold:key,tot:key}); put(w2,'B'+r,fx(x[1],x[2]),{fmt:'int',bold:key,tot:key}); put(w2,'C'+r,x[3],{bold:key,tot:key}); });
  title(w2,'A11','DISTRIBUIÇÃO POR FAIXA ETÁRIA');
  hdr(w2,12,['Faixa etária','No recorte','Acompanhados (proporcional)']);
  const sr = d.rec.reduce((a,c)=>a+c,0)||1;
  FAIXAS.forEach((f,k)=>{ const r=13+k; put(w2,'A'+r,f.k+' — '+f.etapa); put(w2,'B'+r,fx(`${P1}B${8+k}*${P1}D${8+k}`,d.b[k]*S.f[k]),{fmt:'int'});
    put(w2,'C'+r,fx(`IFERROR(ROUND($B$8*B${r}/SUM($B$13:$B$16),0),0)`,Math.round(d.acomp*d.rec[k]/sr)),{fmt:'int'}); });
  title(w2,'A18','RECORTES COMPLEMENTARES');
  hdr(w2,19,['Recorte','Alunos','Observação']);
  put(w2,'A20','Público de Ensino Fundamental I e II (6 a 14 anos)'); put(w2,'B20',fx(`ROUND(${P1}C17*${P1}$B$22*${P1}$B$23*${P1}$B$24,0)`,Math.round(Math.round(b[1]*.8+b[2])*P.rede/100*P.aee*P.adesao/100)),{fmt:'int'}); put(w2,'C20','Escopo da ARP 045/2026');
  title(w2,'A22','CONTEXTO — CENSO 2022 (referência, não editar)');
  hdr(w2,23,['Recorte','Pessoas com TEA','Observação']);
  const uf0 = M[idx[0]] ? M[idx[0]][1] : S.uf; let ufT=0, ufN=0; M.forEach(m=>{ if(m[1]===uf0){ ufT+=m[3]+m[4]+m[5]+m[6]; ufN++; } });
  put(w2,'A24',ctx.label+' — 0 a 19 anos'); put(w2,'B24',fx(P1+'B12',d.total),{fmt:'int'}); put(w2,'C24','Ente desta memória');
  put(w2,'A25',(uf0==='DF'?'Distrito Federal':'Estado '+UFN[uf0][1]+' '+UFN[uf0][0])+' — 0 a 19 anos'); put(w2,'B25',ufT,{fmt:'int'}); put(w2,'C25',int(ufN)+' municípios');
  put(w2,'A26','Brasil — 0 a 19 anos'); put(w2,'B26',1060019,{fmt:'int'}); put(w2,'C26','5.570 municípios');

  /* ---- 3. Plataforma ---- */
  const w3 = wb.addWorksheet('3. Plataforma', {properties:{tabColor:{argb:'FFE0651A'}}});
  w3.columns=[{width:44},{width:18},{width:14},{width:17},{width:18},{width:44},{width:13}];
  title(w3,'A1','CUSTO ANUAL DA SOLUÇÃO TECNOLÓGICA E DOS SERVIÇOS HABILITADORES');
  put(w3,'A2','Preços unitários em AMARELO são propostos e devem ser validados por pesquisa de preços (art. 23 da Lei 14.133/2021). Coluna G: 1 = item integra a proposta.');
  const qF = {esc:P1+'$B$30', edu:P2+'$B$8', sau:P1+'$B$32', fam:`ROUND(${P2}$B$8*${P1}$B$26,0)`, pai:P1+'$B$38', imp:P1+'$B$30', int:P1+'$B$35',
    for:`ROUNDUP(${P1}$B$31/${P1}$B$33,0)`, sup:`ROUND(${P1}$B$34*12,0)`, sus:P1+'$B$36', ava:P2+'$B$8', lab:`ROUNDUP(${P1}$B$30/${P1}$B$37,0)`};
  const ANC = {esc:'Proposto. Sem referência pública direta; calibrar por pesquisa de preços.', edu:'Entre R$ 189 (privado, inferido) e R$ 750,97 (ARP 045/2026).', sau:'Proposto. Materializa a Lei 14.254/2021.', fam:'Proposto. Pode ser ofertado incluso.', pai:'Proposto.',
    imp:'ARP 045/2026, CIMINAS, Item 2 — referência pública direta.', int:'Proposto. Educacenso/INEP, gestão escolar, sistemas de saúde.', for:'ARP 045/2026, CIMINAS, Item 3 — referência pública direta.', sup:'Proposto. Equipe com neuropediatra, TO e psicopedagogo.', sus:'Proposto.',
    ava:'ARP 045/2026, CIMINAS, Item 1 — referência pública direta.', lab:'ARP 045/2026, CIMINAS, Item 4. Um laboratório para cada 4 unidades.'};
  const itemRow = (ws, r, id) => { const l=byId(id), ov=S.ovQ[id]!=null;
    put(ws,'A'+r,DESC[id].replace('Licença de uso','Licença'),{wrap:true}); put(ws,'B'+r,l.un);
    put(ws,'C'+r, ov ? l.q : fx(qF[id], l.qa), ov ? {inp:true,ov:true,fmt:'int',note:'Ajustado manualmente no painel. Cálculo-padrão: '+num(l.qa)} : {fmt:'int'});
    put(ws,'D'+r,l.pu,{inp:true,fmt:'money'}); put(ws,'E'+r,fx(`C${r}*D${r}`,l.tot),{fmt:'money'}); put(ws,'F'+r,ANC[id],{wrap:true}); put(ws,'G'+r,inc(id),{inp:true,fmt:'0'}); };
  hdr(w3,4,['Item','Unidade','Quantidade','Preço unitário (R$)','Total anual (R$)','Âncora de preço','Na proposta (1=sim)']);
  ['esc','edu','sau','fam','pai','imp','int','for','sup','sus'].forEach((id,i)=>itemRow(w3,5+i,id));
  put(w3,'A15','TOTAL ANUAL — PLATAFORMA E SERVIÇOS',{tot:true,bold:true}); put(w3,'E15',fx('SUM(E5:E14)',C.tA.ano),{tot:true,bold:true,fmt:'money'});
  put(w3,'A16','Custo por aluno atendido (R$/ano)'); put(w3,'E16',fx(`IFERROR(E15/${P2}$B$8,0)`,d.acomp?C.tA.ano/d.acomp:0),{fmt:'money'});
  title(w3,'A18','MÓDULO OPCIONAL — AVALIAÇÃO PEDAGÓGICA GAMIFICADA (escopo da ARP 045/2026)');
  hdr(w3,19,['Item','Unidade','Quantidade','Preço unitário (R$)','Total anual (R$)','Âncora de preço','Na proposta (1=sim)']);
  itemRow(w3,20,'ava'); itemRow(w3,21,'lab');
  put(w3,'A22','TOTAL ANUAL — MÓDULO DE AVALIAÇÃO',{tot:true,bold:true}); put(w3,'E22',fx('SUM(E20:E21)',C.tB.ano),{tot:true,bold:true,fmt:'money'});

  /* ---- 4. Cuidado clínico ---- */
  const w4 = wb.addWorksheet('4. Cuidado clinico', {properties:{tabColor:{argb:'FF0E9488'}}});
  w4.columns=[{width:34},{width:20},{width:18},{width:20},{width:58},{width:14}];
  title(w4,'A1','CUSTO ANUAL DO CUIDADO CLÍNICO');
  put(w4,'A2','Valores unitários conforme tabela do contratante. Recomenda-se estratificar por nível de suporte do TEA (1, 2 e 3).');
  title(w4,'A4','A. CUSTO POR ALUNO/ANO — PACOTE DE ATENDIMENTO PROFISSIONAL');
  hdr(w4,5,['Serviço','Valor por consulta (R$)','Quantidade por ano','Total por aluno/ano (R$)','Observação técnica']);
  PAC.forEach((x,i)=>{ const r=6+i; put(w4,'A'+r,/Neuro|Psiq/.test(x.s)?'Consulta com '+x.s.toLowerCase():x.s); put(w4,'B'+r,x.v,{inp:true,fmt:'money'}); put(w4,'C'+r,x.n,{inp:true,fmt:'int'}); put(w4,'D'+r,fx(`B${r}*C${r}`,x.v*x.n),{fmt:'money'}); put(w4,'E'+r,x.o,{wrap:true}); });
  put(w4,'A11','TOTAL POR ALUNO/ANO',{tot:true,bold:true}); put(w4,'D11',fx('SUM(D6:D10)',pacoteTotal()),{tot:true,bold:true,fmt:'money'});
  put(w4,'A12','Nota: medicação não integra este modelo. Ver aba 6.',{it:true});
  title(w4,'A14','B. CUSTO AGREGADO PARA '+ctx.label.toUpperCase());
  hdr(w4,15,['Componente','Alunos','Custo por aluno/ano (R$)','Total anual (R$)','Observação','Na proposta (1=sim)']);
  const cl = byId('cli'), clOvQ = S.ovQ.cli!=null, clOvP = S.ovPU.cli!=null;
  put(w4,'A16','Atendimento profissional');
  put(w4,'B16', clOvQ ? cl.q : fx(P2+'$B$9', cl.qa), clOvQ?{inp:true,ov:true,fmt:'int',note:'Ajustado no painel. Cálculo-padrão: '+num(cl.qa)}:{fmt:'int'});
  put(w4,'C16', clOvP ? cl.pu : fx('D11', pacoteTotal()), clOvP?{inp:true,ov:true,fmt:'money',note:'Ajustado no painel. Pacote-padrão: '+brl(pacoteTotal())}:{fmt:'money'});
  put(w4,'D16',fx('B16*C16',cl.tot),{fmt:'money'}); put(w4,'E16','Aplica o percentual clínico definido na aba 1.',{wrap:true}); put(w4,'F16',inc('cli'),{inp:true,fmt:'0'});
  put(w4,'A17','TOTAL ANUAL — CUIDADO CLÍNICO',{tot:true,bold:true}); put(w4,'D17',fx('SUM(D16:D16)',cl.tot),{tot:true,bold:true,fmt:'money'});

  /* ---- 5. Consolidado ---- */
  const w5 = wb.addWorksheet('5. Consolidado', {properties:{tabColor:{argb:'FF6A3FA0'}}});
  w5.columns=[{width:46},{width:22},{width:22},{width:24}];
  title(w5,'A1','CONSOLIDAÇÃO E CENÁRIOS — '+ctx.label.toUpperCase());
  hdr(w5,3,['Bloco','Total anual (R$)','% do programa','Por aluno atendido (R$/ano)']);
  const tot = C.cen.C.ano||1, ac = d.acomp||1;
  [['Plataforma, implantação, formação e suporte',P3+'$E$15',C.tA.ano],['Módulo opcional de avaliação gamificada',P3+'$E$22',C.tB.ano],['Cuidado clínico — atendimento profissional',P4+'$D$16',C.tC.ano]]
    .forEach((x,i)=>{ const r=4+i; put(w5,'A'+r,x[0]); put(w5,'B'+r,fx(x[1],x[2]),{fmt:'money'}); put(w5,'C'+r,fx(`IFERROR(B${r}/$B$7,0)`,x[2]/tot),{fmt:'0.0%'}); put(w5,'D'+r,fx(`IFERROR(B${r}/${P2}$B$8,0)`,x[2]/ac),{fmt:'money'}); });
  put(w5,'A7','TOTAL DO PROGRAMA INTEGRAL',{tot:true,bold:true}); put(w5,'B7',fx('SUM(B4:B6)',C.cen.C.ano),{tot:true,bold:true,fmt:'money'}); put(w5,'C7',fx('SUM(C4:C6)',1),{tot:true,bold:true,fmt:'0.0%'}); put(w5,'D7',fx(`IFERROR(B7/${P2}$B$8,0)`,C.cen.C.ano/ac),{tot:true,bold:true,fmt:'money'});
  title(w5,'A9','CENÁRIOS DE ESCOPO');
  hdr(w5,10,['Cenário','Composição','Total anual (R$)','Por aluno (R$/ano)']);
  [['A — Coordenação','Somente plataforma de PEI, implantação, formação e suporte.',P3+'$E$15',C.cen.A.ano],['B — Coordenação e avaliação','Cenário A mais o módulo de avaliação e laboratórios.',P3+'$E$15+'+P3+'$E$22',C.cen.B.ano],['C — Programa integral (recomendado)','Cenário B mais o atendimento clínico profissional.',P3+'$E$15+'+P3+'$E$22+'+P4+'$D$16',C.cen.C.ano]]
    .forEach((x,i)=>{ const r=11+i; put(w5,'A'+r,x[0]); put(w5,'B'+r,x[1],{wrap:true}); put(w5,'C'+r,fx(x[2],x[3]),{fmt:'money'}); put(w5,'D'+r,fx(`IFERROR(C${r}/${P2}$B$8,0)`,x[3]/ac),{fmt:'money'}); });
  title(w5,'A15','CONTRATO DE 5 ANOS (serviço contínuo, arts. 106 e 107 da Lei 14.133/2021)');
  hdr(w5,16,['Cenário','Ano 1','Anos 2 a 5 (cada)','Total 5 anos']);
  ['A — Coordenação','B — Coordenação e avaliação','C — Programa integral (recomendado)'].forEach((t,i)=>{ const r=17+i, k='ABC'[i], cz=C.cen[k];
    put(w5,'A'+r,t); put(w5,'B'+r,fx('C'+(11+i),cz.ano),{fmt:'money'}); put(w5,'C'+r,fx(`C${11+i}-${P3}$E$10-${P3}$E$11`,cz.r),{fmt:'money'}); put(w5,'D'+r,fx(`B${r}+C${r}*4`,cz.nr+cz.r*5),{fmt:'money'}); });
  put(w5,'A20','Nos anos 2 a 5 excluem-se implantação e integração, que são despesas não recorrentes. Valores a preços de setembro de 2026, sem reajuste.',{it:true});
  title(w5,'A22','PROPOSTA — ESCOPO E VIGÊNCIA ADOTADOS');
  hdr(w5,23,['Componente','Valor','Cálculo']);
  const sumInc = `SUMPRODUCT(${P3}E5:E14,${P3}G5:G14)+SUMPRODUCT(${P3}E20:E21,${P3}G20:G21)+${P4}D16*${P4}F16`;
  [['Valor anual do escopo (1º ano)',sumInc,C.t.ano,'Soma dos itens marcados com 1'],
   ['Parcela não recorrente (implantação e integração)',`${P3}E10*${P3}G10+${P3}E11*${P3}G11`,C.t.nr,'Somente no 1º ano'],
   ['Parcela recorrente anual','B24-B25',C.t.r,'Repete-se a cada 12 meses'],
   ['Vigência (meses)',P1+'B52',S.vig,'Aba 1, célula B52'],
   ['VALOR TOTAL DO PROJETO','B25+B26*B27/12',C.total,'Não recorrente + recorrente × vigência/12'],
   ['Por estudante/mês',`IFERROR(B28/${P2}B8/B27,0)`,d.acomp?C.total/d.acomp/S.vig:0,'Sobre o público acompanhado']]
   .forEach((x,i)=>{ const r=24+i, key=(r===28); put(w5,'A'+r,x[0],{bold:key,tot:key}); put(w5,'B'+r,fx(x[1],x[2]),{fmt:r===27?'int':'money',bold:key,tot:key}); put(w5,'C'+r,x[3],{tot:key}); });

  /* ---- 6 e 7. Anexos estáticos ---- */
  ['6. Canabidiol','7. Fontes'].forEach(nome=>{
    const src = XLSX_STATIC[nome], ws = wb.addWorksheet(nome, {properties:{tabColor:{argb:'FF8A90AE'}}});
    Object.entries(src.widths).forEach(([k,w])=>{ ws.getColumn(k).width = w; });
    src.cells.forEach(([a,v,bo,fl])=>{ const c=ws.getCell(a); c.value = typeof v==='string' ? v.replace(/Guarujá(\/SP)?/g, ctx.label) : v;
      c.font={name:F,size:/^A1$/.test(a)?12:11,bold:bo}; if(fl && fl!=='00000000') c.fill=fill(fl); c.alignment={wrapText:true,vertical:'top'}; });
    src.merged.forEach(r=>ws.mergeCells(r));
  });

  /* ---- 8. Municipios ---- */
  const w8 = wb.addWorksheet('8. Municipios', {properties:{tabColor:{argb:'FF8A90AE'}}});
  w8.columns=[{width:14},{width:6},{width:34},{width:12},{width:12},{width:13},{width:13},{width:14}];
  title(w8,'A1','BASE MUNICIPAL — PESSOAS COM TEA POR FAIXA ETÁRIA (IBGE, Censo 2022, SIDRA 10145)');
  put(w8,'A2','Municípios que compõem o ente desta memória. A aba 1 soma estas linhas.');
  hdr(w8,4,['Código IBGE','UF','Município','0 a 4 anos','5 a 9 anos','10 a 14 anos','15 a 19 anos','Total 0 a 19']);
  idx.forEach((i,k)=>{ const r=5+k, m=M[i]; w8.getRow(r).values=[m[0],m[1],m[2],m[3],m[4],m[5],m[6],{formula:`SUM(D${r}:G${r})`,result:m[3]+m[4]+m[5]+m[6]}];
    w8.getRow(r).eachCell(c=>{ c.font={name:F,size:11}; }); ['D','E','F','G','H'].forEach(col=>w8.getCell(col+r).numFmt=FM.int); });
  const rt = mR1+1; put(w8,'C'+rt,'TOTAL',{tot:true,bold:true});
  ['D','E','F','G','H'].forEach((col,k)=>put(w8,col+rt,fx(`SUM(${col}${mR0}:${col}${mR1})`,k<4?b[k]:d.total),{tot:true,bold:true,fmt:'int'}));
  w8.views=[{state:'frozen',ySplit:4}];

  wb.worksheets.forEach(ws=>{ ws.pageSetup={paperSize:9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight:0}; });
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

/* =====================================================================
   ETP e TR (.docx)
   ===================================================================== */
/* Padrão de edital para todo Word/PDF gerado: Times New Roman, preto, justificado, entrelinha 1,5, margens ABNT. */
const EDITAL = {fonte:'Times New Roman', corpo:12, cabecalho:12, tabela:10, nota:10, entrelinha:1.5, margem:{sup:3, esq:3, inf:2, dir:2}};
function docKit(ctx, sigla){
  const X = window.docx;
  const {Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel, BorderStyle, VerticalAlign} = X;
  const E = EDITAL, hp = pt => Math.round(pt*2), cm = v => Math.round(v*567), LINE = Math.round(240*E.entrelinha);
  const NAVY='000000', SOFT=null;                       // mantidos por compatibilidade: sem cor, sem fundo
  const PRETO = {style:BorderStyle.SINGLE, size:4, color:'000000'};
  let alinea = 0;                                        // a), b), c)… — reinicia a cada parágrafo, título ou tabela
  const runs = (t, o={}) => String(t).split('**').map((s,i)=>new TextRun({text:s, bold:!!o.bold || i%2===1, italics:!!o.it, size:o.size, font:E.fonte, color:'000000'}));
  const p = (t, o={}) => { alinea = 0; const nota = o.size && o.size < hp(E.corpo);
    return new Paragraph({children:runs(t,{...o, size:nota?hp(E.nota):hp(E.corpo)}), spacing:{after:o.after??120, before:o.before??0, line:nota?240:LINE},
      alignment:o.align??AlignmentType.JUSTIFIED, keepNext:!!o.keepNext}); };
  const h1 = (t, sub) => { alinea = 0; return [new Paragraph({heading:HeadingLevel.HEADING_1, children:[new TextRun({text:t.toUpperCase(), font:E.fonte, bold:true, size:hp(E.corpo), color:'000000'})],
      spacing:{before:360, after:sub?0:120, line:LINE}, keepNext:true}),
    ...(sub ? [new Paragraph({children:[new TextRun({text:sub, italics:true, font:E.fonte, size:hp(E.corpo), color:'000000'})], spacing:{after:120, line:LINE}, keepNext:true})] : [])]; };
  const h2 = t => { alinea = 0; return new Paragraph({heading:HeadingLevel.HEADING_2, children:[new TextRun({text:t, font:E.fonte, bold:true, size:hp(E.corpo), color:'000000'})], spacing:{before:240, after:120, line:LINE}, keepNext:true}); };
  const bl = t => { const l = String.fromCharCode(97 + (alinea++ % 26));
    return new Paragraph({children:[new TextRun({text:l+') ', font:E.fonte, size:hp(E.corpo), color:'000000'}), ...runs(t,{size:hp(E.corpo)})],
      indent:{left:cm(1.25), hanging:cm(0.63)}, spacing:{after:80, line:LINE}, alignment:AlignmentType.JUSTIFIED}); };
  const cell = (c, o={}) => {
    const x = (c && c.__c) ? Object.assign({}, o, c) : Object.assign({}, o, {t:c});
    const paras = String(x.t==null?'':x.t).split('\n').map(line=>new Paragraph({children:runs(line,{bold:x.bold, size:hp(E.tabela)}),
      alignment:x.right?AlignmentType.RIGHT:(x.center?AlignmentType.CENTER:AlignmentType.LEFT), spacing:{after:0, line:240}}));
    return new TableCell({children:paras, columnSpan:x.span, verticalAlign:VerticalAlign.CENTER, margins:{top:40, bottom:40, left:80, right:80},
      width:x.w?{size:x.w, type:WidthType.PERCENTAGE}:undefined});
  };
  const tbl = (head, rows, w, o={}) => { alinea = 0; return new Table({width:{size:100, type:WidthType.PERCENTAGE},
    borders:{top:PRETO, bottom:PRETO, left:PRETO, right:PRETO, insideHorizontal:PRETO, insideVertical:PRETO},
    rows:[...(head?[new TableRow({tableHeader:true, cantSplit:true, children:head.map((h,i)=>cell(h,{bold:true, center:true, w:w&&w[i]}))})]:[]),
      ...rows.map(r=>new TableRow({cantSplit:true, children:r.map((c,i)=>cell(c,{w:w&&w[i], right:(o.right||[]).includes(i)}))}))]}); };
  const kv = pairs => tbl(null, pairs.map(([k,v])=>[{__c:1, t:k, bold:true}, v]), [30,70]);
  const box = (titulo, linhas) => { alinea = 0; return new Table({width:{size:100, type:WidthType.PERCENTAGE},
    borders:{top:PRETO, bottom:PRETO, left:PRETO, right:PRETO, insideHorizontal:PRETO, insideVertical:PRETO},
    rows:[new TableRow({cantSplit:true, children:[new TableCell({margins:{top:100,bottom:100,left:140,right:140},
      children:[new Paragraph({children:[new TextRun({text:titulo.toUpperCase(), bold:true, font:E.fonte, size:hp(E.corpo), color:'000000'})], spacing:{after:80, line:LINE}}), ...linhas.map(t=>p(t,{after:60}))]})]})]}); };
  const gap = (n=120) => new Paragraph({children:[], spacing:{after:n}});
  const centro = (t, o={}) => new Paragraph({children:runs(t,{...o, size:hp(E.corpo)}), alignment:AlignmentType.CENTER, spacing:{after:o.after??0, line:LINE}});
  const assinaturas = () => [gap(240), p(PRE('Local')+', '+PRE('data')+'.',{align:AlignmentType.RIGHT}), gap(480),
    centro('_______________________________________'), centro('Equipe de Planejamento da Contratação',{bold:true}), centro(PRE('Nome / matrícula / cargo'),{after:480}),
    centro('_______________________________________'), centro('Autoridade competente — aprovação',{bold:true}), centro(PRE('Nome / cargo'))];
  const documento = (titulo, children) => new X.Document({creator:'G3 Health Service', title:titulo, description:'Minuta — Ref. '+ctx.ref,
    styles:{default:{document:{run:{font:E.fonte, size:hp(E.corpo), color:'000000'}, paragraph:{spacing:{line:LINE}}},
      heading1:{run:{font:E.fonte, size:hp(E.corpo), bold:true, color:'000000'}}, heading2:{run:{font:E.fonte, size:hp(E.corpo), bold:true, color:'000000'}}}},
    sections:[{properties:{page:{margin:{top:cm(E.margem.sup), left:cm(E.margem.esq), bottom:cm(E.margem.inf), right:cm(E.margem.dir), header:cm(1.25), footer:cm(1.0)}}},
      headers:{default:new X.Header({children:[
        new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:0}, children:[new TextRun({text:ctx.orgao.toUpperCase(), bold:true, font:E.fonte, size:hp(E.cabecalho), color:'000000'})]}),
        new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:120}, border:{bottom:{style:BorderStyle.SINGLE, size:6, color:'000000', space:4}},
          children:[new TextRun({text:titulo.split(' — ')[0]+' — Minuta', font:E.fonte, size:hp(E.cabecalho), color:'000000'})]})]})},
      footers:{default:new X.Footer({children:[new Paragraph({alignment:AlignmentType.CENTER, children:[
        new TextRun({text:'Ref. '+ctx.ref+'   ·   ', font:E.fonte, size:hp(E.nota), color:'000000'}),
        new TextRun({children:['Página ', X.PageNumber.CURRENT, ' de ', X.PageNumber.TOTAL_PAGES], font:E.fonte, size:hp(E.nota), color:'000000'})]})]})},
      children}]});
  const capa = (rotulo, titulo, subtitulo, pares) => [
    centro('**'+titulo.replace(/ — (ETP|TR)$/,' ($1)').toUpperCase()+'**'), centro('(MINUTA)',{after:240}),
    p('**Objeto:** '+subtitulo, {after:200}), kv(pares), gap(200)];
  return {X, p, h1, h2, bl, tbl, kv, box, gap, assinaturas, documento, capa, NAVY, SOFT};
}
function paresCapa(ctx, fundamento, extra){
  const pares = [['FUNDAMENTO', fundamento], ['ÓRGÃO DEMANDANTE', ctx.orgao], ['ESFERA', ctx.esfera], ['TIPO DE CONTRATAÇÃO', MODOS[ctx.modo].t],
    ['PROCESSO ADMINISTRATIVO', S.doc.proc.trim()||PRE()], ['REFERÊNCIA', ctx.ref], ...(extra||[]), ['DATA', ctx.data.charAt(0).toUpperCase()+ctx.data.slice(1)]];
  if(S.doc.g3){ pares.push(['ELABORAÇÃO TÉCNICA','G3 Health Service Ltda. — CNPJ 31.652.744/0001-14']); pares.push(['SOLUÇÃO DE REFERÊNCIA','TEAlliance — ecossistema Inclusi.Via (Core, School, Clinic e Family)']); }
  else pares.push(['ELABORAÇÃO', 'Equipe de Planejamento da Contratação — '+PRE('nomes e matrículas')]);
  return pares;
}
function cadeiaTxt(ctx, atual){
  const it = [['ETP','Estudo Técnico Preliminar — por que contratar, quanto e como (art. 18)'],['ANEXO','Anexo I do ETP — Memória de cálculo em planilha, com fórmulas'],['TR','Termo de Referência — o que, como executar, medir e pagar (art. 6º, XXIII)'],['PROP','Proposta comercial de referência — insumo da pesquisa de preços (art. 23, § 1º, IV)']];
  return 'Cadeia documental da referência '+ctx.ref+': '+it.map(([k,t],i)=>(i+1)+') '+(k===atual?'**'+t+' — este documento**':t)).join('; ')+'. Todas as peças usam os mesmos quantitativos e valores.';
}
function tabelaItens(K, L, cols){
  const rows=[];
  L.forEach(lt=>{
    rows.push([{__c:1, t:'LOTE '+lt.n+' — '+lt.nome.toUpperCase()+'  ·  rubrica: '+lt.rub, span:cols.length, bold:true, color:K.NAVY, fill:K.SOFT, size:17}]);
    lt.itens.forEach(x=>rows.push(cols.map(c=>c.v(x))));
    const sc = cols[cols.length-1];
    if(sc.sum) rows.push([{__c:1, t:'Subtotal do lote '+lt.n, bold:true, span:cols.length-1}, {__c:1, t:sc.sum(lt), bold:true, right:true}]);
  });
  return rows;
}

/* ---------------- ETP ---------------- */
async function gerarETP(ctx){
  await carregar('docx');
  const K = docKit(ctx,'ETP'), {p,h1,h2,bl,tbl,kv,box,gap} = K, {d, L} = ctx;
  const c = [];
  const objeto = 'Contratação de solução integrada de apoio à inclusão escolar e à coordenação do cuidado de educandos com transtorno do espectro autista (TEA) e demais público-alvo da educação especial, compreendendo '+ctx.partes.join('; ');
  c.push(...K.capa('ESTUDO TÉCNICO PRELIMINAR  ·  MINUTA','Estudo Técnico Preliminar — ETP', objeto, paresCapa(ctx,'Art. 18, § 1º, incisos I a XIII, da Lei nº 14.133, de 1º de abril de 2021',[['PÚBLICO ESTIMADO', int(d.acomp)+' educandos acompanhados'+(ctx.cams.has('C')?' · '+int(d.clin)+' com atendimento clínico':'')],['VALOR ESTIMADO', brl(ctx.total)+' para '+ctx.vig+' meses ('+brl(ctx.ano)+' no 1º ano)']])));
  c.push(box('Nota de uso deste documento',[
    'Minuta estruturante, gerada para servir de base técnica ao órgão demandante. Os quantitativos e valores já vêm calculados a partir do Censo 2022 (IBGE) e dos parâmetros do Anexo I — Memória de cálculo; os campos '+PRE()+' dependem de dados locais (Censo Escolar, PCA, dotação orçamentária).',
    'Os valores apresentados no item VI são estimativa preliminar de referência e não substituem a pesquisa de preços exigida pelo art. 23 da Lei nº 14.133/2021, que deve ser realizada pelo órgão antes da deflagração do certame.',
    'O documento foi redigido a partir da necessidade pública e dos deveres legais do ente, e não das características de um fornecedor específico. Os requisitos admitem mais de uma solução no mercado (arts. 9º, I, e 25, § 1º, da Lei nº 14.133/2021).',
    cadeiaTxt(ctx,'ETP')]));

  c.push(...h1('I — Descrição da necessidade pública','(art. 18, § 1º, I)'), h2('I.1. O problema a ser resolvido'));
  c.push(p('Segundo o Censo Demográfico 2022 (IBGE, SIDRA 10145), '+ctx.label+' registra **'+int(d.total)+' pessoas de 0 a 19 anos com diagnóstico de transtorno do espectro autista** informado por profissional de saúde, das quais '+int(d.recorte)+' no recorte etário de '+ctx.recorteTxt+'. Aplicados os parâmetros de rede pública, de ampliação para o público-alvo do AEE e de adesão detalhados no item IV, estima-se em **'+int(d.acomp)+' educandos** o público a ser acompanhado na '+redeTxt()+'.'));
  c.push(p('A '+redeTxt()+' atende '+PRE('nº')+' matrículas na educação básica, das quais '+PRE('nº')+' na educação especial (Censo Escolar '+PRE('ano')+'). O atendimento a esse público hoje se dá de forma fragmentada: a escola registra avaliações e planos em documentos físicos ou planilhas isoladas; a rede de saúde produz laudos e relatórios terapêuticos que raramente chegam ao professor; e a família, que detém a observação mais contínua do desenvolvimento, não dispõe de canal estruturado para contribuir nem para acompanhar. O resultado é conhecido e mensurável:'));
  ['Plano Educacional Individualizado inexistente, desatualizado ou meramente formal, elaborado sem a informação clínica disponível e sem revisão periódica documentada.',
   'Retrabalho e descontinuidade terapêutica — a intervenção realizada no consultório não se reflete na rotina escolar, e o professor não é informado de mudanças no quadro clínico do educando.',
   'Perda de histórico nas transições — troca de professor, mudança de unidade ou passagem entre etapas implicam recomeço do processo de conhecimento do educando.',
   'Ausência de indicadores objetivos de evolução, o que impede a gestão de saber se as medidas adotadas produzem resultado e impede a prestação de contas qualificada.',
   'Exposição a risco jurídico — impossibilidade de demonstrar documentalmente o cumprimento das obrigações legais, com aumento da judicialização individual.'].forEach(t=>c.push(bl(t)));
  if(ctx.cams.has('C')) c.push(p('Soma-se a demanda assistencial: a rede de saúde não dispõe de oferta regular e articulada à escola de consultas especializadas e de terapias para o público-alvo, o que alonga filas e desconecta o plano terapêutico do plano educacional.'));
  c.push(h2('I.2. O dever legal que a contratação cumpre'), p('A necessidade não é discricionária. Decorre de obrigações legais expressas e exigíveis:'));
  c.push(tbl(['Norma','Obrigação imposta ao poder público'],[
    ['Lei nº 14.254/2021','Acompanhamento integral para educandos com dislexia, TDAH ou outro transtorno de aprendizagem, em articulação com os serviços de saúde: identificação precoce, acompanhamento, apoio pedagógico e encaminhamento.'],
    ['Lei nº 13.146/2015 (LBI), art. 28','Serviços e recursos de acessibilidade; medidas individualizadas que maximizem o desenvolvimento; formação continuada; participação da família.'],
    ['Lei nº 12.764/2012','Política Nacional de Proteção dos Direitos da Pessoa com TEA — diagnóstico precoce, atendimento multiprofissional e acesso à educação.'],
    ['Lei nº 9.394/1996 (LDB), arts. 58 a 60','Educação especial preferencialmente na rede regular, com serviços de apoio especializado.'],
    ['Decreto nº 7.611/2011','Atendimento Educacional Especializado, com apoio técnico e financeiro da União e duplo cômputo de matrícula.'],
    ['Lei nº 13.257/2016','Marco Legal da Primeira Infância — articulação intersetorial entre saúde, educação e assistência social.'],
    ['Resolução CNE/CEB nº 4/2009','Diretrizes operacionais do AEE, com plano de atendimento individualizado.'],
    ['Lei nº 13.709/2018 (LGPD)','Tratamento de dados sensíveis de crianças e adolescentes no melhor interesse e com as salvaguardas do art. 14.']],[32,68],{zebra:true}));
  c.push(gap(), box('Síntese da necessidade',['A Administração necessita de infraestrutura digital única que integre as informações de educação, saúde e família em torno de cada educando, gere e mantenha o Plano Educacional Individualizado, permita execução coordenada entre professores, terapeutas e responsáveis e produza indicadores objetivos e auditáveis de evolução'+(ctx.cams.has('C')?', articulada a oferta regular de atendimento multiprofissional':'')+' — cumprindo obrigação legal expressa e reduzindo risco jurídico e retrabalho.'],'EEF0F8'));

  c.push(...h1('II — Previsão no Plano de Contratações Anual','(art. 18, § 1º, II)'));
  c.push(p('A contratação encontra-se prevista no Plano de Contratações Anual — PCA '+(S.doc.pca.trim()||PRE('exercício e nº do item'))+', alinhada ao Plano '+(S.tipo==='uf'?'Estadual':'Municipal')+' de Educação e às metas de inclusão escolar nele estabelecidas. Caso não conste do PCA vigente, deverá ser promovida sua inclusão antes da deflagração do procedimento.'));

  c.push(...h1('III — Requisitos da contratação','(art. 18, § 1º, III)'), h2('III.1. Requisitos de negócio'));
  ['Solução em nuvem (SaaS), sem necessidade de instalação de servidores ou de infraestrutura própria pela Administração.',
   'Módulos distintos e integrados para escola, saúde e família, operando sobre base de dados única por educando.',
   'Geração assistida por inteligência artificial do PEI, com revisão e homologação humana obrigatórias antes da validação.',
   'Registro estruturado de objetivos, metas, estratégias, recursos de acessibilidade e responsáveis, com controle de versões.',
   'Painel de indicadores por educando, por unidade escolar e consolidado para a Secretaria.',
   'Implantação, formação continuada, supervisão de casos e suporte técnico como serviços integrantes do objeto — o resultado depende da operação, e não apenas do software.'].forEach(t=>c.push(bl(t)));
  c.push(h2('III.2. Requisitos tecnológicos'));
  c.push(tbl(['Frente','Requisito mínimo'],[
    ['Arquitetura','Aplicação web responsiva, compatível com as duas últimas versões estáveis dos navegadores de mercado; hospedagem em território nacional.'],
    ['Disponibilidade','Índice mensal mínimo de 99,5%, apurado conforme o Acordo de Nível de Serviço do Termo de Referência.'],
    ['Interoperabilidade','API REST documentada em OpenAPI 3.x; integração com o sistema de gestão escolar, o Educacenso/INEP e sistemas de saúde; HL7 FHIR R4 desejável.'],
    ['Autenticação','SSO (SAML 2.0 e/ou OAuth 2.0 com OpenID Connect); perfis granulares; duplo fator para perfis administrativos.'],
    ['Acessibilidade','eMAG e WCAG 2.1 nível AA.'],
    ['Segurança','TLS 1.2 ou superior; criptografia em repouso; trilha de auditoria imutável; backup com RPO e RTO declarados.'],
    ['Proteção de dados','LGPD; RIPD na implantação; encarregado nomeado; matriz controlador-operador; consentimento parental (art. 14).'],
    ['Governança de IA','Documentação dos modelos; explicabilidade; registro de revisão humana; canal de contestação.'],
    ['Portabilidade','Exportação integral em formato aberto (CSV, JSON ou XML), a qualquer tempo, sem custo.']],[24,76],{zebra:true}));
  c.push(h2('III.3. Requisitos de sustentabilidade'), p('A solução dispensa equipamentos e infraestrutura física própria e elimina a impressão e o arquivamento em papel do PEI e dos registros de acompanhamento, em atendimento ao art. 11, IV, da Lei nº 14.133/2021.'));
  c.push(h2('III.4. Requisitos de habilitação técnica'));
  ['Atestado(s) de capacidade técnica de fornecimento de solução compatível com o objeto.','Comprovação de titularidade ou de direito de uso do software ofertado.','Equipe técnica mínima com profissional de educação especial ou psicopedagogia e responsável técnico de TI.','Prova de conceito (PoC), com roteiro objetivo previamente divulgado.'].forEach(t=>c.push(bl(t)));
  if(ctx.cams.has('B')){ c.push(h2('III.5. Requisitos específicos da avaliação pedagógica')); ['Instrumento de avaliação pedagógica de habilidades — leitura, escrita, numeramento e atenção — com resultado de natureza pedagógica.','Vedação à emissão de hipótese diagnóstica ou à nomeação de condição de saúde: o instrumento sinaliza necessidade de aprofundamento e encaminha ao profissional habilitado, o que mantém a solução fora do regime de software como dispositivo médico (RDC ANVISA nº 657/2022 e nº 751/2022).','Importação estruturada dos resultados pela plataforma de gestão do PEI.'].forEach(t=>c.push(bl(t))); }
  if(ctx.cams.has('C')){ c.push(h2('III.'+(ctx.cams.has('B')?6:5)+'. Requisitos específicos do atendimento multiprofissional')); ['Profissionais legalmente habilitados, com registro ativo no conselho de classe (CRM, com Registro de Qualificação de Especialista em neurologia ou psiquiatria; CRFa; CRP) e formação comprovada em psicopedagogia.','Estabelecimentos cadastrados no CNES; teleatendimento admitido nos termos da Lei nº 14.510/2022.','Registro de cada atendimento no módulo de saúde da plataforma, com devolutiva ao plano educacional.','Periodicidade de referência: 3 consultas/ano de neurologia e de psiquiatria e 12 sessões/ano de fonoaudiologia, psicologia e psicopedagogia, ajustável por nível de suporte.'].forEach(t=>c.push(bl(t))); }

  c.push(...h1('IV — Estimativa das quantidades','(art. 18, § 1º, IV)'));
  c.push(p('As quantidades foram dimensionadas pelo público efetivamente atendido pela educação especial, e não pela matrícula total — critério tecnicamente mais defensável, que reduz o valor global e evita questionamento por superdimensionamento. A memória completa, com fórmulas, integra este ETP como **Anexo I — Memória de cálculo**.'));
  c.push(h2('IV.1. Base populacional — Censo 2022'));
  const sr = d.rec.reduce((a,x)=>a+x,0)||1;
  c.push(tbl(['Faixa etária','Etapa de ensino','Pessoas com TEA','% no recorte','No recorte','Acompanhados'],
    FAIXAS.map((f,k)=>[f.k, f.etapa, int(d.b[k]), num(S.f[k]*100,0)+'%', int(d.rec[k]), int(Math.round(d.acomp*d.rec[k]/sr))]).concat([[{__c:1,t:'Total',bold:true},'',{__c:1,t:int(d.total),bold:true},'',{__c:1,t:int(d.recorte),bold:true},{__c:1,t:int(d.acomp),bold:true}]]),
    [16,32,14,12,12,14],{right:[2,3,4,5]}));
  c.push(h2('IV.2. Do total populacional ao público atendido'));
  c.push(tbl(['Etapa','Educandos','Parâmetro aplicado'],[
    ['Pessoas com TEA no recorte etário',int(d.recorte),'IBGE, Censo 2022 — distribuição uniforme dentro da faixa quinquenal'],
    ['Em rede pública',int(d.rede),num(P.rede,1)+'% — a confirmar no Censo Escolar'],
    ['Ampliado para o público-alvo do AEE',int(d.aee),'Fator '+num(P.aee,2)+' — demais deficiências, TGD e altas habilidades'],
    ['Com adesão efetiva',{__c:1,t:int(d.acomp),bold:true},num(P.adesao,1)+'% — recusa familiar e perda de seguimento'],
    ...(ctx.cams.has('C')?[['Com atendimento clínico',int(d.clin),num(P.clin,1)+'% do público acompanhado — parâmetro de política']]:[])],[40,16,44],{right:[1]}));
  c.push(h2('IV.3. Itens, quantidades e memória de cálculo'));
  c.push(tbl(['#','Item','Unidade','Qtd. (ano 1)','Memória de cálculo'], tabelaItens(K, L, [{v:x=>String(x.n)},{v:x=>x.desc},{v:x=>x.un},{v:x=>({__c:1,t:num(x.q),right:true})},{v:x=>x.memo}]), [5,33,13,11,38]));
  if(ctx.part.length){ c.push(h2('IV.4. Estimativa por município participante'));
    c.push(tbl(['Município','TEA 0–19','No recorte','Acompanhados','Valor anual estimado'], ctx.part.map(x=>[x.nome,int(x.total),int(x.recorte),int(x.acomp),brl(x.ano)]).concat([[{__c:1,t:'Total',bold:true},int(ctx.part.reduce((a,x)=>a+x.total,0)),int(ctx.part.reduce((a,x)=>a+x.recorte,0)),int(ctx.part.reduce((a,x)=>a+x.acomp,0)),{__c:1,t:brl(ctx.part.reduce((a,x)=>a+x.ano,0)),bold:true}]]),[34,14,14,16,22],{right:[1,2,3,4]}));
    c.push(p('A estimativa individual serve à Intenção de Registro de Preços e à distribuição das quantidades entre participantes. A soma pode diferir ligeiramente do consolidado por arredondamento e pelo mínimo de uma unidade escolar por município.',{size:18}));
  }
  c.push(gap(), box('Advertência de dimensionamento',['Os parâmetros de rede, de ampliação e de estrutura escolar (unidades, profissionais, horas) são provisórios e estão sinalizados no Anexo I. Devem ser substituídos pelos dados do Censo Escolar e da Secretaria antes da publicação. '+(['srp','srpcons'].includes(ctx.modo)?'Em Registro de Preços, as quantidades são estimativa máxima, sem obrigação de contratação; recomenda-se implantação faseada, com o primeiro exercício cobrindo de 20% a 30% da rede.':'Dimensionamento inflado é a causa mais comum de impugnação e glosa.')]));

  c.push(...h1('V — Levantamento de mercado e justificativa da escolha','(art. 18, § 1º, V)'), h2('V.1. Alternativas examinadas'));
  c.push(tbl(['Alternativa','Descrição','Avaliação'],[
    ['A — Manutenção da situação atual','PEI em papel ou planilha, sem integração com a saúde e sem indicadores.','Descartada. Não cumpre a Lei nº 14.254/2021 nem produz evidência auditável.'],
    ['B — Desenvolvimento interno','Construção de sistema próprio pela TI do órgão.','Descartada. Sem equipe de desenvolvimento nem base de conhecimento pedagógico-clínico; custo e prazo incompatíveis.'],
    ['C — Módulo do sistema de gestão escolar','Ampliação do sistema de gestão escolar contratado.','Descartada. Trata matrícula, frequência e notas; sem camada clínica nem geração assistida de PEI.'],
    ['D — Licença isolada de software','Aquisição de licenças, sem serviço de implantação, formação e supervisão.','Descartada. O mercado público de referência compra serviço especializado com plataforma acoplada; licença sem operação não produz adesão nem resultado.'],
    ['E — Serviço especializado com plataforma acoplada','Solução SaaS com IA, módulos escola, saúde e família, com implantação, formação, supervisão e suporte'+(ctx.cams.has('B')?', avaliação pedagógica':'')+(ctx.cams.has('C')?' e atendimento multiprofissional':'')+'.',{__c:1,t:'Adotada. Única alternativa que endereça integralmente a necessidade do item I.',bold:true}]],[24,40,36],{zebra:true}));
  c.push(h2('V.2. Contratações públicas identificadas'));
  c.push(tbl(['Órgão','Objeto','Valor'],[
    ['CIMPAR — Consórcio Intermunicipal do Vale do Paraibuna','Solução integrada de apoio técnico-pedagógico ao AEE nos municípios consorciados — Pregão Eletrônico 07/2026, homologado em 16/07/2026','R$ 116.389.239,08'],
    ['CIMINAS — Consórcio Interfederativo Minas Gerais','Plataforma de identificação com avaliação gamificada, implantação, treinamento e laboratórios itinerantes — ARP 045/2026','R$ 80.400.000,00'],
    ['Chorozinho/CE','Locação de software de gestão educacional, sem módulo de PEI ou AEE — 12 meses','R$ 360.000,00']],[30,50,20],{right:[2]}));
  c.push(p('As contratações de maior porte foram celebradas por consórcios públicos e compram serviço e operação, não licença. O levantamento identificou pluralidade de fornecedores de soluções de gestão de PEI; nenhum requisito deste ETP é exclusivo de fornecedor único. A pesquisa formal de fornecedores e de preços deverá seguir o art. 23 da Lei nº 14.133/2021. Recomenda-se consulta direta ao PNCP.'));
  c.push(p('Quanto ao enquadramento regulatório: soluções que se propõem a identificar condições de saúde nomeadas tendem a configurar Software como Dispositivo Médico (RDC ANVISA nº 657/2022 e nº 751/2022). A solução especificada opera sobre dados clínicos laudados por profissional habilitado e gera produto de natureza pedagógica, o que afasta aquele risco.'));

  c.push(...h1('VI — Estimativa do valor da contratação','(art. 18, § 1º, VI)'), h2('VI.1. Método'));
  c.push(p('O valor estimado deverá ser apurado por pesquisa de preços conduzida pelo órgão, nos termos do art. 23 da Lei nº 14.133/2021: (a) sistemas oficiais e contratações similares da Administração; (b) contratações similares de outros entes nos 12 meses anteriores; (c) mídia ou sítio especializado; e (d) pesquisa direta com no mínimo três fornecedores. Os valores abaixo são **estimativa preliminar de referência**.'));
  c.push(h2('VI.2. Parâmetro público de referência — ARP nº 045/2026 (CIMINAS)'));
  c.push(K.kv([['Origem','Concorrência Eletrônica nº 004/2026, firmada em 27/08/2026'],['Licença de plataforma de avaliação','R$ 750,97 por aluno/ano'],['Implementação por unidade educacional','R$ 1.162,49 por unidade'],['Treinamento (turma de até 20)','R$ 3.520,84 por turma'],['Locação de laboratório itinerante','R$ 2.542,89 por serviço'],['Valor total registrado','R$ 80.400.000,00']]));
  c.push(p('O objeto da ARP 045/2026 não é idêntico: é útil para a ordem de grandeza dos itens de implantação, formação e avaliação, e não deve ser transposto diretamente para o licenciamento de gestão de PEI, dimensionado aqui por educando com PEI ativo.',{size:18, before:80}));
  c.push(h2('VI.3. Planilha de formação do valor estimado'));
  c.push(tbl(['#','Item','Unidade','Qtd.','Unitário','Total anual'], tabelaItens(K, L, [{v:x=>String(x.n)},{v:x=>x.desc+(x.rec?'':' *')},{v:x=>x.un},{v:x=>({__c:1,t:num(x.q),right:true})},{v:x=>({__c:1,t:brl(x.pu),right:true})},{v:x=>({__c:1,t:brl(x.tot),right:true}),sum:lt=>brl(lt.ano)}]), [5,37,13,10,15,20]));
  c.push(p('* Item não recorrente: incide apenas no primeiro ano.',{size:17, before:60}));
  c.push(h2('VI.4. Síntese do valor'));
  c.push(K.kv([['Valor do primeiro ano',brl(ctx.ano)],['Parcela não recorrente',brl(ctx.nr)],['Parcela recorrente anual',brl(ctx.ano-ctx.nr)],['Vigência considerada',ctx.vig+' meses'],['VALOR ESTIMADO PARA A VIGÊNCIA','**'+brl(ctx.total)+'**'],['Custo por educando/mês',brl(d.acomp?ctx.total/d.acomp/ctx.vig:0)]]));
  c.push(p('O detalhamento dos preços unitários poderá ter caráter sigiloso, nos termos do art. 24 da Lei nº 14.133/2021.',{size:18, before:80}));

  c.push(...h1('VII — Descrição da solução como um todo','(art. 18, § 1º, VII)'));
  c.push(p('A solução é descrita em seu ciclo de vida completo, e não apenas no momento da entrega do software. O que o ente contrata é a coordenação do desenvolvimento de cada educando; a plataforma é o instrumento.'));
  const comp = [['Núcleo de integração e inteligência','Base única por educando; cruzamento com base estruturada de conhecimento pedagógico, clínico e legal; IA para identificar necessidades, recomendar estratégias e gerar proposta de plano.'],['Módulo Escola','Ciclo completo do PEI: elaboração assistida, homologação humana, execução, revisão periódica e encerramento, com controle de prazos legais.'],['Módulo Saúde','Acesso controlado de profissionais da rede de saúde e conveniada, com registro de condutas, metas terapêuticas e devolutivas.'],['Módulo Família','Registro de observações do cotidiano, acompanhamento da evolução e orientação sobre estratégias domiciliares.'],['Painel de governança','Indicadores agregados para a Secretaria: cobertura de PEI, tempestividade, revisão no prazo, evolução.'],['Formação continuada','Trilhas por perfil (regente, AEE, coordenação, gestão, saúde), turmas de até 20, carga horária certificada.'],['Supervisão e suporte','Mentoria multiprofissional de casos e central de atendimento com níveis de serviço.']];
  if(ctx.cams.has('B')) comp.push(['Avaliação pedagógica','Avaliação gamificada de habilidades aplicada com laboratórios itinerantes; resultados importados para o PEI; sem finalidade diagnóstica.']);
  if(ctx.cams.has('C')) comp.push(['Atendimento multiprofissional','Neurologia, psiquiatria, fonoaudiologia, psicologia e psicopedagogia, com registro no módulo de saúde e devolutiva à escola.']);
  c.push(tbl(['Componente','Descrição funcional'], comp, [26,74], {zebra:true}));
  c.push(h2('VII.2. Escopo excluído — medicação'), p('Não integra o objeto o fornecimento de medicação à base de canabidiol: não incorporado ao SUS para TEA, sem registro na ANVISA dos produtos usualmente prescritos (logo, sem preço CMED de referência) e fora das indicações da Resolução CFM nº 2.324/2022. A plataforma registra a medicação prescrita, função legítima de coordenação do cuidado, sem custeio pelo contrato.'));

  c.push(...h1('VIII — Justificativa para o parcelamento ou não da solução','(art. 18, § 1º, VIII)'));
  c.push(p('Nos termos do art. 40, V, "b", da Lei nº 14.133/2021, o parcelamento é a regra quando tecnicamente viável e economicamente vantajoso.'));
  if(L.length>1){
    c.push(p('Propõe-se a divisão em **'+L.length+' lotes**, segundo a natureza do serviço e a rubrica orçamentária: '+L.map(x=>'Lote '+x.n+' — '+x.nome+' ('+x.rub+')').join('; ')+'.'));
    c.push(p('O **Lote 1** é indivisível: os módulos operam sobre base de dados única, e o fracionamento entre software, implantação, formação e suporte criaria zona cinzenta de responsabilização e perda do efeito de coordenação — que é o próprio objeto. Os itens são cotados individualmente dentro do lote, preservando a transparência dos preços unitários.'));
    c.push(p('A separação dos demais lotes amplia a competitividade e mantém rubricas apartadas: '+(ctx.cams.has('B')?'a avaliação pedagógica é complementar e pode ser prestada por fornecedor distinto, desde que seus resultados sejam importados pela plataforma; ':'')+(ctx.cams.has('C')?'o atendimento multiprofissional é serviço de saúde, custeado pela Saúde, sem comprometer o limite de manutenção e desenvolvimento do ensino.':'')));
  } else {
    c.push(p('Propõe-se **lote único**: indivisibilidade funcional (base de dados única), unicidade de responsabilidade sobre o resultado, formação vinculada ao produto e economia de escala na gestão contratual. A adoção de lote único não restringe a competitividade, admitida a subcontratação de parcelas acessórias nos limites do art. 122 da Lei nº 14.133/2021. Os itens são cotados individualmente dentro do lote.'));
    if(ctx.cams.has('C') && (ctx.cams.has('A')||ctx.cams.has('B'))) c.push(p('Havendo parcela de saúde no mesmo lote, a despesa deverá ser empenhada em rubricas distintas (Educação e Saúde), conforme a natureza de cada item.'));
  }

  c.push(...h1('IX — Resultados pretendidos','(art. 18, § 1º, IX)'));
  const res = [['Universalização do PEI','% de educandos do público-alvo com PEI ativo e homologado','≥ 95% ao final do 1º ano de operação plena'],['Tempestividade','Tempo médio entre identificação e homologação do PEI','≤ 30 dias'],['Revisão periódica efetiva','% de PEI revisados no prazo definido pela Secretaria','≥ 90%'],['Integração intersetorial','% de PEI com ao menos um registro da saúde no período','≥ 60% ao final do 2º ano'],['Participação da família','% de famílias com registro no semestre','≥ 70%'],['Capacitação','% de profissionais capacitados e certificados','≥ 90%'],['Evidência para gestão e controle','Relatórios gerenciais emitidos pela plataforma','Mensal e consolidado anual']];
  if(ctx.cams.has('C')) res.push(['Acesso ao cuidado','% dos educandos com pacote clínico atendidos na periodicidade do plano','≥ 85%']);
  c.push(tbl(['Resultado','Indicador','Meta de referência'], res, [28,44,28], {zebra:true}));

  c.push(...h1('X — Providências prévias à celebração do contrato','(art. 18, § 1º, X)'));
  ['Designação formal de gestor e de fiscais do contrato (art. 117 da Lei nº 14.133/2021).','Consolidação da base de educandos do público-alvo a partir do Censo Escolar e dos registros da rede.','Definição da base legal de tratamento de dados e do fluxo de consentimento parental (art. 14 da LGPD), com validação do encarregado.','Formalização da matriz controlador-operador e do acordo de tratamento de dados.','Mapeamento dos sistemas legados a integrar e disponibilização da documentação técnica.','Verificação da conectividade das unidades escolares, com plano de contingência.','Articulação formal com a Saúde para o acesso dos profissionais ao Módulo Saúde'+(ctx.cams.has('C')?' e para a regulação do atendimento multiprofissional':'')+'.','Publicação de ato normativo disciplinando o uso da plataforma e os prazos do PEI.','Reserva orçamentária e compatibilidade com o PPA, a LDO e a LOA.']
    .concat(ctx.modo==='srpcons'?['Procedimento de Intenção de Registro de Preços (IRP) junto aos municípios consorciados, com manifestação das quantidades de cada participante.']:[])
    .concat(ctx.modo==='adesao'?['Consulta e aceite prévios do órgão gerenciador e do fornecedor beneficiário da ata (art. 86, § 2º).']:[])
    .forEach(t=>c.push(bl(t)));

  c.push(...h1('XI — Contratações correlatas e interdependentes','(art. 18, § 1º, XI)'));
  c.push(tbl(['Contratação','Relação','Providência'],[['Sistema de gestão escolar','Interdependente — fonte de matrícula, turma e frequência','Prever integração e obter documentação da API ou layout de exportação.'],['Conectividade das unidades','Correlata — condição de uso','Verificar contrato vigente e prever contingência.'],['Equipamentos de acesso','Correlata','Verificar suficiência do parque; a solução web dispensa aquisição dedicada.'],['Sistemas de informação em saúde','Interdependente','Articular com a Saúde o fluxo de dados e perfis de acesso.'],['Formação continuada já contratada','Potencialmente sobreposta','Verificar sobreposição para evitar duplicidade de despesa.']],[28,32,40],{zebra:true}));

  c.push(...h1('XII — Possíveis impactos ambientais e medidas mitigadoras','(art. 18, § 1º, XII)'));
  c.push(p('Serviço em nuvem, sem fornecimento de bens físicos, sem resíduos sólidos ou passivo ambiental relevante. O consumo energético de data center é mitigado pela preferência por provedores com certificação de eficiência; a digitalização do PEI elimina impressão e arquivamento em papel; a solução opera sobre o parque existente, sem descarte eletrônico adicional (art. 11, IV, da Lei nº 14.133/2021).'));

  c.push(...h1('XIII — Posicionamento conclusivo','(art. 18, § 1º, XIII)'));
  c.push(p('Considerando a necessidade decorrente de obrigação legal expressa (item I); o levantamento de mercado e o descarte fundamentado das alternativas (item V); a pluralidade de fornecedores aptos; e a mensurabilidade dos resultados pretendidos (item IX) —'));
  const recModo = {srp:'a adoção do Sistema de Registro de Preços, por pregão eletrônico, nos termos dos arts. 82 e seguintes da Lei nº 14.133/2021, dada a conveniência de contratação por demanda',
    pregao:'a contratação por pregão eletrônico, com contrato de serviço contínuo de 12 meses prorrogável até 5 anos (arts. 106 e 107 da Lei nº 14.133/2021)',
    srpcons:'o Registro de Preços compartilhado, conduzido por '+(S.consNome.trim()||'consórcio público')+' como órgão gerenciador, com os municípios consorciados como participantes (art. 86 da Lei nº 14.133/2021 e Lei nº 11.107/2005)',
    adesao:'a adesão a Ata de Registro de Preços vigente, observadas as condições e os limites do art. 86, §§ 2º a 5º, da Lei nº 14.133/2021, inclusive quanto à esfera do órgão gerenciador',
    direta:'a contratação direta de piloto, com fundamento na hipótese legal a ser definida pela assessoria jurídica (arts. 74 ou 75), instruída na forma do art. 72 da Lei nº 14.133/2021'}[ctx.modo];
  c.push(box('Conclusão',['Conclui-se pela viabilidade técnica e pela adequação da contratação, no valor estimado de **'+brl(ctx.total)+'** para '+ctx.vig+' meses. Recomenda-se: (a) '+recModo+'; (b) implantação faseada, com validação no primeiro lote antes da expansão; e (c) o prosseguimento do feito com a elaboração do Termo de Referência, do qual este ETP e seu Anexo I são parte integrante.'],'EEF0F8'));
  c.push(...K.assinaturas());
  const doc = K.documento('Estudo Técnico Preliminar — '+ctx.label, c);
  return K.X.Packer.toBlob(doc);
}

/* ---------------- TR ---------------- */
async function gerarTR(ctx){
  await carregar('docx');
  const K = docKit(ctx,'TR'), {p,h1,h2,bl,tbl,box,gap} = K, {d, L} = ctx;
  const c = [], srp = ['srp','srpcons'].includes(ctx.modo);
  const prefixo = {srp:'Registro de preços para eventual contratação de ', srpcons:'Registro de preços compartilhado para eventual contratação de ', adesao:'Contratação, por adesão a Ata de Registro de Preços, de ', pregao:'Contratação de ', direta:'Contratação direta, em caráter de piloto, de '}[ctx.modo];
  const objeto = prefixo+'solução integrada de apoio à inclusão escolar e à coordenação do cuidado de educandos com TEA e demais público-alvo da educação especial, compreendendo '+ctx.partes.join('; ');
  c.push(...K.capa('TERMO DE REFERÊNCIA  ·  MINUTA','Termo de Referência — TR', objeto, paresCapa(ctx,'Art. 6º, XXIII, alíneas "a" a "j", da Lei nº 14.133, de 1º de abril de 2021',[['DOCUMENTOS VINCULADOS','Estudo Técnico Preliminar e Anexo I — Memória de cálculo (Ref. '+ctx.ref+')'],['VALOR ESTIMADO',brl(ctx.total)+' para '+ctx.vig+' meses']])));
  c.push(box('Nota de uso deste documento',['Minuta estruturante para adequação pela área técnica do órgão demandante. Os requisitos foram formulados por função e desempenho, e não por características exclusivas de fornecedor (arts. 40, I, e 25, § 1º, da Lei nº 14.133/2021). Antes da publicação, a área técnica deve confirmar que cada requisito é atendível por mais de um fornecedor e submeter a minuta à assessoria jurídica (art. 53).', cadeiaTxt(ctx,'TR')]));

  c.push(...h1('1. Do objeto','(art. 6º, XXIII, "a")'), h2('1.1. Objeto'));
  c.push(p(objeto+', conforme condições, quantidades e exigências deste Termo de Referência e de seus anexos.'));
  c.push(h2('1.2. Natureza do objeto'));
  L.forEach(lt=>c.push(bl('**Lote '+lt.n+'** — '+lt.natServ+(lt.rub!=='Saúde'?', passível de descrição objetiva por especificações usuais de mercado (art. 6º, XIII, da Lei nº 14.133/2021)':'')+'.')));
  c.push(h2('1.3. Itens e quantitativos'));
  const fq = x => num(x.rec ? x.q*ctx.vig/12 : x.q, 1);
  c.push(tbl(['#','Descrição do item','Unidade','Qtd. ano 1','Qtd. vigência','CATMAT/CATSER'], tabelaItens(K, L, [{v:x=>String(x.n)},{v:x=>x.desc},{v:x=>x.un},{v:x=>({__c:1,t:num(x.q),right:true})},{v:x=>({__c:1,t:fq(x),right:true})},{v:()=>'[   ]'}]), [5,40,13,11,12,19]));
  c.push(p('Quantidades estimadas a partir do Censo 2022 e dos parâmetros do Anexo I do ETP'+(srp?'; em Registro de Preços, representam estimativa máxima, sem obrigação de contratação integral':'')+'. Itens de licença medidos em unidade-ano: a coluna "Qtd. vigência" já considera '+ctx.vig+' meses.',{size:18, before:60}));
  c.push(h2('1.4. Prazo de vigência'));
  if(srp) c.push(bl('Ata de Registro de Preços: 1 (um) ano, contado da divulgação no PNCP, prorrogável por igual período mediante comprovação de preço vantajoso (art. 84 da Lei nº 14.133/2021).'));
  c.push(bl('Contrato: '+ctx.vig+' meses'+(ctx.modo==='direta'?', em caráter de piloto, sem prorrogação automática':', prorrogável até o limite de 5 (cinco) anos, por se tratar de serviço contínuo (arts. 106 e 107), mediante vantajosidade e disponibilidade orçamentária')+'.'));
  c.push(h2('1.5. Garantia'), p('O contratado garantirá o pleno funcionamento da solução durante toda a vigência, incluindo correção de defeitos, atualizações e manutenção dos níveis de serviço, sem custo adicional. Aos serviços de implantação, integração e formação aplica-se garantia de 90 dias do respectivo termo de aceite.'));
  if(ctx.part.length){ c.push(h2('1.6. Órgão gerenciador e participantes'), p('Órgão gerenciador: **'+(S.consNome.trim()||PRE('consórcio público'))+'**. Participantes e estimativa de demanda:'));
    c.push(tbl(['Município participante','Código IBGE','Educandos acompanhados','Valor anual estimado'], ctx.part.map(x=>[x.nome,x.cod,int(x.acomp),brl(x.ano)]),[40,18,20,22],{right:[2,3]})); }

  c.push(...h1('2. Da fundamentação da contratação','(art. 6º, XXIII, "b")'));
  c.push(p('A contratação fundamenta-se no dever legal de garantir acompanhamento integral aos educandos com transtornos de aprendizagem e ao público-alvo da educação especial (Lei nº 14.254/2021; Lei nº 13.146/2015, art. 28; Lei nº 12.764/2012; Lei nº 9.394/1996, arts. 58 a 60; Decreto nº 7.611/2011; Resolução CNE/CEB nº 4/2009).'));
  c.push(p('O ETP demonstrou que o modelo atual não atende àquelas obrigações e dimensionou o público em **'+int(d.acomp)+' educandos** ('+ctx.label+', Censo 2022, '+ctx.recorteTxt+'). Foram descartadas fundamentadamente a manutenção da situação atual, o desenvolvimento interno, a ampliação do sistema de gestão escolar e a aquisição de licença isolada, sem serviço. A contratação consta do PCA sob o item '+(S.doc.pca.trim()||PRE())+'.'));

  c.push(...h1('3. Da descrição da solução como um todo','(art. 6º, XXIII, "c")'));
  const comp=[['Núcleo de integração e inteligência','Base única por educando; base estruturada de conhecimento pedagógico, clínico e legal; IA para recomendação de estratégias e geração assistida do plano.'],['Módulo Escola','Ciclo completo do PEI com homologação humana e controle de prazos legais.'],['Módulo Saúde','Acesso controlado de profissionais de saúde, com condutas, metas terapêuticas e devolutivas.'],['Módulo Família','Observações do cotidiano, evolução e orientação domiciliar.'],['Painel de governança','Indicadores agregados para a Secretaria e prestação de contas.'],['Serviços','Implantação faseada, integração, formação certificada por perfil, supervisão multiprofissional e suporte por níveis de serviço.']];
  if(ctx.cams.has('B')) comp.push(['Avaliação pedagógica','Avaliação gamificada de habilidades, com laboratórios itinerantes e resultados importados para o PEI, sem finalidade diagnóstica.']);
  if(ctx.cams.has('C')) comp.push(['Atendimento multiprofissional','Consultas e sessões por profissionais habilitados, reguladas a partir do plano e registradas no módulo de saúde.']);
  c.push(tbl(['Componente','Entrega'], comp, [28,72], {zebra:true}));

  c.push(...h1('4. Dos requisitos da contratação','(art. 6º, XXIII, "d")'), h2('4.1. Requisitos funcionais mínimos'), p('Todos verificáveis em prova de conceito:'));
  const RF=[['Ficha única do educando, com dados escolares, clínicos e familiares e histórico preservado nas transferências.','Cadastro, transferência e recuperação de histórico.'],['Registro estruturado de anamnese pedagógica, avaliação de habilidades e observações, com anexos.','Preenchimento e anexação.'],['Recepção de laudos, relatórios e condutas terapêuticas, com identificação do profissional e data.','Registro por perfil de saúde.'],['Geração assistida por IA de proposta de PEI: objetivos, metas, estratégias, recursos, responsáveis e prazos.','Geração a partir de caso-teste do órgão.'],['Homologação humana obrigatória: a proposta não produz efeitos antes da validação por profissional identificado.','Proposta em estado não homologado até a validação.'],['Explicabilidade das recomendações em linguagem acessível.','Demonstração em caso-teste.'],['Controle de versões do PEI, com autor, data e comparação.','Versionamento e comparação.'],['Controle de prazos legais e alertas de revisão, parametrizáveis.','Disparo de alerta em demonstração.'],['Registro de execução por professores, terapeutas e responsáveis, com objetivos compartilhados.','Registro por cada perfil.'],['Mensuração da evolução por indicadores, com série histórica e gráfico.','Evolução com dois ciclos.'],['Canal da família em interface distinta e simplificada.','Acesso familiar.'],['Relatórios por educando, turma, unidade e rede, em PDF e formato estruturado.','Emissão e exportação.'],['Painel de governança: cobertura, tempestividade, revisão no prazo, participação da saúde e da família.','Demonstração do painel.'],['Perfis de acesso com segregação de informação sensível.','Perfil docente sem acesso a dado clínico restrito.'],['Trilha de auditoria imutável, exportável.','Trilha de período determinado.'],['Gestão do consentimento parental, com bloqueio automático na revogação.','Ciclo completo de consentimento.'],['Recepção de resultados de instrumentos de triagem ou avaliação de terceiros como insumo do PEI.','Importação de arquivo-modelo.'],['Biblioteca de estratégias pedagógicas e recursos de acessibilidade vinculada às necessidades.','Busca e vinculação ao plano.'],['Exportação integral dos dados em formato aberto, sem custo.','Exportação completa.']];
  c.push(tbl(['RF','Requisito funcional','Verificação'], RF.map((r,i)=>['RF-'+String(i+1).padStart(2,'0'),r[0],r[1]]), [9,62,29], {zebra:true}));
  c.push(h2('4.2. Requisitos não funcionais'));
  c.push(tbl(['RNF','Requisito','Parâmetro'],[['RNF-01','Fornecimento','SaaS em nuvem, sem servidores da Administração.'],['RNF-02','Compatibilidade','Web responsiva; duas últimas versões dos navegadores; desktop, tablet e smartphone.'],['RNF-03','Localização dos dados','Hospedagem e processamento em território nacional.'],['RNF-04','Disponibilidade','Mínimo de 99,5% mensais.'],['RNF-05','Desempenho','Até 3 s para 95% das consultas.'],['RNF-06','Capacidade','Mínimo de '+int(Math.max(50,Math.round(d.acomp*0.15)))+' usuários simultâneos (estimativa: 15% do público), com escalabilidade elástica.'],['RNF-07','Segurança','TLS 1.2 ou superior; criptografia em repouso.'],['RNF-08','Autenticação','SSO (SAML 2.0/OIDC); duplo fator para administradores.'],['RNF-09','Interoperabilidade','API REST OpenAPI 3.x; Educacenso/INEP; HL7 FHIR R4 desejável.'],['RNF-10','Acessibilidade','eMAG e WCAG 2.1 AA, com laudo.'],['RNF-11','Backup e continuidade','Backup diário, retenção de 90 dias, RPO ≤ 24 h, RTO ≤ 8 h.'],['RNF-12','Proteção de dados','LGPD; RIPD na implantação; acordo de tratamento; eliminação comprovada ao término.'],['RNF-13','Governança de IA','Explicabilidade; revisão humana; vedação ao uso dos dados para treinar modelos de terceiros.'],['RNF-14','Reversibilidade','Exportação integral e plano de transição sem custo.'],['RNF-15','Registro','Logs de auditoria com retenção mínima de 12 meses.']],[11,24,65],{zebra:true}));
  c.push(h2('4.3. Requisitos de habilitação técnica'));
  ['Atestado(s) de capacidade técnica de solução compatível, em quantitativo não inferior a '+PRE('20% a 30%')+' do total licitado.','Declaração de titularidade do software ou do direito de comercialização e suporte.','Equipe técnica mínima: responsável pedagógico (Pedagogia, Psicopedagogia ou Educação Especial), responsável técnico de TI, encarregado de dados e gerente de projeto.','Laudo de conformidade de acessibilidade (eMAG / WCAG 2.1 AA).','Aprovação em prova de conceito (item 8).'].forEach(t=>c.push(bl(t)));
  let sec=4;
  if(ctx.cams.has('B')){ c.push(h2('4.'+(sec++)+'. Requisitos específicos — avaliação pedagógica')); ['Avaliação de leitura, escrita, numeramento e atenção, com resultado de natureza pedagógica.','Vedado emitir hipótese diagnóstica ou nomear condição de saúde; o resultado sinaliza aprofundamento e encaminha ao profissional habilitado.','Laboratórios itinerantes com equipamentos e mediação, na razão de um para cada '+num(P.escLab)+' unidades escolares.','Consentimento parental prévio e integração dos resultados ao PEI (RF-17).'].forEach(t=>c.push(bl(t))); }
  if(ctx.cams.has('C')){ c.push(h2('4.'+(sec++)+'. Requisitos específicos — atendimento multiprofissional')); ['Profissionais com registro ativo no conselho de classe (CRM com RQE em neurologia ou psiquiatria; CRFa; CRP) e formação comprovada em psicopedagogia.','Estabelecimentos no CNES; teleatendimento admitido nos termos da Lei nº 14.510/2022, quando clinicamente adequado.','Atendimento a partir de encaminhamento regulado, vinculado ao PEI do educando.','Registro de cada atendimento no módulo de saúde, com devolutiva à escola.','Periodicidade de referência: '+PAC.map(x=>x.s.toLowerCase()+' '+x.n+'/ano').join('; ')+'; estratificável por nível de suporte.'].forEach(t=>c.push(bl(t))); }
  c.push(h2('4.'+sec+'. Vedações'));
  ['Uso dos dados para finalidade diversa da execução, inclusive treinamento de IA para terceiros, salvo anonimização irreversível e autorização expressa.','Transferência internacional de dados sem autorização e sem observância dos arts. 33 a 36 da LGPD.','Subcontratação do licenciamento e da operação da plataforma (parcela principal); parcelas acessórias até '+PRE('%')+', com autorização prévia (art. 122).','Cobrança de qualquer valor de famílias, educandos ou servidores.','Fornecimento de medicação pelo contrato.'].forEach(t=>c.push(bl(t)));

  c.push(...h1('5. Do modelo de execução do objeto','(art. 6º, XXIII, "e")'));
  c.push(p('Regime de empreitada por preço unitário, iniciando-se com a ordem de serviço e desenvolvendo-se em fases:'));
  const fases=[['1 — Planejamento','Reunião inicial; plano de implantação; matriz LGPD; RIPD; cronograma físico.','Até 30 dias da ordem de serviço'],['2 — Provisionamento','Ambientes de produção e homologação; SSO; perfis.','Até 15 dias do aceite da fase 1'],['3 — Integração','Integração com sistemas legados; carga inicial validada.','Até 45 dias do aceite da fase 2'],['4 — Implantação por lote','Parametrização por unidade; habilitação de usuários; termo de aceite.','Lotes de até '+PRE()+' unidades'],['5 — Capacitação','Turmas de até 20 por perfil, com avaliação e certificação.','Concomitante à fase 4'],['6 — Operação assistida','Ajustes, suporte reforçado e mentoria de casos.','90 dias após cada lote']];
  if(ctx.cams.has('B')) fases.push(['7 — Avaliação pedagógica','Aplicação com laboratórios itinerantes; importação dos resultados ao PEI.','Ciclos semestrais']);
  if(ctx.cams.has('C')) fases.push([(ctx.cams.has('B')?'8':'7')+' — Atendimento multiprofissional','Início dos atendimentos regulados a partir dos PEIs homologados.','A partir do 3º mês']);
  fases.push(['Operação plena','Uso regular; SLA; relatórios mensais.','Até o término da vigência'],['Encerramento','Exportação integral; transferência de conhecimento; eliminação segura dos dados.','Últimos 60 dias']);
  c.push(tbl(['Fase','Atividades','Prazo'], fases, [24,50,26], {zebra:true}));
  c.push(h2('5.2. Formação continuada'), p('Turmas de até 20 participantes, com carga mínima de 20 horas; trilhas para professor regente, AEE, coordenação, gestão, saúde e orientação às famílias; conteúdo mínimo: marco legal da inclusão, fundamentos do PEI, operação da plataforma, uso da informação clínica no planejamento, proteção de dados e uso responsável de IA; certificado individual.'));
  c.push(h2('5.3. Suporte e níveis de serviço'));
  c.push(tbl(['Severidade','Caracterização','Resposta','Solução'],[['Crítica','Indisponibilidade total ou exposição de dados','30 min','4 h'],['Alta','Funcionalidade essencial sem contorno','2 h','8 h úteis'],['Média','Erro com contorno disponível','4 h úteis','3 dias úteis'],['Baixa','Dúvida ou melhoria','8 h úteis','10 dias úteis']],[16,48,16,20],{zebra:true}));
  c.push(p('Incidentes de segurança com dados pessoais devem ser comunicados em até 2 horas, com relatório preliminar em 24 horas (art. 48 da LGPD).',{before:80}));

  c.push(...h1('6. Do modelo de gestão do contrato','(art. 6º, XXIII, "f")'));
  c.push(tbl(['Agente','Atribuições'],[['Gestor do contrato','Coordenação da execução, alterações, relatórios de fiscalização e processos sancionatórios.'],['Fiscal técnico','Conformidade dos entregáveis, apuração dos níveis de serviço, termos de aceite.'],['Fiscal administrativo','Regularidade fiscal, trabalhista e previdenciária; conferência das notas fiscais.'],['Fiscal setorial','Verificação local da implantação e da capacitação.'],...(ctx.cams.has('C')?[['Fiscal da Saúde','Verificação dos atendimentos realizados e da regularidade dos profissionais.']]:[]),['Encarregado de dados','Obrigações de proteção de dados, RIPD e incidentes.'],['Preposto do contratado','Representação perante a Administração; resposta em até 2 dias úteis.']],[26,74],{zebra:true}));
  c.push(p('Relatório mensal de serviços até o 5º dia útil; reuniões mensais nas fases de implantação e trimestrais na operação plena; recebimento provisório em até 5 dias úteis e definitivo em até 10 dias úteis.',{before:100}));
  c.push(h2('6.2. Matriz de riscos'));
  c.push(tbl(['Risco','Mitigação','Alocação'],[['Baixa adesão dos profissionais','Capacitação obrigatória; operação assistida; indicador de uso.','Compartilhado'],['Conectividade insuficiente','Levantamento prévio; contingência; priorização de unidades.','Contratante'],['Falha de integração','Prova de integração na fase 3; exportação manual.','Compartilhado'],['Incidente de segurança','Criptografia; auditoria; RIPD; comunicação em 2 h.','Contratado'],['Recomendação de IA inadequada','Homologação humana; explicabilidade; contestação.','Compartilhado'],['Variação da demanda',srp?'Registro de Preços, sem obrigação de consumo integral.':'Medição por quantidade efetivamente prestada.','Contratante']],[30,50,20],{zebra:true}));

  c.push(...h1('7. Dos critérios de medição e pagamento','(art. 6º, XXIII, "g")'));
  c.push(p('O serviço é medido e pago pelo que foi efetivamente prestado no período:'));
  // agrupa itens consecutivos com o mesmo critério (ex.: as licenças 1 a 4)
  const grp=[]; L.forEach(lt=>lt.itens.forEach(x=>{ const crit=MEDICAO[x.id]==='idem'?MEDICAO.esc:MEDICAO[x.id], g=grp[grp.length-1];
    if(g && g.crit===crit && g.b===x.n-1){ g.b=x.n; g.nomes.push(x.desc); } else grp.push({a:x.n, b:x.n, crit, id:x.id, nomes:[x.desc]}); }));
  const rotulo = g => g.id==='cli' ? 'Atendimentos multiprofissionais' : g.nomes.length>1 && /^Licença/.test(g.nomes[0]) ? 'Licenças de uso' : g.nomes.join('; ');
  c.push(tbl(['#','Item','Critério de medição'], grp.map(g=>[g.a===g.b?String(g.a):g.a+' a '+g.b, rotulo(g), g.crit]), [9,39,52], {zebra:true}));
  c.push(h2('7.2. Glosas por nível de serviço'), p('Aplicáveis cumulativamente, limitadas a 30% do valor mensal: disponibilidade entre 98% e 99,5% — 1% por ponto; abaixo de 98% — 3% por ponto; atraso de solução crítica — 0,5% por hora; alta — 0,5% por dia útil; média/baixa — 0,2% por dia útil; atraso do relatório mensal — 0,2% por dia útil; atraso na comunicação de incidente — 1% por hora.'));
  c.push(h2('7.3. Pagamento e reajuste'), p('Pagamento em até 30 dias do recebimento definitivo, sem pagamento antecipado nem por licença não ativada. Reajuste após 12 meses da data-base da proposta, pelo '+PRE('IPCA/IBGE ou índice setorial')+' (arts. 25, § 7º, e 92, § 3º).'));

  c.push(...h1('8. Da forma e dos critérios de seleção do fornecedor','(art. 6º, XXIII, "h")'));
  const sel = {
    srp:['Pregão, na forma eletrônica, destinado a Registro de Preços, por se tratar de serviço comum, com modo de disputa aberto (arts. 6º, XLI, 29 e 82 e seguintes da Lei nº 14.133/2021).'],
    pregao:['Pregão, na forma eletrônica, por se tratar de serviço comum, com modo de disputa aberto (arts. 6º, XLI, e 29 da Lei nº 14.133/2021), para contrato de serviço contínuo.'],
    srpcons:['Pregão eletrônico para Registro de Preços compartilhado, conduzido por '+(S.consNome.trim()||PRE('consórcio'))+' como órgão gerenciador, precedido de Intenção de Registro de Preços junto aos municípios consorciados (art. 86 da Lei nº 14.133/2021 e Lei nº 11.107/2005). Cada participante contrata a sua parcela, no limite das quantidades informadas.'],
    adesao:['Adesão à Ata de Registro de Preços nº '+PRE()+', gerenciada por '+PRE('órgão gerenciador')+', condicionada: (I) a este ETP, que demonstra ganhos de eficiência, viabilidade e economicidade; (II) à demonstração de que os valores registrados são compatíveis com os de mercado (art. 23); e (III) à consulta e ao aceite prévios do órgão gerenciador e do fornecedor (art. 86, § 2º).','Observar os limites de quantitativo por órgão aderente e em conjunto (art. 86, §§ 4º e 5º) e as regras aplicáveis à esfera do órgão gerenciador.'],
    direta:['Contratação direta, com fundamento em '+PRE('hipótese legal: art. 74 (inexigibilidade) ou art. 75 (dispensa)')+', instruída com os documentos do art. 72 da Lei nº 14.133/2021 e validada pela assessoria jurídica.','A dispensa por valor (art. 75, II) está sujeita ao teto atualizado anualmente; a inexigibilidade exige demonstração de inviabilidade de competição. O piloto destina-se a gerar evidência de resultado e atestado de capacidade técnica.']}[ctx.modo];
  c.push(h2('8.1. Modalidade'), ...sel.map(t=>p(t)));
  if(!['adesao','direta'].includes(ctx.modo)){
    c.push(h2('8.2. Critério de julgamento'), p('Menor preço '+(L.length>1?'global por lote':'global do lote')+', justificado pela indivisibilidade funcional demonstrada no ETP. Alternativamente, técnica e preço (art. 36), com peso de 30% para a técnica, mediante justificativa nos autos.'));
    c.push(h2('8.3. Prova de conceito'), p('A licitante provisoriamente classificada em primeiro lugar será convocada para prova de conceito em até 10 dias úteis, com verificação binária dos requisitos RF-01 a RF-19; exige-se 100% dos essenciais e 80% dos desejáveis.'));
  }
  if(ctx.cams.has('C')) c.push(box('Atenção — lote de atendimento multiprofissional',['Para a parcela de saúde, avaliar a adequação do credenciamento (art. 79 da Lei nº 14.133/2021), que admite múltiplos prestadores a preço fixado, e observar as regras de participação complementar da iniciativa privada no SUS (art. 199, § 1º, da Constituição e arts. 24 a 26 da Lei nº 8.080/1990).']));
  c.push(h2('8.'+(['adesao','direta'].includes(ctx.modo)?2:4)+'. Sanções'), p('Arts. 155 a 163 da Lei nº 14.133/2021, com gradação proporcional e consideração das glosas já aplicadas.'));

  c.push(...h1('9. Das estimativas do valor da contratação','(art. 6º, XXIII, "i")'));
  c.push(p('O valor total estimado é de **'+brl(ctx.total)+'** para a vigência de '+ctx.vig+' meses ('+brl(ctx.ano)+' no primeiro ano), conforme a planilha de formação de preços do item VI do ETP e o Anexo I. A pesquisa de preços foi realizada nos termos do art. 23 mediante '+PRE('fontes, datas e responsável')+'.'));
  c.push(tbl(['Lote','Rubrica','Primeiro ano','Vigência ('+ctx.vig+' meses)'], L.map(lt=>['Lote '+lt.n+' — '+lt.nome, lt.rub, brl(lt.ano), brl(lt.vig)]).concat([[{__c:1,t:'Total',bold:true},'',{__c:1,t:brl(ctx.ano),bold:true,right:true},{__c:1,t:brl(ctx.total),bold:true,right:true}]]),[46,14,20,20],{right:[2,3]}));

  c.push(...h1('10. Da adequação orçamentária','(art. 6º, XXIII, "j")'));
  if(srp) c.push(p('Tratando-se de Registro de Preços, a indicação da dotação orçamentária somente será exigida para a formalização do contrato ou instrumento equivalente.'));
  c.push(tbl(['Lote','Unidade orçamentária','Natureza da despesa','Fonte de recursos'], L.map(lt=>['Lote '+lt.n+' ('+lt.rub+')', S.doc.uo.trim()||PRE(), lt.nat, lt.rub==='Saúde'?'Bloco de custeio da Saúde; recursos próprios; emendas':'FUNDEB; Salário-Educação; MDE; FNDE; emendas']),[18,22,32,28]));
  c.push(p('As matrículas da educação especial em AEE admitem duplo cômputo no FUNDEB (Lei nº 14.113/2020; Decreto nº 7.611/2011). A separação entre Educação e Saúde permite custear a parcela clínica sem comprometer o limite de manutenção e desenvolvimento do ensino.',{before:100}));

  c.push(...h1('11. Das disposições finais'));
  c.push(p('Integram este Termo, independentemente de transcrição: o Estudo Técnico Preliminar; o **Anexo I — Memória de cálculo** (Ref. '+ctx.ref+'); o roteiro da prova de conceito; o modelo de acordo de tratamento de dados; o modelo de termo de aceite'+(srp?'; e a minuta da Ata de Registro de Preços':'')+'. A titularidade dos dados é exclusiva da Administração e dos titulares. Esta minuta deve ser adequada aos dados e à regulamentação locais e submetida à assessoria jurídica (art. 53 da Lei nº 14.133/2021).'));
  c.push(...K.assinaturas());
  const doc = K.documento('Termo de Referência — '+ctx.label, c);
  return K.X.Packer.toBlob(doc);
}

/* =====================================================================
   UI da seção de documentos
   ===================================================================== */
function renderDocs(){
  const el = $('docModo'); if(!el || !LAST) return;
  const ctx = {label:enteLabel(), modo:docModo(), dem:docDem()};
  $('docEnte').textContent = ctx.label; $('docEsfera').textContent = esferaTxt();
  $('docOrgao').textContent = orgaoDemandante();
  $('docDem').value = ctx.dem;
  el.innerHTML = Object.entries(MODOS).map(([k,m])=>`<label class="opt ${ctx.modo===k?'on':''}"><input type="radio" name="docmodo" ${ctx.modo===k?'checked':''} onchange="S.doc.modo='${k}';update()"><div><b>${m.t}</b><span>${m.s}</span></div></label>`).join('');
  $('docLotes').value = S.doc.lotes;
  const L = montarLotes(LAST);
  $('docLotesTb').innerHTML = '<thead><tr><th>Lote</th><th>Rubrica</th><th class="n">Itens</th><th class="n">1º ano</th><th class="n">Vigência</th></tr></thead><tbody>'+
    L.map(x=>`<tr><td><b>Lote ${x.n}</b><small class="et">${x.nome}</small></td><td>${x.rub}</td><td class="n">${x.itens.length}</td><td class="n">${brlK(x.ano)}</td><td class="n">${brlK(x.vig)}</td></tr>`).join('')+
    `<tr class="tot"><td colspan="3">Total · ${S.vig} meses</td><td class="n">${brlK(L.reduce((a,x)=>a+x.ano,0))}</td><td class="n">${brlK(L.reduce((a,x)=>a+x.vig,0))}</td></tr></tbody>`;
  $('docWarn').style.display = S.doc.g3 ? '' : 'none';
}
async function baixarDoc(tipo){
  const st=$('docStatus'), btns=document.querySelectorAll('.docbtn');
  if(!LAST || !LAST.d.acomp){ st.style.color='#B42318'; st.textContent='Selecione um ente com público no recorte etário.'; return; }
  btns.forEach(b=>b.disabled=true); st.style.color='#14225A';
  try{
    const ctx = ctxDoc();
    const nomes = {etp:'1_ETP_Inclusao_Escolar_TEA_'+ctx.slug+'.docx', xlsx:'2_Anexo_I_Memoria_de_Calculo_'+ctx.slug+'.xlsx', tr:'3_TR_Inclusao_Escolar_TEA_'+ctx.slug+'.docx', pdf:'4_Proposta_Referencia_'+ctx.slug+'.pdf'};
    const gera = {etp:()=>gerarETP(ctx), xlsx:()=>gerarXlsx(ctx), tr:()=>gerarTR(ctx), pdf:async()=>{ const r=montarPDF(); return r.doc.output('blob'); }};
    if(tipo==='zip'){
      st.textContent='Montando o pacote (ETP, memória de cálculo, TR e proposta)…';
      await carregar('zip'); const z = new JSZip();
      for(const k of ['etp','xlsx','tr','pdf']) z.file(nomes[k], await gera[k]());
      z.file('LEIA-ME.txt', 'Pacote documental da contratação — '+ctx.label+'\r\nReferência: '+ctx.ref+'\r\nGerado em '+new Date().toLocaleString('pt-BR')+' pelo painel Inclusi.Via (G3 Health Service)\r\n\r\n'+
        '1) ETP — Estudo Técnico Preliminar (art. 18 da Lei 14.133/2021): por que contratar, quanto e como.\r\n'+
        '2) Anexo I do ETP — Memória de cálculo com fórmulas vivas. Células amarelas são editáveis.\r\n'+
        '3) TR — Termo de Referência (art. 6º, XXIII): objeto, requisitos, execução, medição e pagamento.\r\n'+
        '4) Proposta comercial de referência — insumo da pesquisa de preços (art. 23, § 1º, IV).\r\n\r\n'+
        'Ente: '+ctx.label+' | Esfera: '+ctx.esfera+' | Órgão demandante: '+ctx.orgao+'\r\nTipo de contratação: '+MODOS[ctx.modo].t+'\r\n'+
        'Público acompanhado: '+int(ctx.d.acomp)+' educandos | Vigência: '+ctx.vig+' meses | Valor estimado: '+brl(ctx.total)+'\r\n\r\n'+
        'Campos [PREENCHER] dependem de dados locais. Minutas sujeitas à pesquisa de preços (art. 23) e à análise da assessoria jurídica (art. 53).\r\n');
      salvarBlob(await z.generateAsync({type:'blob'}), 'Pacote_Contratacao_TEA_'+ctx.slug+'.zip');
    } else {
      st.textContent = tipo==='xlsx' ? 'Gerando a memória de cálculo…' : 'Gerando o '+tipo.toUpperCase()+'…';
      salvarBlob(await gera[tipo](), nomes[tipo]);
    }
    st.style.color='#0F7B55'; st.textContent='✅ Documento gerado — confira a pasta de downloads.';
  }catch(e){ console.error(e); st.style.color='#B42318'; st.textContent='Não foi possível gerar: '+e.message; }
  finally{ btns.forEach(b=>b.disabled=false); }
}
