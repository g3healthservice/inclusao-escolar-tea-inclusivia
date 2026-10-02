/* ===== Inclusi.Via v2 — distribui as seções do painel em etapas (modo aplicativo) =====
   Roda depois do script principal: só move nós do DOM (ids e handlers continuam os mesmos)
   e acrescenta a barra de resumo. Toda a lógica de cálculo, PDF, ETP/TR e planilha é a do template. */
(function(){
  const $ = id => document.getElementById(id);
  const content = $('v2content');
  const ORDEM = ['ente','publico','mods','valor','proposta','docs','rank'];
  const NOMES = {ente:'Ente', publico:'Público', mods:'Módulos', valor:'Valor', proposta:'Proposta', docs:'ETP e TR', rank:'Ranking', sobre:'Sobre o programa', metodo:'Metodologia'};
  const SEC2PANE = {top:'sobre', problema:'sobre', servico:'sobre', publico:'ente', custos:'mods', valor:'valor', proposta:'proposta', documentos:'docs', ranking:'rank', metodo:'metodo'};
  const etapa = id => { const k = ORDEM.indexOf(id); return k < 0 ? '' : 'Etapa '+(k+1)+' de '+ORDEM.length; };

  function pane(id){ const p = document.createElement('div'); p.className = 'v2pane'; p.id = 'v2p-'+id; content.appendChild(p); return p; }
  function navHTML(id){
    const k = ORDEM.indexOf(id); if(k < 0) return '<div class="v2nav"><button class="btn ghost" data-go="ente">← Voltar às etapas</button><span></span></div>';
    const prev = ORDEM[k-1], next = ORDEM[k+1];
    return '<div class="v2nav">'+(prev?'<button class="btn ghost" data-go="'+prev+'">← '+NOMES[prev]+'</button>':'<span></span>')+
      (next?'<button class="btn sec" data-go="'+next+'">Próximo: '+NOMES[next].toLowerCase()+' →</button>':'<span></span>')+'</div>';
  }
  function etiqueta(sec, id){ const sh = sec.querySelector('.sh'); if(sh) sh.textContent = etapa(id)+' · '+sh.textContent.replace(/^\d+\s*·\s*/,''); }
  function addNav(p, id){ p.insertAdjacentHTML('beforeend', navHTML(id)); }

  // ---- Sobre o programa: abertura + problema + modelo de serviço
  const sobre = pane('sobre');
  sobre.append(document.querySelector('.hero'), $('problema'), $('servico'));
  sobre.insertAdjacentHTML('beforeend', '<div class="v2nav"><span></span><button class="btn" data-go="ente">Começar: escolher o ente →</button></div>');

  // ---- 1 Ente: busca, consórcio, estado + cartão do ente
  const ente = pane('ente');
  ente.innerHTML = '<div class="v2h"><div class="sh">'+etapa('ente')+' · Ente público</div><h2>Para quem é a proposta?</h2>'+
    '<p class="lead2">Escolha um município, monte um consórcio de vários municípios ou selecione um estado. Todas as etapas seguintes são calculadas para este ente.</p></div>'+
    '<div class="card" id="v2enteCard"></div><div class="v2stats" id="v2stats"></div>'+
    '<div class="note" id="v2statsNote" style="margin-top:10px"></div>';
  const pub = $('publico'), pcard = pub.querySelector('.card'), ecard = $('v2enteCard');
  [pcard.querySelector('.tabs'), $('r-busca'), $('chips'), $('r-uf'), pcard.querySelector('.ente'), $('capHint')].forEach(n => n && ecard.appendChild(n));
  ecard.insertAdjacentHTML('beforeend', '<div class="note" style="margin-top:10px">Não sabe por onde começar? Veja onde está o público no <button class="v2go" data-go="rank">ranking de oportunidade →</button></div>');
  addNav(ente, 'ente');

  // ---- 2 Público
  const p2 = pane('publico'); p2.appendChild(pub); etiqueta(pub, 'publico');
  const lead = pub.querySelector('.lead2'); if(lead) lead.innerHTML = lead.innerHTML.replace(/\s*Escolha um município, monte um consórcio ou selecione um estado\.?/, '');
  addNav(p2, 'publico');

  // ---- 3 a 7
  [['mods','custos'],['valor','valor'],['proposta','proposta'],['docs','documentos'],['rank','ranking']].forEach(([id, sid]) => {
    const p = pane(id), s = $(sid); p.appendChild(s); etiqueta(s, id); addNav(p, id);
  });
  const met = pane('metodo'); met.appendChild($('metodo')); addNav(met, 'metodo');
  const foot = document.querySelector('.foot'); if(foot) $('v2footer').appendChild(foot), foot.style.display = 'block', foot.style.padding = '0';

  // ---- navegação
  function go(id, opt){
    if(!$('v2p-'+id)) id = 'ente';
    document.querySelectorAll('.v2pane').forEach(p => p.classList.toggle('on', p.id === 'v2p-'+id));
    const k = ORDEM.indexOf(id);
    document.querySelectorAll('.v2st').forEach(b => { const j = ORDEM.indexOf(b.dataset.go); b.classList.toggle('on', j === k); b.classList.toggle('done', k >= 0 && j < k); });
    document.querySelectorAll('.v2link').forEach(b => b.classList.toggle('on', b.dataset.go === id));
    try{ sessionStorage.setItem('v2_pane', id); }catch(e){}
    if(!(opt && opt.keepScroll)) window.scrollTo(0, 0);
  }
  window.v2go = go;
  document.addEventListener('click', e => {
    const g = e.target.closest('[data-go]'); if(g){ e.preventDefault(); go(g.dataset.go); return; }
    const t = e.target.closest('[data-theme-set]'); if(t){ setTheme(t.dataset.themeSet); return; }
    const a = e.target.closest('a[href^="#"]'); if(a){ const id = a.getAttribute('href').slice(1); if(SEC2PANE[id]){ e.preventDefault(); go(SEC2PANE[id]); } }
  }, true);
  // loadMun/loadUF do ranking mudam o hash para #publico: leva ao ente e devolve o link direto (#m=…)
  window.addEventListener('hashchange', () => {
    const id = location.hash.slice(1); if(!SEC2PANE[id]) return;
    go(SEC2PANE[id]); history.replaceState(null, '', location.pathname + location.search); update();
  });

  // ---- tema
  function setTheme(t){
    t = t === 'dark' ? 'dark' : 'light';
    document.documentElement.dataset.theme = t;
    try{ localStorage.setItem('inclusivia_theme', t); }catch(e){}
    document.querySelectorAll('[data-theme-set]').forEach(b => b.classList.toggle('on', b.dataset.themeSet === t));
    const dk = t === 'dark', tick = dk ? '#A3ABD3' : '#59618A', grid = dk ? '#25305F' : '#EEF0F6', surf = dk ? '#131C42' : '#fff';
    Chart.defaults.color = tick;
    [chFaixa, chAnos].forEach(c => { if(!c) return; c.options.scales.y.grid.color = grid; c.options.scales.x.ticks.color = tick; c.options.scales.y.ticks.color = tick; });
    if(chFaixa){ chFaixa.data.datasets[0].borderColor = surf; chFaixa.data.datasets[1].backgroundColor = dk ? '#3A4470' : '#D3D7E5'; }
    if(chComp) chComp.data.datasets[0].borderColor = surf;
    [chFaixa, chComp, chAnos].forEach(c => c && c.update('none'));
  }
  window.setTheme = setTheme;

  // ---- barra de resumo, sempre sincronizada com o cálculo
  const IC = {
    u:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>',
    b:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2zM22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>',
    s:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m3 21 9-18 9 18M7 13h10"/></svg>',
    r:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8M12 18V6"/></svg>'
  };
  function sync(){
    const C = LAST; if(!C) return; const d = C.d, ce = C.ceAll, G = C.gasto;
    $('v2tNome').textContent = $('enteNome').textContent;
    $('v2tTipo').textContent = $('enteTipo').textContent;
    $('v2kA').textContent = int(d.acomp); $('v2kF').textContent = int(d.familias); $('v2kE').textContent = int(d.escolas);
    $('v2kV').textContent = brlK(C.total); $('v2kVl').textContent = 'Valor · '+S.vig+(S.vig === 1 ? ' mês' : ' meses');
    const st = [
      ['', IC.u, ce ? int(ce.tea) : '—', 'Estudantes com TEA (todas as redes)'],
      ['t', IC.b, ce ? int(ce.rede) : '—', 'Matrículas na educação especial da rede'],
      ['p', IC.s, int(d.escolas), S.base === 'censo' ? 'Escolas da rede com educação especial' : 'Unidades escolares (estimativa)'],
      ['o', IC.r, G && G.emp > 0 ? brlK(G.emp) : '—', 'Gasto atual com educação especial/ano']
    ];
    $('v2stats').innerHTML = st.map(x => '<div class="v2stat '+x[0]+'"><span class="ic">'+x[1]+'</span><small>'+x[3]+'</small><b>'+x[2]+'</b></div>').join('');
    $('v2statsNote').textContent = 'Fontes: INEP, Censo Escolar 2025 · Tesouro Nacional, SICONFI (subfunção 12.367). O detalhamento está nas etapas Público e Valor.';
  }
  const _update = window.update;
  window.update = function(){ const r = _update.apply(this, arguments); sync(); return r; };

  let ini = 'ente'; try{ ini = sessionStorage.getItem('v2_pane') || 'ente'; }catch(e){}
  setTheme(document.documentElement.dataset.theme);
  sync(); go(ini, {keepScroll:true});
})();
