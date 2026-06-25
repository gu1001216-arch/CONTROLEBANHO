// Lógica compartilhada do dashboard admin e do painel público.
// Define window.DASH_ENDPOINT antes de carregar este arquivo.
let chartTipo, chartProc, chartDia, chartPeso, chartArea;
const ESTADOS={
  PREPARANDO:{t:'Preparando',c:'#F59E0B',i:'hourglass-split'},
  PREENCHER:{t:'Preencher',c:'#64748b',i:'pencil-square'},
  FILA_BANHO:{t:'Na fila',c:'#7C5CFC',i:'list-ol'},
  EM_BANHO:{t:'Em banho',c:'#2563EB',i:'droplet-fill'},
};

function getRange(){
  const de=document.getElementById('fDe')?.value||'';
  const ate=document.getElementById('fAte')?.value||'';
  const p=new URLSearchParams();
  if(de)p.set('de',de); if(ate)p.set('ate',ate);
  return p.toString()?('?'+p.toString()):'';
}

async function atualizar(){
  const r=await fetch(window.DASH_ENDPOINT+getRange());
  const d=await r.json();
  set('kTotal',d.total); set('kAnd',d.em_andamento);
  set('kBanhoNormal',d.banho_normal!==undefined?d.banho_normal:d.normais);
  set('kBanhoRetrab',d.banho_retrabalho!==undefined?d.banho_retrabalho:d.retrabalhos);
  set('kNormais',d.normais); set('kRetrab',d.retrabalhos);
  set('kPrep',d.media_prep); set('kBanho',d.media_banho);
  set('kEspera',d.media_espera!==undefined?d.media_espera:0);
  set('kPecas',d.pecas_total_geral!==undefined?d.pecas_total_geral:0);
  set('kPeso',d.peso_total_geral!==undefined?d.peso_total_geral:0);
  set('kArea',d.area_total_geral!==undefined?d.area_total_geral:0);
  renderAtivos(d.ativos);
  renderCharts(d);
  if(document.getElementById('tbody')) renderTabela(d.registros);
  if(document.getElementById('histBody')) renderHistorico(d.registros);
}
function renderHistorico(regs){
  const tb=document.getElementById('histBody');
  if(!regs||!regs.length){tb.innerHTML='<tr><td colspan="11" class="empty">Nenhum cesto concluído ainda.</td></tr>';return;}
  tb.innerHTML=regs.map(r=>{
    // texto pesquisável: todas as OPs, códigos e descrições do cesto
    const itens=(r.itens&&r.itens.length)?r.itens:[{ordem:r.ordem,material:r.material,texto_breve:r.texto_breve}];
    const busca=(r.numero_cesto+' '+itens.map(it=>(it.ordem||'')+' '+(it.material||'')+' '+(it.texto_breve||'')).join(' ')+' '+(r.processo||'')+' '+(r.tipo||'')).toLowerCase();
    return `<tr class="${r.tipo==='Retrabalho'?'retrab':''}" data-busca="${busca.replace(/"/g,'')}">
    <td>${r.id}</td><td><strong>${r.numero_cesto}</strong></td><td>${r.n_itens>1?(r.ordem+' +'+(r.n_itens-1)):(r.ordem||'—')}</td>
    <td>${r.material||'—'}</td><td><span class="small">${r.texto_breve||'—'}</span></td><td>${r.qtd_total}</td>
    <td>${r.processo||'—'}</td><td><span class="pill ${r.tipo==='Retrabalho'?'pill-retrab':'pill-normal'}">${r.tipo}</span></td>
    <td class="mono">${r.prep_minutos}</td><td class="mono">${r.banho_minutos}</td>
    <td><span class="small">${r.banho_fim||''}</span></td>
  </tr>`;}).join('');
  if(typeof aplicarBuscaHistorico==='function')aplicarBuscaHistorico();
}
function set(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}

function renderAtivos(ativos){
  const c=document.getElementById('andamento'); if(!c)return;
  if(!ativos.length){c.innerHTML='<div class="empty"><i class="bi bi-clipboard-check"></i>Nenhum cesto em andamento.</div>';return;}
  c.innerHTML=ativos.map(a=>{
    const e=ESTADOS[a.estado]||{t:a.estado,c:'#888',i:'circle'};
    return `<div class="and-chip" style="border-left:4px solid ${e.c}">
      <div class="and-top"><span class="and-cesto">${a.numero_cesto}</span><i class="bi bi-${e.i}" style="color:${e.c}"></i></div>
      <div class="and-st" style="color:${e.c}">${e.t}</div>
      <div class="and-op">OP ${a.ordem||'—'}</div>
    </div>`;
  }).join('');
}

function renderCharts(d){
  const labels=Object.keys(d.por_processo), data=Object.values(d.por_processo);
  const diaL=Object.keys(d.por_dia), diaD=Object.values(d.por_dia);
  if(!chartTipo){
    chartTipo=new Chart(document.getElementById('chartTipo'),{type:'doughnut',
      data:{labels:['Normal','Retrabalho'],datasets:[{data:[d.normais,d.retrabalhos],backgroundColor:['#2BA45C','#1668C0'],borderWidth:0}]},
      options:{plugins:{legend:{position:'bottom'}},cutout:'66%'}});
    chartProc=new Chart(document.getElementById('chartProc'),{type:'bar',
      data:{labels,datasets:[{data,backgroundColor:'#1668C0',borderRadius:7}]},
      options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}}});
    if(document.getElementById('chartDia'))
      chartDia=new Chart(document.getElementById('chartDia'),{type:'line',
        data:{labels:diaL,datasets:[{data:diaD,borderColor:'#2BA45C',backgroundColor:'rgba(43,164,92,.15)',fill:true,tension:.3,borderWidth:3,pointRadius:4,pointBackgroundColor:'#2BA45C'}]},
        options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}}});
  }else{
    chartTipo.data.datasets[0].data=[d.normais,d.retrabalhos];chartTipo.update();
    chartProc.data.labels=labels;chartProc.data.datasets[0].data=data;chartProc.update();
    if(chartDia){chartDia.data.labels=diaL;chartDia.data.datasets[0].data=diaD;chartDia.update();}
  }
  // gráficos de peso e área por dia (linha) — cada um só se o canvas existir
  const pesoL=Object.keys(d.peso_por_dia||{}), pesoD=Object.values(d.peso_por_dia||{});
  const areaL=Object.keys(d.area_por_dia||{}), areaD=Object.values(d.area_por_dia||{});
  if(document.getElementById('chartPeso')){
    if(!chartPeso){
      chartPeso=new Chart(document.getElementById('chartPeso'),{type:'line',
        data:{labels:pesoL,datasets:[{label:'kg',data:pesoD,borderColor:'#2BA45C',backgroundColor:'rgba(43,164,92,.15)',fill:true,tension:.3,borderWidth:3,pointRadius:4,pointBackgroundColor:'#2BA45C'}]},
        options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}}});
    }else{chartPeso.data.labels=pesoL;chartPeso.data.datasets[0].data=pesoD;chartPeso.update();}
  }
  if(document.getElementById('chartArea')){
    if(!chartArea){
      chartArea=new Chart(document.getElementById('chartArea'),{type:'line',
        data:{labels:areaL,datasets:[{label:'m²',data:areaD,borderColor:'#1668C0',backgroundColor:'rgba(22,104,192,.15)',fill:true,tension:.3,borderWidth:3,pointRadius:4,pointBackgroundColor:'#1668C0'}]},
        options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}}});
    }else{chartArea.data.labels=areaL;chartArea.data.datasets[0].data=areaD;chartArea.update();}
  }
}

function renderTabela(regs){
  const tb=document.getElementById('tbody');
  if(!regs.length){tb.innerHTML='<tr><td colspan="11" class="empty">Nenhum registro no período.</td></tr>';return;}
  tb.innerHTML=regs.map(r=>`<tr class="${r.tipo==='Retrabalho'?'retrab':''}"
    data-txt="${(r.numero_cesto+' '+r.ordem+' '+r.material+' '+(r.texto_breve||'')+' '+r.operador_prep+' '+r.operador_banho).toLowerCase()}"
    data-tipo="${r.tipo}" data-proc="${r.processo}">
    <td>${r.id}</td><td><strong>${r.numero_cesto}</strong></td><td>${r.n_itens>1?(r.ordem+' +'+(r.n_itens-1)):(r.ordem||'—')}</td>
    <td>${r.material||'—'}</td><td><span class="small">${r.texto_breve||'—'}</span></td><td>${r.qtd_total}</td>
    <td>${r.processo||'—'}</td><td><span class="pill ${r.tipo==='Retrabalho'?'pill-retrab':'pill-normal'}">${r.tipo}</span></td>
    <td class="mono">${r.prep_minutos}</td><td class="mono">${r.banho_minutos}</td>
    <td><span class="small">${r.banho_fim||''}</span></td>
  </tr>`).join('');
  aplicarFiltroTexto();
}

function aplicarFiltroTexto(){
  const txt=(document.getElementById('fTexto')?.value||'').toLowerCase();
  const tipo=document.getElementById('fTipo')?.value||'';
  const proc=document.getElementById('fProc')?.value||'';
  document.querySelectorAll('#tbody tr[data-txt]').forEach(tr=>{
    const ok=(!txt||tr.dataset.txt.includes(txt))&&(!tipo||tr.dataset.tipo===tipo)&&(!proc||tr.dataset.proc===proc);
    tr.style.display=ok?'':'none';
  });
}

document.addEventListener('DOMContentLoaded',()=>{
  ['fTexto','fTipo','fProc'].forEach(id=>{const e=document.getElementById(id);if(e)e.addEventListener(id==='fTexto'?'input':'change',aplicarFiltroTexto);});
  ['fDe','fAte'].forEach(id=>{const e=document.getElementById(id);if(e)e.addEventListener('change',atualizar);});
  atualizar();
  setInterval(atualizar,8000);
});
window.atualizar=atualizar;
