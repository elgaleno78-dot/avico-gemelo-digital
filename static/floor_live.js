// Fuente 02 · Censo Universal HIBPC · corte 2026-09-17
// Solo estado agregado de camas; no contiene identificadores de pacientes.
const FLOOR_LIVE_CUTOFF='2026-09-17';
const FLOOR_LIVE_OCCUPIED=[4,5,7,8,9,10,11,14,15,16,17,19,21,24,25];

pages.areas=function(){
  stopRefresh(); T.textContent='Áreas · estado real de Piso';
  const floorOcc=FLOOR_LIVE_OCCUPIED.length;
  A.innerHTML=`<div class="sim-kpis"><div><strong>27</strong><span>Camas censables</span></div><div class="kpi-green"><strong>${floorOcc}</strong><span>Piso ocupado</span></div><div><strong>${27-floorOcc}</strong><span>Piso disponible</span></div><div><strong>${Math.round(floorOcc/27*1000)/10}%</strong><span>Ocupación Piso</span></div></div><div class="sim-head"><div><b class="green">● REAL · CENSO UNIVERSAL · ${FLOOR_LIVE_CUTOFF}</b></div></div><div class="floorplan"><section class="room triage-room"><h3>TRIAGE OBSTÉTRICO</h3><small>1 consultorio</small>${bed('CONS-01',false)}</section><section class="room labor-room"><h3>LABOR</h3><small>9 camas · L02 = Código Mater</small><div class="bed-grid">${Array.from({length:9},(_,i)=>bed(i===1?'L02 · CÓDIGO MATER':'L'+String(i+1).padStart(2,'0'),false)).join('')}</div></section><section class="room exp-room"><h3>SALA DE EXPULSIÓN</h3>${bed('E01',false)}</section><section class="room mix-room"><h3>SALA MIXTA</h3>${bed('M01',false)}</section><section class="room qx-room"><h3>PUERPERIO DE BAJO RIESGO</h3><small>5 camas</small>${bedGroup('PBR',5,0)}</section><section class="room floor-room"><h3>PISO DE GINECOLOGÍA Y OBSTETRICIA</h3><small>27 camas censables · 25–27 aislamiento</small><div class="bed-grid floor-beds">${Array.from({length:27},(_,i)=>bed(String(i+1).padStart(2,'0'),FLOOR_LIVE_OCCUPIED.includes(i+1),i>=24)).join('')}</div></section><section class="room rec-room"><h3>RECUPERACIÓN</h3>${bedGroup('RP',5,0)}</section></div><div class="sim-bottom"><div><span class="legend free"></span>Libre <span class="legend occupied-dot"></span>Ocupada</div><div>REAL · Piso actualizado al ${FLOOR_LIVE_CUTOFF} · sin datos identificables</div></div>`;
};

pages.monitor=async function(){
  stopRefresh(); T.textContent='Monitor · estado real';
  async function draw(){
    const r=await fetch('/api/real-state',{cache:'no-store'}).then(x=>x.json()), b=r.births,l=r.labor,u=r.triage,v=x=>x==null?'—':x;
    const n=FLOOR_LIVE_OCCUPIED.length;
    A.innerHTML=`<div class="card full"><h3>CORTE OPERATIVO</h3><b class="green">REAL · ${r.cutoff_date}</b><p class="note">Última consulta: ${new Date(r.as_of).toLocaleString('es-MX')} · actualización automática cada 60 segundos.</p></div><div class="grid">${card('Ingresos a Urgencias / Triage',v(u.value),u.last_valid_date?'Corte '+u.last_valid_date:'Fuente aún no conectada')}${card('Labor',l.value==null?'—':l.value+'/'+l.capacity,l.last_valid_date?'Corte '+l.last_valid_date:'Fuente aún no conectada')}${card('Piso',n+'/27','Censo Universal · corte '+FLOOR_LIVE_CUTOFF+' · '+Math.round(n/27*1000)/10+'%')}${card('Nacimientos',v(b.value),b.last_valid_date?'Corte '+b.last_valid_date+(b.vaginal!=null?' · Partos '+b.vaginal+' · Cesáreas '+b.cesareans:''):'Sin corte')}</div>`;
  }
  await draw(); refreshTimer=setInterval(draw,60000);
};

// Corrige el primer render de Monitor, que app.js ejecuta antes de cargar este parche.
if(document.querySelector('nav button.active')?.dataset.page==='monitor') pages.monitor();
