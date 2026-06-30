let rawRows = [];
let filteredRows = [];
let charts = {};

const els = {
  fileInput: document.getElementById('fileInput'),
  status: document.getElementById('status'),
  dateFrom: document.getElementById('dateFrom'),
  dateTo: document.getElementById('dateTo'),
  mainMagnitude: document.getElementById('mainMagnitude'),
  applyFilters: document.getElementById('applyFilters'),
  resetFilters: document.getElementById('resetFilters'),
  kpiGrid: document.getElementById('kpiGrid'),
  reactiveKpis: document.getElementById('reactiveKpis'),
  bandsTable: document.getElementById('bandsTable'),
  topDaysTable: document.getElementById('topDaysTable'),
  topPeaksTable: document.getElementById('topPeaksTable'),
  exportCsv: document.getElementById('exportCsv')
};

document.querySelectorAll('.tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    Object.values(charts).forEach(c => c && c.resize());
  });
});

els.fileInput.addEventListener('change', handleFile);
els.applyFilters.addEventListener('click', applyFilters);
els.resetFilters.addEventListener('click', resetFilters);
els.exportCsv.addEventListener('click', exportSummaryCsv);

async function handleFile(event){
  const file = event.target.files[0];
  if(!file) return;
  setStatus('Leyendo archivo...', 'empty');
  try{
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type:'array', cellDates:true });
    const sheetName = workbook.SheetNames.includes('Measures') ? 'Measures' : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval:null });
    rawRows = normalizeRows(rows);
    if(rawRows.length === 0) throw new Error('No encontré registros válidos con FECHA-HORA, MAGNITUD y VALOR.');
    setupDateInputs(rawRows);
    applyFilters();
    const mags = [...new Set(rawRows.map(r => r.magnitude))].join(', ');
    setStatus(`Archivo cargado: ${file.name}. Registros válidos: ${rawRows.length.toLocaleString('es-UY')}. Magnitudes: ${mags}.`, 'ok');
  }catch(err){
    console.error(err);
    setStatus(`Error: ${err.message}`, 'error');
  }
}

function normalizeRows(rows){
  return rows.map(r => {
    const dateValue = findValue(r, ['FECHA-HORA','FECHA HORA','FECHA','DATE','DATETIME']);
    const magnitude = String(findValue(r, ['MAGNITUD','MAG','MAGNITUDE']) ?? '').trim().toUpperCase();
    const valueRaw = findValue(r, ['VALOR','VALUE','ENERGIA','CONSUMO']);
    const intervalRaw = findValue(r, ['INTERVALO','INTERVAL']);
    const validation = String(findValue(r, ['RESULT VALIDACION','VALIDACION','VALIDATION']) ?? '').trim();
    const date = parseExcelDate(dateValue);
    const value = parseNumber(valueRaw);
    const intervalMinutes = parseInterval(intervalRaw) || 15;
    if(!date || !magnitude || !Number.isFinite(value)) return null;
    return { date, magnitude, value, intervalMinutes, validation };
  }).filter(Boolean).sort((a,b) => a.date - b.date);
}

function findValue(row, possible){
  const keys = Object.keys(row);
  for(const name of possible){
    const found = keys.find(k => normalizeKey(k) === normalizeKey(name));
    if(found) return row[found];
  }
  return null;
}
function normalizeKey(k){ return String(k).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,''); }
function parseExcelDate(v){
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v === 'number') return XLSX.SSF.parse_date_code(v) ? new Date(Math.round((v - 25569) * 86400 * 1000)) : null;
  if(typeof v === 'string'){
    const d = new Date(v.replace(' ', 'T'));
    if(!isNaN(d)) return d;
  }
  return null;
}
function parseNumber(v){
  if(typeof v === 'number') return v;
  if(v === null || v === undefined || v === '') return NaN;
  return Number(String(v).replace(/\./g,'').replace(',', '.'));
}
function parseInterval(v){
  if(typeof v === 'number') return v;
  if(!v) return null;
  const m = String(v).match(/(\d+)/);
  return m ? Number(m[1]) : null;
}

function setupDateInputs(rows){
  const min = rows[0].date;
  const max = rows[rows.length-1].date;
  els.dateFrom.value = toDateInput(min);
  els.dateTo.value = toDateInput(max);
}
function toDateInput(d){ return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
function pad(n){ return String(n).padStart(2,'0'); }

function applyFilters(){
  if(!rawRows.length) return;
  const from = els.dateFrom.value ? new Date(els.dateFrom.value + 'T00:00:00') : null;
  const to = els.dateTo.value ? new Date(els.dateTo.value + 'T23:59:59') : null;
  filteredRows = rawRows.filter(r => (!from || r.date >= from) && (!to || r.date <= to));
  renderAll();
}
function resetFilters(){ setupDateInputs(rawRows); els.mainMagnitude.value = 'AE'; applyFilters(); }

function renderAll(){
  const mainMag = els.mainMagnitude.value;
  const mainRows = filteredRows.filter(r => r.magnitude === mainMag);
  const aeRows = filteredRows.filter(r => r.magnitude === 'AE');
  const q1Rows = filteredRows.filter(r => r.magnitude === 'Q1');
  renderKpis(mainRows, mainMag);
  renderDashboardCharts(mainRows);
  renderBands(aeRows);
  renderReactive(aeRows, q1Rows);
  renderRankings(aeRows);
}

function renderKpis(rows, mag){
  const total = sum(rows, 'value');
  const max = Math.max(...rows.map(r => r.value), 0);
  const avg = rows.length ? total / rows.length : 0;
  const peakKw = rows.length ? Math.max(...rows.map(r => powerKw(r)), 0) : 0;
  const unit = mag === 'AE' ? 'kWh' : 'kvarh';
  const dates = rows.length ? `${fmtDate(rows[0].date)} a ${fmtDate(rows[rows.length-1].date)}` : '-';
  els.kpiGrid.innerHTML = [
    kpi('Total ' + mag, `${fmt(total)} ${unit}`),
    kpi('Promedio por intervalo', `${fmt(avg)} ${unit}`),
    kpi('Máximo por intervalo', `${fmt(max)} ${unit}`),
    kpi('Pico estimado', `${fmt(peakKw)} kW`),
    kpi('Registros', rows.length.toLocaleString('es-UY')),
    kpi('Período analizado', dates),
    kpi('Días con datos', countUnique(rows, r => dayKey(r.date)).toLocaleString('es-UY')),
    kpi('Validaciones no VALID', rows.filter(r => r.validation && r.validation.toUpperCase() !== 'VALID').length.toLocaleString('es-UY'))
  ].join('');
}
function kpi(label, value){ return `<div class="kpi"><span>${label}</span><strong>${value}</strong></div>`; }

function renderDashboardCharts(rows){
  makeBar('monthlyChart', groupSum(rows, monthKey), 'Mes', 'Energía');
  makeBar('dailyChart', groupSum(rows, dayKey), 'Día', 'Energía');
  makeBar('hourlyChart', groupSum(rows, r => pad(r.date.getHours()) + ':00'), 'Hora', 'Energía');
  makeLine('loadCurveChart', rows.slice(-1200).map(r => ({ label: fmtDateTime(r.date), value: powerKw(r) })), 'kW estimado');
}

function renderBands(aeRows){
  const bands = { valle:0, llano:0, puntaHabiles:0, puntaFinSemana:0 };
  aeRows.forEach(r => bands[getBand(r.date)] += r.value);
  const total = Object.values(bands).reduce((a,b)=>a+b,0);
  const data = [
    ['Valle', bands.valle],
    ['Llano', bands.llano],
    ['Punta hábiles', bands.puntaHabiles],
    ['Punta sábado/domingo', bands.puntaFinSemana]
  ];
  makePie('bandsChart', data);
  els.bandsTable.innerHTML = table(['Tramo','kWh','%'], data.map(([name,val]) => [name, fmt(val), total ? fmt(val/total*100)+'%' : '0%']));
}
function getBand(d){
  const h = d.getHours() + d.getMinutes()/60;
  const weekend = d.getDay() === 0 || d.getDay() === 6;
  if(h >= 0 && h < 7) return 'valle';
  if(h >= 18 && h < 22) return weekend ? 'puntaFinSemana' : 'puntaHabiles';
  return 'llano';
}

function renderReactive(aeRows, q1Rows){
  const ae = sum(aeRows,'value');
  const q1 = sum(q1Rows,'value');
  const ratio = ae ? q1 / ae : 0;
  els.reactiveKpis.innerHTML = [
    kpi('AE total', `${fmt(ae)} kWh`),
    kpi('Q1 total', `${fmt(q1)} kvarh`),
    kpi('Q1 / AE', fmt(ratio)),
    kpi('Referencia 0,426', ratio > 0.426 ? 'Supera referencia' : 'Dentro de referencia')
  ].join('');
  const aeMonthly = groupSum(aeRows, monthKey);
  const q1Monthly = groupSum(q1Rows, monthKey);
  const labels = [...new Set([...Object.keys(aeMonthly), ...Object.keys(q1Monthly)])].sort();
  makeMultiBar('reactiveMonthlyChart', labels, [
    { label:'AE kWh', data:labels.map(l => aeMonthly[l] || 0) },
    { label:'Q1 kvarh', data:labels.map(l => q1Monthly[l] || 0) }
  ]);
  makeLine('ratioChart', labels.map(l => ({ label:l, value:(aeMonthly[l] ? (q1Monthly[l] || 0) / aeMonthly[l] : 0) })), 'Q1/AE');
}

function renderRankings(aeRows){
  const days = Object.entries(groupSum(aeRows, dayKey)).sort((a,b)=>b[1]-a[1]).slice(0,10);
  els.topDaysTable.innerHTML = table(['Día','kWh'], days.map(([d,v]) => [d, fmt(v)]));
  const peaks = [...aeRows].sort((a,b)=>powerKw(b)-powerKw(a)).slice(0,10);
  els.topPeaksTable.innerHTML = table(['Fecha-hora','kW estimado','kWh intervalo'], peaks.map(r => [fmtDateTime(r.date), fmt(powerKw(r)), fmt(r.value)]));
}

function groupSum(rows, fn){
  return rows.reduce((acc,r) => { const k = fn(r); acc[k] = (acc[k] || 0) + r.value; return acc; }, {});
}
function sum(rows, prop){ return rows.reduce((a,r)=>a+(Number(r[prop])||0),0); }
function countUnique(rows, fn){ return new Set(rows.map(fn)).size; }
function powerKw(r){ return r.value * (60 / (r.intervalMinutes || 15)); }
function dayKey(r){ const d = r.date || r; return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
function monthKey(r){ const d = r.date || r; return `${d.getFullYear()}-${pad(d.getMonth()+1)}`; }
function fmt(n){ return Number(n || 0).toLocaleString('es-UY', { maximumFractionDigits:2 }); }
function fmtDate(d){ return d.toLocaleDateString('es-UY'); }
function fmtDateTime(d){ return d.toLocaleString('es-UY', { dateStyle:'short', timeStyle:'short' }); }
function table(headers, rows){
  return `<table class="table"><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}

function destroyChart(id){ if(charts[id]) charts[id].destroy(); }
function makeBar(id, obj, xLabel, yLabel){
  const labels = Object.keys(obj).sort();
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'bar', data:{ labels, datasets:[{ label:yLabel, data:labels.map(l=>obj[l]) }] }, options:{ responsive:true, plugins:{ legend:{ display:false } }, scales:{ x:{ title:{ display:true, text:xLabel } }, y:{ beginAtZero:true } } } });
}
function makeMultiBar(id, labels, datasets){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'bar', data:{ labels, datasets }, options:{ responsive:true, scales:{ y:{ beginAtZero:true } } } });
}
function makePie(id, data){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'doughnut', data:{ labels:data.map(d=>d[0]), datasets:[{ data:data.map(d=>d[1]) }] }, options:{ responsive:true } });
}
function makeLine(id, data, label){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'line', data:{ labels:data.map(d=>d.label), datasets:[{ label, data:data.map(d=>d.value), pointRadius:0, tension:.2 }] }, options:{ responsive:true, scales:{ y:{ beginAtZero:true } } } });
}
function setStatus(msg, type){ els.status.className = `status ${type}`; els.status.innerHTML = msg; }
function exportSummaryCsv(){
  if(!filteredRows.length) return;
  const aeRows = filteredRows.filter(r => r.magnitude === 'AE');
  const q1Rows = filteredRows.filter(r => r.magnitude === 'Q1');
  const bands = { valle:0, llano:0, puntaHabiles:0, puntaFinSemana:0 };
  aeRows.forEach(r => bands[getBand(r.date)] += r.value);
  const lines = [
    ['Indicador','Valor'],
    ['AE total kWh', sum(aeRows,'value')],
    ['Q1 total kvarh', sum(q1Rows,'value')],
    ['Valle kWh', bands.valle],
    ['Llano kWh', bands.llano],
    ['Punta habiles kWh', bands.puntaHabiles],
    ['Punta fin semana kWh', bands.puntaFinSemana]
  ];
  const csv = lines.map(r => r.join(';')).join('\n');
  const blob = new Blob([csv], { type:'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'resumen-consumos.csv'; a.click();
  URL.revokeObjectURL(url);
}
