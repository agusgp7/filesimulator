let rawRows = [];
let filteredRows = [];
let charts = {};
let curveMode = '1200';

const els = {
  fileInput: document.getElementById('fileInput'),
  dropZone: document.getElementById('dropZone'),
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
  exportCsv: document.getElementById('exportCsv'),
  insights: document.getElementById('insights'),
  heatmapGrid: document.getElementById('heatmapGrid'),
  periodBadge: document.getElementById('periodBadge'),
  themeToggle: document.getElementById('themeToggle')
};

document.querySelectorAll('.nav-item').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    setTimeout(() => Object.values(charts).forEach(c => c && c.resize()), 80);
  });
});

document.querySelectorAll('.chip[data-curve]').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.chip[data-curve]').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    curveMode = btn.dataset.curve;
    renderDashboardCharts(filteredRows.filter(r => r.magnitude === els.mainMagnitude.value));
  });
});

els.fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
els.applyFilters.addEventListener('click', applyFilters);
els.resetFilters.addEventListener('click', resetFilters);
els.exportCsv.addEventListener('click', exportSummaryCsv);
els.themeToggle.addEventListener('click', toggleTheme);

['dragenter','dragover'].forEach(ev => els.dropZone.addEventListener(ev, e => {
  e.preventDefault();
  els.dropZone.classList.add('dragover');
}));
['dragleave','drop'].forEach(ev => els.dropZone.addEventListener(ev, e => {
  e.preventDefault();
  els.dropZone.classList.remove('dragover');
}));
els.dropZone.addEventListener('drop', e => {
  const file = e.dataTransfer.files?.[0];
  if(file) handleFile(file);
});

function toggleTheme(){
  const html = document.documentElement;
  const next = html.dataset.theme === 'dark' ? 'light' : 'dark';
  html.dataset.theme = next;
  els.themeToggle.textContent = next === 'dark' ? '☀️ Modo claro' : '🌙 Modo oscuro';
  Object.values(charts).forEach(c => c && c.update());
}

async function handleFile(file){
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
    setStatus(`Archivo cargado: <strong>${file.name}</strong>. Registros válidos: ${rawRows.length.toLocaleString('es-UY')}. Magnitudes: ${mags}.`, 'ok');
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
    const intervalRaw = findValue(r, ['INTERVALO','INTERVAL','PERIODO']);
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
  if(typeof v === 'number') {
    const parsed = XLSX.SSF.parse_date_code(v);
    return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, Math.floor(parsed.S || 0)) : null;
  }
  if(typeof v === 'string'){
    const direct = new Date(v.replace(' ', 'T'));
    if(!isNaN(direct)) return direct;
    const m = v.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if(m){
      const y = Number(m[3].length === 2 ? '20'+m[3] : m[3]);
      return new Date(y, Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0));
    }
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
  renderHeatmap(aeRows);
  renderInsights(aeRows, q1Rows);
}

function renderKpis(rows, mag){
  const total = sum(rows, 'value');
  const max = rows.length ? Math.max(...rows.map(r => r.value)) : 0;
  const avg = rows.length ? total / rows.length : 0;
  const peakKw = rows.length ? Math.max(...rows.map(r => powerKw(r))) : 0;
  const unit = mag === 'AE' ? 'kWh' : 'kvarh';
  const dates = rows.length ? `${fmtDate(rows[0].date)} a ${fmtDate(rows[rows.length-1].date)}` : '-';
  els.periodBadge.textContent = dates;
  els.kpiGrid.innerHTML = [
    kpi('Total ' + mag, `${fmt(total)} ${unit}`, 'good'),
    kpi('Promedio por intervalo', `${fmt(avg)} ${unit}`),
    kpi('Máximo por intervalo', `${fmt(max)} ${unit}`),
    kpi('Pico estimado', `${fmt(peakKw)} kW`, 'warn'),
    kpi('Registros', rows.length.toLocaleString('es-UY')),
    kpi('Período analizado', dates),
    kpi('Días con datos', countUnique(rows, r => dayKey(r.date)).toLocaleString('es-UY')),
    kpi('Validaciones no VALID', rows.filter(r => r.validation && r.validation.toUpperCase() !== 'VALID').length.toLocaleString('es-UY'))
  ].join('');
}
function kpi(label, value, cls=''){ return `<div class="kpi ${cls}"><span>${label}</span><strong>${value}</strong></div>`; }

function renderDashboardCharts(rows){
  makeBar('monthlyChart', groupSum(rows, monthKey), 'Mes', 'Energía');
  makeBar('dailyChart', groupSum(rows, dayKey), 'Día', 'Energía');
  makeBar('hourlyChart', groupSum(rows, r => pad(r.date.getHours()) + ':00'), 'Hora', 'Energía');
  makeBar('weekdayChart', groupSum(rows, r => weekdayName(r.date)), 'Día', 'Energía', weekdayOrder);
  const curveRows = curveMode === 'all' ? rows : rows.slice(-1200);
  makeLine('loadCurveChart', curveRows.map(r => ({ label: fmtDateTime(r.date), value: powerKw(r) })), 'kW estimado');
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
    kpi('AE total', `${fmt(ae)} kWh`, 'good'),
    kpi('Q1 total', `${fmt(q1)} kvarh`),
    kpi('Q1 / AE', fmt(ratio), ratio > .426 ? 'danger' : 'good'),
    kpi('Referencia 0,426', ratio > 0.426 ? 'Supera referencia' : 'Dentro de referencia', ratio > .426 ? 'danger' : 'good')
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

function renderInsights(aeRows, q1Rows){
  if(!aeRows.length){ els.insights.innerHTML = '<div class="insight">Cargá un archivo para ver conclusiones automáticas.</div>'; return; }
  const byDay = Object.entries(groupSum(aeRows, dayKey)).sort((a,b)=>b[1]-a[1]);
  const byHour = Object.entries(groupSum(aeRows, r => pad(r.date.getHours()) + ':00')).sort((a,b)=>b[1]-a[1]);
  const total = sum(aeRows,'value');
  const q1 = sum(q1Rows,'value');
  const ratio = total ? q1/total : 0;
  const bands = { valle:0, llano:0, puntaHabiles:0, puntaFinSemana:0 };
  aeRows.forEach(r => bands[getBand(r.date)] += r.value);
  const puntaTotal = bands.puntaHabiles + bands.puntaFinSemana;
  els.insights.innerHTML = [
    insight('Día de mayor consumo', `${byDay[0]?.[0] || '-'} con ${fmt(byDay[0]?.[1] || 0)} kWh.`),
    insight('Hora con mayor energía acumulada', `${byHour[0]?.[0] || '-'} concentra ${fmt(byHour[0]?.[1] || 0)} kWh.`),
    insight('Consumo en Punta', `${fmt(puntaTotal)} kWh, equivalente a ${fmt(total ? puntaTotal/total*100 : 0)}% del total AE.`),
    insight('Reactiva Q1/AE', `${fmt(ratio)}. ${ratio > .426 ? 'Conviene revisar reactiva.' : 'Está dentro de referencia 0,426.'}`)
  ].join('');
}
function insight(title, text){ return `<div class="insight"><strong>${title}</strong>${text}</div>`; }

const weekdayOrder = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
function weekdayName(d){ return weekdayOrder[(d.getDay()+6)%7]; }

function renderHeatmap(aeRows){
  if(!aeRows.length){ els.heatmapGrid.innerHTML = 'Sin datos.'; return; }
  const buckets = {};
  aeRows.forEach(r => {
    const key = `${weekdayName(r.date)}|${r.date.getHours()}`;
    if(!buckets[key]) buckets[key] = { sum:0, count:0 };
    buckets[key].sum += r.value; buckets[key].count++;
  });
  const avgs = {};
  let max = 0;
  Object.entries(buckets).forEach(([k,v]) => { avgs[k] = v.sum / v.count; max = Math.max(max, avgs[k]); });
  let html = '<div class="heat-table"><div></div>';
  for(let h=0; h<24; h++) html += `<div class="heat-head">${pad(h)}</div>`;
  weekdayOrder.forEach(day => {
    html += `<div class="heat-day">${day.slice(0,3)}</div>`;
    for(let h=0; h<24; h++){
      const val = avgs[`${day}|${h}`] || 0;
      const intensity = max ? val / max : 0;
      const bg = heatColor(intensity);
      html += `<div class="heat-cell" title="${day} ${pad(h)}:00 - ${fmt(val)} kWh promedio" style="background:${bg}">${val ? fmt(val) : ''}</div>`;
    }
  });
  html += '</div>';
  els.heatmapGrid.innerHTML = html;
}
function heatColor(i){
  const light = 95 - Math.round(i * 42);
  const sat = 95;
  const hue = 210 - Math.round(i * 170);
  return `hsl(${hue} ${sat}% ${light}%)`;
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
function makeBar(id, obj, xLabel, yLabel, order=null){
  let labels = Object.keys(obj).sort();
  if(order) labels = order.filter(l => Object.prototype.hasOwnProperty.call(obj,l));
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'bar', data:{ labels, datasets:[{ label:yLabel, data:labels.map(l=>obj[l]), borderWidth:1 }] }, options:{ responsive:true, plugins:{ legend:{ display:false } }, scales:{ x:{ title:{ display:true, text:xLabel } }, y:{ beginAtZero:true } } } });
}
function makeMultiBar(id, labels, datasets){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'bar', data:{ labels, datasets }, options:{ responsive:true, scales:{ y:{ beginAtZero:true } } } });
}
function makePie(id, data){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'doughnut', data:{ labels:data.map(d=>d[0]), datasets:[{ data:data.map(d=>d[1]) }] }, options:{ responsive:true, cutout:'62%' } });
}
function makeLine(id, data, label){
  destroyChart(id);
  charts[id] = new Chart(document.getElementById(id), { type:'line', data:{ labels:data.map(d=>d.label), datasets:[{ label, data:data.map(d=>d.value), pointRadius:0, tension:.25, fill:true }] }, options:{ responsive:true, animation:false, scales:{ y:{ beginAtZero:true } } } });
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
    ['Q1 / AE', sum(aeRows,'value') ? sum(q1Rows,'value') / sum(aeRows,'value') : 0],
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
