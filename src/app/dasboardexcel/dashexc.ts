import { Component, AfterViewInit, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Chart, registerables } from 'chart.js';
import * as XLSX from 'xlsx';

Chart.register(...registerables);

@Component({
  selector: 'app-dashexc',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashexc.html',
  styleUrls: ['./dashexc.css']
})
export class DashexcComponent implements AfterViewInit {
  @ViewChild('chartTransp') chartTranspRef!: ElementRef;
  @ViewChild('chartAlmacen') chartAlmacenRef!: ElementRef;
  @ViewChild('chartTiendas') chartTiendasRef!: ElementRef;
  @ViewChild('chartAgri') chartAgriRef!: ElementRef;

  charts: any = {};
  PALETTE = ['#2dd4bf','#f2a93c','#e2564f','#7aa2f7','#c792ea','#9ece6a'];
  MESES_ORDER = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  // --- ESTADOS INICIALES (Tus datos hardcodeados exactos) ---
  makandState: any = {
    src: 'VIAJEROS_Y_TERCEROS_MAKAND.xlsx · JUNIO 2026 (mes más reciente completo)', status: '', statusClass: '',
    note: '258 viajes registrados en junio, repartidos entre 3 transportadoras. Los % de cumplimiento de llegada corresponden a la llegada del vehículo a planta según la hora citada.',
    kpiTotal: '258', kpiPeriodo: 'junio 2026', kpiCumple: '80.6%', kpiCajas: '115,432', kpiCajasSub: 'MAKAND+ARSITRANS+POLAR', kpiLider: 'MAKAND', kpiLiderSub: '51.2% de los viajes',
    table: [
      { name: 'MAKAND', color: '#2dd4bf', total: 132, pct: '51.2%', llegada: '93.2%', llegadaClass: 'good', cargue: '49.2%', cargueClass: 'mid' },
      { name: 'ARSITRANS', color: '#f2a93c', total: 78, pct: '30.2%', llegada: '71.8%', llegadaClass: 'good', cargue: '48.7%', cargueClass: 'mid' },
      { name: 'POLAR', color: '#e2564f', total: 48, pct: '18.6%', llegada: '60.4%', llegadaClass: 'mid', cargue: '89.6%', cargueClass: 'good' }
    ]
  };

  tiendasState: any = {
    src: 'ON_TIME_EN_TIENDAS.xlsx · JUNIO 2026 (mes más reciente completo)', status: '', statusClass: '',
    note: '929 registros de llegada/salida en ruta durante junio. El cumplimiento mide si el vehículo permaneció en el punto dentro del tiempo estándar asignado.',
    kpiTotal: '929', kpiPeriodo: 'junio 2026', kpiCumple: '33.8%', kpiCumpleSub: '314 registros', kpiParcial: '12.6%', kpiParcialSub: '117 registros', kpiNoCumple: '53.6%', kpiNoCumpleSub: '498 registros', rutasMesLabel: '(junio)',
    tableCedi: [
      { name: 'Tiendas', total: 602, pct: '64.8%' }, { name: 'CEDI 2', total: 200, pct: '21.5%' },
      { name: 'CEDI ARA', total: 74, pct: '8.0%' }, { name: 'CEDI 1', total: 53, pct: '5.7%' }
    ],
    tableRutas: [
      { name: 'Plataforma Siberia', total: 63, pct: '—', pctClass: 'bad' },
      { name: 'Ara Cota', total: 45, pct: '26.7%', pctClass: 'mid' },
      { name: 'Olímpica', total: 37, pct: '—', pctClass: 'bad' },
      { name: 'Plataforma Cencosud', total: 34, pct: '—', pctClass: 'bad' },
      { name: 'D1 Tocancipá', total: 30, pct: '—', pctClass: 'bad' }
    ]
  };

  agriState: any = {
    src: 'ON_TIME_AGRICULTORES_2026_ACT.xlsx · Semana 23 (más reciente disponible)', status: '', statusClass: '',
    note: 'Cumplimiento semanal por agricultor en tres momentos del proceso: llegada del transporte, tiempo en finca y llegada a planta.',
    kpiSemanaLabel: 'semana 23', kpiTotal: '70', kpiTotalSub: '6 agricultores', kpiLlegada: '54.3%', kpiTiempo: '21.4%', kpiPlanta: '0.0%', chartLabel: '(S23)', tableLabel: 'semana 23',
    table: [
      { name: 'Diego', total: 6, llegada: '16.7%', llegadaClass: 'bad', tiempo: '16.7%', tiempoClass: 'bad', planta: '0.0%', plantaClass: 'bad' },
      { name: 'Ferrucas', total: 47, llegada: '65.9%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '0.0%', plantaClass: 'bad' },
      { name: 'Gabriel', total: 7, llegada: '0.0%', llegadaClass: 'bad', tiempo: '85.7%', tiempoClass: 'good', planta: '0.0%', plantaClass: 'bad' },
      { name: 'Georgeth', total: 1, llegada: '0.0%', llegadaClass: 'bad', tiempo: '100.0%', tiempoClass: 'good', planta: '0.0%', plantaClass: 'bad' },
      { name: 'José Tibaquicha', total: 3, llegada: '0.0%', llegadaClass: 'bad', tiempo: '33.3%', tiempoClass: 'mid', planta: '0.0%', plantaClass: 'bad' },
      { name: 'Lechugas del Día', total: 6, llegada: '100.0%', llegadaClass: 'good', tiempo: '100.0%', tiempoClass: 'good', planta: '—', plantaClass: 'mid' }
    ]
  };

  constructor() {
    Chart.defaults.color = '#8b96a5';
    Chart.defaults.font.family = "'IBM Plex Mono', monospace";
    Chart.defaults.font.size = 11;
  }

  ngAfterViewInit(): void {
    // Renderizado estático inicial de tus gráficas originales
    this.makeDoughnut('chartTransp', this.chartTranspRef.nativeElement, ['MAKAND','ARSITRANS','POLAR'], [132,78,48], ['#2dd4bf','#f2a93c','#e2564f']);
    this.makeStackedBar('chartAlmacen', this.chartAlmacenRef.nativeElement, ['D1','ARA','ÉXITO'], [
      {label:'MAKAND', data:[96,10,26], backgroundColor:'#2dd4bf'},
      {label:'ARSITRANS', data:[32,32,14], backgroundColor:'#f2a93c'},
      {label:'POLAR', data:[48,0,0], backgroundColor:'#e2564f'}
    ]);
    this.makeDoughnut('chartTiendas', this.chartTiendasRef.nativeElement, ['Cumple','Cumple parcial','No cumple'], [314,117,498], ['#2dd4bf','#f2a93c','#e2564f']);
    this.makeHBar('chartAgri', this.chartAgriRef.nativeElement, ['Diego','Ferrucas','Gabriel','Georgeth','José T.','Lechugas'], [16.7,65.9,0,0,0,100], ['#e2564f','#2dd4bf','#e2564f','#e2564f','#e2564f','#2dd4bf']);
  }

  // ---------- TUS HELPERS EXACTOS ----------
  pillClass(pct: any): string {
    if(pct===null||pct===undefined||isNaN(pct)) return 'mid';
    if(pct>=70) return 'good'; if(pct>=40) return 'mid'; return 'bad';
  }
  fmtPct(pct: any): string { return (pct===null||pct===undefined||isNaN(pct)) ? '—' : pct.toFixed(1)+'%'; }
  findSheet(wb: any, patterns: string[]): string | null {
    const names = wb.SheetNames;
    for(const p of patterns){ const hit = names.find((n:string) => n.toUpperCase().replace(/\s+/g,' ').trim().includes(p)); if(hit) return hit; }
    return null;
  }
  sheetRows(wb: any, name: string): any[] { return XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:null, raw:true}); }
  colIndex(header: any[], name: string): number { return header.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase()); }
  pickLatestCompleteMonth(counts: any): string | null {
    const months = Object.keys(counts).filter(m=>this.MESES_ORDER.includes(m));
    months.sort((a,b)=>this.MESES_ORDER.indexOf(a)-this.MESES_ORDER.indexOf(b));
    if(!months.length) return null;
    const maxCount = Math.max(...months.map(m=>counts[m]));
    let chosen = months[0];
    for(const m of months){ if(counts[m] >= 0.5*maxCount) chosen = m; }
    return chosen;
  }

  // ---------- TUS CREADORES DE GRÁFICAS ----------
  makeDoughnut(id: string, canvasRef: any, labels: any, data: any, colors: any) {
    if(this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(canvasRef, {
      type:'doughnut', data:{ labels, datasets:[{ data, backgroundColor:colors, borderColor:'#171e27', borderWidth:3 }] },
      options:{ plugins:{legend:{position:'bottom', labels:{boxWidth:10, padding:16}}}, cutout:'62%' }
    });
  }
  makeStackedBar(id: string, canvasRef: any, labels: any, datasets: any) {
    if(this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(canvasRef, {
      type:'bar', data:{ labels, datasets },
      options:{ responsive:true, maintainAspectRatio:false, scales:{ x:{stacked:true, grid:{display:false}}, y:{stacked:true, grid:{color:'#2a3340'}} }, plugins:{legend:{position:'bottom', labels:{boxWidth:10,padding:14}}} }
    });
  }
  makeHBar(id: string, canvasRef: any, labels: any, data: any, colors: any, max: number=100) {
    if(this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(canvasRef, {
      type:'bar', data:{ labels, datasets:[{ data, backgroundColor:colors }] },
      options:{ responsive:true, maintainAspectRatio:false, indexAxis:'y', scales:{ x:{max:max, grid:{color:'#2a3340'}}, y:{grid:{display:false}} }, plugins:{legend:{display:false}} }
    });
  }

  // ---------- EL UPLOAD ----------
  setStatus(key: string, msg: string, cls: string = '') {
    const state = key === 'makand' ? this.makandState : key === 'tiendas' ? this.tiendasState : this.agriState;
    state.status = msg; state.statusClass = cls ? `upload-status ${cls}` : 'upload-status';
  }

  handleUpload(evt: any, kind: string) {
    const file = evt.target.files[0]; if(!file) return;
    this.setStatus(kind, 'Leyendo '+file.name+'…');
    const reader = new FileReader();
    reader.onload = (e: any) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:false});
        if(kind==='makand'){ const d = this.parseMakand(wb); this.renderMakand(d, file.name); }
        else if(kind==='tiendas'){ const d = this.parseTiendas(wb); this.renderTiendas(d, file.name); }
        else { const d = this.parseAgri(wb); this.renderAgri(d, file.name); }
        this.setStatus(kind, '✓ actualizado ' + new Date().toLocaleString('es-CO'), 'ok');
      } catch(err: any){ this.setStatus(kind, 'Error: '+err.message, 'err'); }
    };
    reader.onerror = () => this.setStatus(kind, 'No se pudo leer el archivo.', 'err');
    reader.readAsArrayBuffer(file);
  }

  // ---------- PARSER 1: MAKAND ----------
  parseMakand(wb: any) {
    const sheetName = this.findSheet(wb, ['DT VIAJEROS']); if(!sheetName) throw new Error('No encontré la hoja "DT VIAJEROS."');
    const rows = this.sheetRows(wb, sheetName); const header = rows[0];
    const iMes = this.colIndex(header,'MES'), iTransp = this.colIndex(header,'TRANSPORTE'), iAlmacen = this.colIndex(header,'ALMACEN'), iLlegada = this.colIndex(header,'CUMPLIMIENTO LLEGADA VEHICULO'), iCargue = this.colIndex(header,'CUMPLIMIENTO CARGUE VEHICULO'), iCajas = this.colIndex(header,'TOTAL CAJAS');
    if(iMes<0||iTransp<0) throw new Error('No encontré las columnas MES/TRANSPORTE esperadas.');
    const data = rows.slice(1).filter(r=>r && r[iMes] && r[iTransp]);
    const counts: any = {}; data.forEach(r=>{ const m=String(r[iMes]).trim().toUpperCase(); counts[m]=(counts[m]||0)+1; });
    const month = this.pickLatestCompleteMonth(counts); const rowsMonth = data.filter(r=>String(r[iMes]).trim().toUpperCase()===month);
    const transpMap: any = {}; const almacenSet = new Set<string>();
    rowsMonth.forEach(r=>{
      const t = String(r[iTransp]).trim().toUpperCase();
      if(!transpMap[t]) transpMap[t] = {total:0, llegadaOk:0, cargueOk:0, cajas:0, almacenes:{}};
      transpMap[t].total++;
      if(iLlegada>=0 && String(r[iLlegada]).trim().toUpperCase()==='CUMPLE') transpMap[t].llegadaOk++;
      if(iCargue>=0 && String(r[iCargue]).trim().toUpperCase()==='CUMPLE') transpMap[t].cargueOk++;
      if(iCajas>=0){ const v=parseFloat(r[iCajas]); if(!isNaN(v)) transpMap[t].cajas += v; }
      const alm = iAlmacen>=0 ? String(r[iAlmacen]||'').trim() : 'N/D';
      almacenSet.add(alm); transpMap[t].almacenes[alm] = (transpMap[t].almacenes[alm]||0)+1;
    });
    const total = rowsMonth.length; const transportadoras = Object.keys(transpMap).sort((a,b)=>transpMap[b].total-transpMap[a].total);
    const overallLlegada = total ? ((Object.values(transpMap) as any[]).reduce((s: number, t: any) => s + (t.llegadaOk || 0), 0) / total * 100) : 0;
    const totalCajas = Object.values(transpMap).reduce((s:any,t:any)=>s+t.cajas,0); const lider = transportadoras[0];
    return { month, total, transportadoras, transpMap, almacenes:[...almacenSet], overallLlegada, totalCajas, lider, liderPct: total ? (transpMap[lider].total/total*100) : 0 };
  }

  renderMakand(d: any, sourceLabel: string) {
    this.makandState.src = sourceLabel + ' · ' + d.month + ' (mes más reciente completo)';
    this.makandState.note = `${d.total} viajes registrados en ${d.month.toLowerCase()}, repartidos entre ${d.transportadoras.length} transportadora(s).`;
    this.makandState.kpiTotal = d.total; this.makandState.kpiPeriodo = d.month.toLowerCase();
    this.makandState.kpiCumple = this.fmtPct(d.overallLlegada);
    this.makandState.kpiCajas = Math.round(d.totalCajas).toLocaleString('es-CO');
    this.makandState.kpiLider = d.lider; this.makandState.kpiLiderSub = this.fmtPct(d.liderPct)+' de los viajes';
    
    this.makandState.table = d.transportadoras.map((t:any, i:number) => {
      const m = d.transpMap[t]; const pctTotal = m.total/d.total*100; const pctLlegada = m.total ? m.llegadaOk/m.total*100 : 0; const pctCargue = m.total ? m.cargueOk/m.total*100 : 0;
      return { name: t, color: this.PALETTE[i%this.PALETTE.length], total: m.total, pct: this.fmtPct(pctTotal), llegada: this.fmtPct(pctLlegada), llegadaClass: this.pillClass(pctLlegada), cargue: this.fmtPct(pctCargue), cargueClass: this.pillClass(pctCargue) };
    });
    this.makeDoughnut('chartTransp', this.chartTranspRef.nativeElement, d.transportadoras, d.transportadoras.map((t:any)=>d.transpMap[t].total), d.transportadoras.map((_:any,i:number)=>this.PALETTE[i%this.PALETTE.length]));
    this.makeStackedBar('chartAlmacen', this.chartAlmacenRef.nativeElement, d.almacenes, d.transportadoras.map((t:any,i:number)=>({ label:t, backgroundColor:this.PALETTE[i%this.PALETTE.length], data:d.almacenes.map((a:any)=>d.transpMap[t].almacenes[a]||0) })));
  }

  // ---------- PARSER 2: TIENDAS ----------
  parseTiendas(wb: any) {
    const sheetName = this.findSheet(wb, ['DETALLE REGIONES']); if(!sheetName) throw new Error('No encontré la hoja "Detalle regiones"');
    const rows = this.sheetRows(wb, sheetName); const header = rows[0];
    const iMes = this.colIndex(header,'MES'), iCump = this.colIndex(header,'CUMPLIMIENTO'), iRegion = this.colIndex(header,'REGIÓN'), iCedi = this.colIndex(header,'CEDI');
    if(iMes<0||iCump<0) throw new Error('No encontré las columnas MES/CUMPLIMIENTO.');
    const data = rows.slice(1).filter(r=>r && r[iMes]);
    const counts: any = {}; data.forEach(r=>{ const m=String(r[iMes]).trim().toUpperCase(); counts[m]=(counts[m]||0)+1; });
    const month = this.pickLatestCompleteMonth(counts); const rowsMonth = data.filter(r=>String(r[iMes]).trim().toUpperCase()===month);
    const dist: any = {}; const cediMap: any = {}; const regionMap: any = {};
    rowsMonth.forEach(r=>{
      const c = iCump>=0 ? String(r[iCump]||'').trim().toUpperCase() : 'N/D'; dist[c] = (dist[c]||0)+1;
      if(iCedi>=0){ const cd=String(r[iCedi]||'').trim()||'N/D'; cediMap[cd]=(cediMap[cd]||0)+1; }
      if(iRegion>=0){ const rg = String(r[iRegion]||'').trim()||'N/D'; if(!regionMap[rg]) regionMap[rg] = {total:0, cumple:0}; regionMap[rg].total++; if(c==='CUMPLE') regionMap[rg].cumple++; }
    });
    const total = rowsMonth.length; const topRegiones = Object.keys(regionMap).sort((a,b)=>regionMap[b].total-regionMap[a].total).slice(0,5); const cedis = Object.keys(cediMap).sort((a,b)=>cediMap[b]-cediMap[a]);
    return {month, total, dist, cediMap, cedis, regionMap, topRegiones};
  }

  renderTiendas(d: any, sourceLabel: string) {
    this.tiendasState.src = sourceLabel + ' · ' + d.month + ' (mes más reciente completo)';
    this.tiendasState.note = `${d.total} registros de llegada/salida en ruta durante ${d.month.toLowerCase()}.`;
    this.tiendasState.rutasMesLabel = '('+d.month.toLowerCase()+')';
    const cumple = d.dist['CUMPLE']||0, parcial = d.dist['CUMPLE PARCIAL']||0, noCumple = d.dist['NO CUMPLE']||0;
    this.tiendasState.kpiTotal = d.total; this.tiendasState.kpiPeriodo = d.month.toLowerCase();
    this.tiendasState.kpiCumple = this.fmtPct(cumple/d.total*100); this.tiendasState.kpiCumpleSub = cumple+' registros';
    this.tiendasState.kpiParcial = this.fmtPct(parcial/d.total*100); this.tiendasState.kpiParcialSub = parcial+' registros';
    this.tiendasState.kpiNoCumple = this.fmtPct(noCumple/d.total*100); this.tiendasState.kpiNoCumpleSub = noCumple+' registros';

    this.tiendasState.tableCedi = d.cedis.map((cd:any) => ({ name: cd, total: d.cediMap[cd], pct: this.fmtPct(d.cediMap[cd]/d.total*100) }));
    this.tiendasState.tableRutas = d.topRegiones.map((rg:any) => {
      const info = d.regionMap[rg]; const pct = info.total ? info.cumple/info.total*100 : null;
      return { name: rg, total: info.total, pct: this.fmtPct(pct), pctClass: this.pillClass(pct) };
    });

    const labels = Object.keys(d.dist);
    this.makeDoughnut('chartTiendas', this.chartTiendasRef.nativeElement, labels, labels.map(l=>d.dist[l]), labels.map((l,i)=>{
      if(l==='CUMPLE') return '#2dd4bf'; if(l==='CUMPLE PARCIAL') return '#f2a93c'; if(l==='NO CUMPLE') return '#e2564f'; return this.PALETTE[i%this.PALETTE.length];
    }));
  }

  // ---------- PARSER 3: AGRICULTORES ----------
  parseAgri(wb: any) {
    const sheetName = this.findSheet(wb, ['RESUMEN']); if(!sheetName) throw new Error('No encontré la hoja "RESUMEN"');
    const rows = this.sheetRows(wb, sheetName); let headerRow = -1;
    for(let i=0;i<rows.length;i++){ if(rows[i] && String(rows[i][1]||'').trim().toUpperCase()==='AGRICULTOR'){ headerRow = i; break; } }
    if(headerRow<0) throw new Error('No encontré la tabla "Agricultor".');
    const weekRow = rows[headerRow+1]; const weekCols = [];
    for(let c=3;c<weekRow.length;c++){ if(weekRow[c]!==null && weekRow[c]!==undefined && weekRow[c]!=='') weekCols.push(c); }
    if(!weekCols.length) throw new Error('No encontré columnas de semana.');
    const nWeeks = Math.floor(weekCols.length/3); const llegadaCols = weekCols.slice(0,nWeeks); const lastWeekCol = llegadaCols[llegadaCols.length-1];
    const tiempoCols = weekCols.slice(nWeeks, 2*nWeeks); const plantaCols = weekCols.slice(2*nWeeks, 3*nWeeks);
    const lastTiempoCol = tiempoCols[tiempoCols.length-1]; const lastPlantaCol = plantaCols[plantaCols.length-1];
    const weekLabel = weekRow[lastWeekCol];
    const toPct = (v:any) => { if(v===null||v===undefined||v==='-') return null; if(typeof v==='string' && v.trim().endsWith('%')) return parseFloat(v); if(typeof v==='number') return v<=1 ? v*100 : v; return null; };
    const agricultores: any[] = [];
    let r = headerRow+2;
    while(r < rows.length){
      const name = rows[r] && rows[r][1]; if(!name) break; const totalViajes = rows[r][lastWeekCol];
      const pctLlegada = toPct(rows[r+3] ? rows[r+3][lastWeekCol] : null); const pctTiempo = toPct(rows[r+3] ? rows[r+3][lastTiempoCol] : null); const pctPlanta = toPct(rows[r+3] ? rows[r+3][lastPlantaCol] : null);
      if(typeof totalViajes === 'number'){ agricultores.push({name:String(name).trim(), viajes:totalViajes, pctLlegada, pctTiempo, pctPlanta}); } r += 4;
    }
    const totalViajes = agricultores.reduce((s,a)=>s+a.viajes,0);
    const weighted = (key:any) => { let num=0, den=0; agricultores.forEach((a:any)=>{ if(a[key]!==null){ num += a[key]*a.viajes; den += a.viajes; } }); return den ? num/den : null; };
    return { weekLabel: weekLabel || '', agricultores, totalViajes, avgLlegada: weighted('pctLlegada'), avgTiempo: weighted('pctTiempo'), avgPlanta: weighted('pctPlanta') };
  }

  renderAgri(d: any, sourceLabel: string) {
    const weekNum = String(d.weekLabel).replace(/[^0-9]/g,'') || d.weekLabel; const wLabel = 'Semana ' + weekNum;
    this.agriState.src = sourceLabel + ' · ' + wLabel + ' (más reciente disponible)';
    this.agriState.kpiSemanaLabel = 'semana ' + weekNum; this.agriState.chartLabel = '(S'+weekNum+')'; this.agriState.tableLabel = wLabel.toLowerCase();
    this.agriState.kpiTotal = d.totalViajes; this.agriState.kpiTotalSub = d.agricultores.length + ' agricultores';
    this.agriState.kpiLlegada = this.fmtPct(d.avgLlegada); this.agriState.kpiTiempo = this.fmtPct(d.avgTiempo); this.agriState.kpiPlanta = this.fmtPct(d.avgPlanta);
    
    this.agriState.table = d.agricultores.map((a:any) => ({
      name: a.name, total: a.viajes, llegada: this.fmtPct(a.pctLlegada), llegadaClass: this.pillClass(a.pctLlegada),
      tiempo: this.fmtPct(a.pctTiempo), tiempoClass: this.pillClass(a.pctTiempo), planta: this.fmtPct(a.pctPlanta), plantaClass: this.pillClass(a.pctPlanta)
    }));

    this.makeHBar('chartAgri', this.chartAgriRef.nativeElement, d.agricultores.map((a:any)=>a.name), d.agricultores.map((a:any)=>a.pctLlegada||0), d.agricultores.map((a:any)=>this.pillClass(a.pctLlegada)==='good' ? '#2dd4bf' : this.pillClass(a.pctLlegada)==='mid' ? '#f2a93c' : '#e2564f'));
  }
}