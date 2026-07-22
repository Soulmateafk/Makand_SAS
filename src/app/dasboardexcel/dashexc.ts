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
  PALETTE = ['#2dd4bf', '#f2a93c', '#e2564f', '#7aa2f7', '#c792ea', '#9ece6a'];
  MESES_ORDER = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // --- ESTADOS INICIALES ---
  makandState: any = {
    src: 'VIAJEROS_Y_TERCEROS_MAKAND.xlsx · JUNIO 2026 (mes más reciente completo)',
    status: '',
    statusClass: '',
    note: '258 viajes registrados en junio, repartidos entre 3 transportadoras. Los % de cumplimiento de llegada corresponden a la llegada del vehículo a planta según la hora citada.',
    kpiTotal: '258',
    kpiPeriodo: 'junio 2026',
    kpiCumple: '80.6%',
    kpiCajas: '115,432',
    kpiCajasSub: 'MAKAND+ARSITRANS+POLAR',
    kpiLider: 'MAKAND',
    kpiLiderSub: '51.2% de los viajes',
    table: [
      { name: 'MAKAND', color: '#2dd4bf', total: 132, pct: '51.2%', llegada: '93.2%', llegadaClass: 'good', cargue: '49.2%', cargueClass: 'mid' },
      { name: 'ARSITRANS', color: '#f2a93c', total: 78, pct: '30.2%', llegada: '71.8%', llegadaClass: 'good', cargue: '48.7%', cargueClass: 'mid' },
      { name: 'POLAR', color: '#e2564f', total: 48, pct: '18.6%', llegada: '60.4%', llegadaClass: 'mid', cargue: '89.6%', cargueClass: 'good' }
    ]
  };

  tiendasState: any = {
    src: 'ON_TIME_EN_TIENDAS.xlsx · JUNIO 2026 (mes más reciente completo)',
    status: '',
    statusClass: '',
    note: '929 registros de llegada/salida en ruta durante junio. El cumplimiento mide si el vehículo permaneció en el punto dentro del tiempo estándar asignado.',
    kpiTotal: '929',
    kpiPeriodo: 'junio 2026',
    kpiCumple: '33.8%',
    kpiCumpleSub: '314 registros',
    kpiParcial: '12.6%',
    kpiParcialSub: '117 registros',
    kpiNoCumple: '53.6%',
    kpiNoCumpleSub: '498 registros',
    rutasMesLabel: '(junio)',
    tableCedi: [
      { name: 'Tiendas', total: 602, pct: '64.8%' },
      { name: 'CEDI 2', total: 200, pct: '21.5%' },
      { name: 'CEDI ARA', total: 74, pct: '8.0%' },
      { name: 'CEDI 1', total: 53, pct: '5.7%' }
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
    src: 'ON_TIME_AGRICULTORES_2026_ACT.xlsx · JUNIO 2026 (mes más reciente completo)',
    status: '',
    statusClass: '',
    note: 'Cumplimiento de 478 registros durante junio, en tres momentos del proceso: llegada del transporte, tiempo en finca y llegada a planta. Los agricultores "FERRUCAS 1" a "FERRUCAS 10" se agrupan en uno solo (FERRUCAS).',
    kpiSemanaLabel: 'junio 2026',
    kpiTotal: '478',
    kpiTotalSub: '15 agricultores',
    kpiLlegada: '73.4%',
    kpiTiempo: '25.5%',
    kpiPlanta: '34.5%',
    chartLabel: '(junio)',
    tableLabel: 'junio 2026',
    table: [
      { name: 'Ferrucas', total: 172, llegada: '72.7%', llegadaClass: 'good', tiempo: '27.9%', tiempoClass: 'bad', planta: '25.3%', plantaClass: 'bad' },
      { name: 'Andrés Cadena', total: 42, llegada: '90.5%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '9.5%', plantaClass: 'bad' },
      { name: 'Wilson', total: 30, llegada: '90.0%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '60.0%', plantaClass: 'mid' },
      { name: 'José Tibaquicha', total: 30, llegada: '66.7%', llegadaClass: 'mid', tiempo: '23.3%', tiempoClass: 'bad', planta: '40.0%', plantaClass: 'mid' },
      { name: 'Mario Acevedo', total: 30, llegada: '90.0%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '60.0%', plantaClass: 'mid' },
      { name: 'Juan Pablo', total: 30, llegada: '90.0%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '80.0%', plantaClass: 'good' },
      { name: 'Jesús', total: 28, llegada: '100.0%', llegadaClass: 'good', tiempo: '100.0%', tiempoClass: 'good', planta: '85.7%', plantaClass: 'good' },
      { name: 'Gabriel', total: 28, llegada: '17.9%', llegadaClass: 'bad', tiempo: '35.7%', tiempoClass: 'bad', planta: '10.7%', plantaClass: 'bad' },
      { name: 'Severo', total: 27, llegada: '88.9%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '18.5%', plantaClass: 'bad' },
      { name: 'Diego', total: 25, llegada: '12.0%', llegadaClass: 'bad', tiempo: '80.0%', tiempoClass: 'good', planta: '8.0%', plantaClass: 'bad' },
      { name: 'Jorge Baracaldo', total: 16, llegada: '87.5%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '12.5%', plantaClass: 'bad' },
      { name: 'Georgeth', total: 9, llegada: '44.4%', llegadaClass: 'mid', tiempo: '88.9%', tiempoClass: 'good', planta: '22.2%', plantaClass: 'bad' },
      { name: 'Ana Cely', total: 9, llegada: '77.8%', llegadaClass: 'good', tiempo: '0.0%', tiempoClass: 'bad', planta: '77.8%', plantaClass: 'good' }
    ]
  };

  constructor() {
  Chart.defaults.color = '#e2e8f0'; 
  Chart.defaults.font.family = "'IBM Plex Mono', monospace";
  Chart.defaults.font.size = 11;
}

  ngAfterViewInit(): void {
    this.makeDoughnut('chartTransp', this.chartTranspRef.nativeElement, ['MAKAND', 'ARSITRANS', 'POLAR'], [132, 78, 48], ['#2dd4bf', '#f2a93c', '#e2564f']);
    this.makeStackedBar('chartAlmacen', this.chartAlmacenRef.nativeElement, ['D1', 'ARA', 'ÉXITO'], [
      { label: 'MAKAND', data: [96, 10, 26], backgroundColor: '#2dd4bf' },
      { label: 'ARSITRANS', data: [32, 32, 14], backgroundColor: '#f2a93c' },
      { label: 'POLAR', data: [48, 0, 0], backgroundColor: '#e2564f' }
    ]);
    this.makeDoughnut('chartTiendas', this.chartTiendasRef.nativeElement, ['Cumple', 'Cumple parcial', 'No cumple'], [314, 117, 498], ['#2dd4bf', '#f2a93c', '#e2564f']);
    this.makeHBar('chartAgri', this.chartAgriRef.nativeElement,
      ['Ferrucas', 'Andrés Cadena', 'Wilson', 'José T.', 'Mario Acevedo', 'Juan Pablo', 'Jesús', 'Gabriel', 'Severo', 'Diego', 'Jorge Baracaldo', 'Georgeth', 'Ana Cely'],
      [72.7, 90.5, 90.0, 66.7, 90.0, 90.0, 100.0, 17.9, 88.9, 12.0, 87.5, 44.4, 77.8],
      ['#2dd4bf', '#2dd4bf', '#2dd4bf', '#f2a93c', '#2dd4bf', '#2dd4bf', '#2dd4bf', '#e2564f', '#2dd4bf', '#e2564f', '#2dd4bf', '#f2a93c', '#2dd4bf']);
  }

  // ---------- HELPERS ----------
  pillClass(pct: any): string {
    if (pct === null || pct === undefined || isNaN(pct)) return 'mid';
    if (pct >= 70) return 'good';
    if (pct >= 40) return 'mid';
    return 'bad';
  }

  fmtPct(pct: any): string {
    return (pct === null || pct === undefined || isNaN(pct)) ? '—' : pct.toFixed(1) + '%';
  }

  findSheet(wb: any, patterns: string[]): string | null {
    const names = wb.SheetNames;
    for (const p of patterns) {
      const hit = names.find((n: string) => n.toUpperCase().replace(/\s+/g, ' ').trim().includes(p));
      if (hit) return hit;
    }
    return null;
  }

  sheetRows(wb: any, name: string): any[] {
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true });
  }

  colIndex(header: any[], name: string): number {
    return header.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
  }

  pickLatestCompleteMonth(counts: any): string | null {
    const months = Object.keys(counts).filter(m => this.MESES_ORDER.includes(m));
    months.sort((a, b) => this.MESES_ORDER.indexOf(a) - this.MESES_ORDER.indexOf(b));
    if (!months.length) return null;
    const maxCount = Math.max(...months.map(m => counts[m]));
    let chosen = months[0];
    for (const m of months) {
      if (counts[m] >= 0.5 * maxCount) chosen = m;
    }
    return chosen;
  }

  // ---------- CREADORES DE GRÁFICAS ----------
  makeDoughnut(id: string, canvasRef: any, labels: any, data: any, colors: any) {
  if (this.charts[id]) this.charts[id].destroy();
  this.charts[id] = new Chart(canvasRef, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: colors, borderColor: '#171e27', borderWidth: 3 }]
    },
    options: {
      responsive: true,           
      maintainAspectRatio: false,  
      plugins: {
        legend: {
          position: 'bottom',
          labels: { boxWidth: 12, padding: 20 }
        }
      },
      cutout: '60%' 
    }
  });
}

  makeStackedBar(id: string, canvasRef: any, labels: any, datasets: any) {
    if (this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(canvasRef, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, grid: { color: '#2a3340' } }
        },
        plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, padding: 14 } } }
      }
    });
  }

  makeHBar(id: string, canvasRef: any, labels: any, data: any, colors: any, max: number = 100) {
    if (this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(canvasRef, {
      type: 'bar',
      data: {
        labels,
        datasets: [{ data, backgroundColor: colors }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        scales: {
          x: { max: max, grid: { color: '#2a3340' } },
          y: { grid: { display: false } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }

  // ---------- OPERACIONES DE CARGA (UPLOAD) ----------
  setStatus(key: string, msg: string, cls: string = '') {
    const state = key === 'makand' ? this.makandState : key === 'tiendas' ? this.tiendasState : this.agriState;
    state.status = msg;
    state.statusClass = cls ? `upload-status ${cls}` : 'upload-status';
  }

  handleUpload(evt: any, kind: string) {
    const file = evt.target.files[0];
    if (!file) return;

    this.setStatus(kind, 'Leyendo ' + file.name + '…');
    const reader = new FileReader();

    reader.onload = (e: any) => {
      try {
        const parseDates = (kind !== 'makand' && kind !== 'tiendas');
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: parseDates });

        if (kind === 'makand') {
          const d = this.parseMakand(wb);
          this.renderMakand(d, file.name);
        } else if (kind === 'tiendas') {
          const d = this.parseTiendas(wb);
          this.renderTiendas(d, file.name);
        } else {
          const d = this.parseAgri(wb);
          this.renderAgri(d, file.name);
        }
        this.setStatus(kind, '✓ actualizado ' + new Date().toLocaleString('es-CO'), 'ok');
      } catch (err: any) {
        this.setStatus(kind, 'Error: ' + err.message, 'err');
      }
    };

    reader.onerror = () => this.setStatus(kind, 'No se pudo leer el archivo.', 'err');
    reader.readAsArrayBuffer(file);
  }

  // ---------- PARSER 1: MAKAND ----------
  parseMakand(wb: any) {
    const sheetName = this.findSheet(wb, ['DT VIAJEROS']);
    if (!sheetName) throw new Error('No encontré la hoja "DT VIAJEROS".');

    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iMes = this.colIndex(header, 'MES'),
          iTransp = this.colIndex(header, 'TRANSPORTE'),
          iAlmacen = this.colIndex(header, 'ALMACEN'),
          iLlegada = this.colIndex(header, 'CUMPLIMIENTO LLEGADA VEHICULO'),
          iCargue = this.colIndex(header, 'CUMPLIMIENTO CARGUE VEHICULO'),
          iCajas = this.colIndex(header, 'TOTAL CAJAS');

    if (iMes < 0 || iTransp < 0) throw new Error('No encontré las columnas MES/TRANSPORTE esperadas.');

    const data = rows.slice(1).filter(r => r && r[iMes] && r[iTransp]);
    const counts: any = {};
    data.forEach(r => {
      const m = String(r[iMes]).trim().toUpperCase();
      counts[m] = (counts[m] || 0) + 1;
    });

    const month = this.pickLatestCompleteMonth(counts) || 'JUNIO';
    const rowsMonth = data.filter(r => String(r[iMes]).trim().toUpperCase() === month);

    const transpMap: any = {};
    const almacenSet = new Set<string>();

    rowsMonth.forEach(r => {
      const t = String(r[iTransp]).trim().toUpperCase();
      if (!transpMap[t]) transpMap[t] = { total: 0, llegadaOk: 0, cargueOk: 0, cajas: 0, almacenes: {} };
      transpMap[t].total++;

      if (iLlegada >= 0 && String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE') transpMap[t].llegadaOk++;
      if (iCargue >= 0 && String(r[iCargue]).trim().toUpperCase() === 'CUMPLE') transpMap[t].cargueOk++;
      
      if (iCajas >= 0) {
        const v = parseFloat(r[iCajas]);
        if (!isNaN(v)) transpMap[t].cajas += v;
      }

      const alm = iAlmacen >= 0 ? String(r[iAlmacen] || '').trim() : 'N/D';
      almacenSet.add(alm);
      transpMap[t].almacenes[alm] = (transpMap[t].almacenes[alm] || 0) + 1;
    });

    const total = rowsMonth.length;
    const transportadoras = Object.keys(transpMap).sort((a, b) => transpMap[b].total - transpMap[a].total);
    const overallLlegada = total ? (Object.keys(transpMap).reduce((s: number, k: string) => s + (transpMap[k].llegadaOk || 0), 0) / total * 100) : 0;
    const totalCajas = Object.keys(transpMap).reduce((s: number, k: string) => s + transpMap[k].cajas, 0);
    const lider = transportadoras[0] || 'N/D';

    return {
      month,
      total,
      transportadoras,
      transpMap,
      almacenes: [...almacenSet],
      overallLlegada,
      totalCajas,
      lider,
      liderPct: total ? (transpMap[lider].total / total * 100) : 0
    };
  }

  renderMakand(d: any, sourceLabel: string) {
    this.makandState.src = sourceLabel + ' · ' + d.month + ' (mes más reciente completo)';
    this.makandState.note = `${d.total} viajes registrados en ${d.month.toLowerCase()}, repartidos entre ${d.transportadoras.length} transportadora(s).`;
    this.makandState.kpiTotal = d.total;
    this.makandState.kpiPeriodo = d.month.toLowerCase();
    this.makandState.kpiCumple = this.fmtPct(d.overallLlegada);
    this.makandState.kpiCajas = Math.round(d.totalCajas).toLocaleString('es-CO');
    this.makandState.kpiLider = d.lider;
    this.makandState.kpiLiderSub = this.fmtPct(d.liderPct) + ' de los viajes';
    
    this.makandState.table = d.transportadoras.map((t: any, i: number) => {
      const m = d.transpMap[t];
      const pctTotal = m.total / d.total * 100;
      const pctLlegada = m.total ? m.llegadaOk / m.total * 100 : 0;
      const pctCargue = m.total ? m.cargueOk / m.total * 100 : 0;
      return {
        name: t,
        color: this.PALETTE[i % this.PALETTE.length],
        total: m.total,
        pct: this.fmtPct(pctTotal),
        llegada: this.fmtPct(pctLlegada),
        llegadaClass: this.pillClass(pctLlegada),
        cargue: this.fmtPct(pctCargue),
        cargueClass: this.pillClass(pctCargue)
      };
    });

    this.makeDoughnut('chartTransp', this.chartTranspRef.nativeElement, d.transportadoras, d.transportadoras.map((t: any) => d.transpMap[t].total), d.transportadoras.map((_: any, i: number) => this.PALETTE[i % this.PALETTE.length]));
    this.makeStackedBar('chartAlmacen', this.chartAlmacenRef.nativeElement, d.almacenes, d.transportadoras.map((t: any, i: number) => ({
      label: t,
      backgroundColor: this.PALETTE[i % this.PALETTE.length],
      data: d.almacenes.map((a: any) => d.transpMap[t].almacenes[a] || 0)
    })));
  }

  // ---------- PARSER 2: TIENDAS ----------
  parseTiendas(wb: any) {
    const sheetName = this.findSheet(wb, ['DETALLE REGIONES']);
    if (!sheetName) throw new Error('No encontré la hoja "Detalle regiones"');

    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iMes = this.colIndex(header, 'MES'),
          iCump = this.colIndex(header, 'CUMPLIMIENTO'),
          iRegion = this.colIndex(header, 'REGIÓN'),
          iCedi = this.colIndex(header, 'CEDI');

    if (iMes < 0 || iCump < 0) throw new Error('No encontré las columnas MES/CUMPLIMIENTO.');

    const data = rows.slice(1).filter(r => r && r[iMes]);
    const counts: any = {};
    data.forEach(r => {
      const m = String(r[iMes]).trim().toUpperCase();
      counts[m] = (counts[m] || 0) + 1;
    });

    const month = this.pickLatestCompleteMonth(counts) || 'JUNIO';
    const rowsMonth = data.filter(r => String(r[iMes]).trim().toUpperCase() === month);

    const dist: any = {};
    const cediMap: any = {};
    const regionMap: any = {};

    rowsMonth.forEach(r => {
      const c = iCump >= 0 ? String(r[iCump] || '').trim().toUpperCase() : 'N/D';
      dist[c] = (dist[c] || 0) + 1;

      if (iCedi >= 0) {
        const cd = String(r[iCedi] || '').trim() || 'N/D';
        cediMap[cd] = (cediMap[cd] || 0) + 1;
      }
      if (iRegion >= 0) {
        const rg = String(r[iRegion] || '').trim() || 'N/D';
        if (!regionMap[rg]) regionMap[rg] = { total: 0, cumple: 0 };
        regionMap[rg].total++;
        if (c === 'CUMPLE') regionMap[rg].cumple++;
      }
    });

    const total = rowsMonth.length;
    const topRegiones = Object.keys(regionMap).sort((a, b) => regionMap[b].total - regionMap[a].total).slice(0, 5);
    const cedis = Object.keys(cediMap).sort((a, b) => cediMap[b] - cediMap[a]);

    return { month, total, dist, cediMap, cedis, regionMap, topRegiones };
  }

  renderTiendas(d: any, sourceLabel: string) {
    this.tiendasState.src = sourceLabel + ' · ' + d.month + ' (mes más reciente completo)';
    this.tiendasState.note = `${d.total} registros de llegada/salida en ruta durante ${d.month.toLowerCase()}.`;
    this.tiendasState.rutasMesLabel = '(' + d.month.toLowerCase() + ')';

    const cumple = d.dist['CUMPLE'] || 0,
          parcial = d.dist['CUMPLE PARCIAL'] || 0,
          noCumple = d.dist['NO CUMPLE'] || 0;

    this.tiendasState.kpiTotal = d.total;
    this.tiendasState.kpiPeriodo = d.month.toLowerCase();
    this.tiendasState.kpiCumple = this.fmtPct(cumple / d.total * 100);
    this.tiendasState.kpiCumpleSub = cumple + ' registros';
    this.tiendasState.kpiParcial = this.fmtPct(parcial / d.total * 100);
    this.tiendasState.kpiParcialSub = parcial + ' registros';
    this.tiendasState.kpiNoCumple = this.fmtPct(noCumple / d.total * 100);
    this.tiendasState.kpiNoCumpleSub = noCumple + ' registros';

    this.tiendasState.tableCedi = d.cedis.map((cd: any) => ({
      name: cd,
      total: d.cediMap[cd],
      pct: this.fmtPct(d.cediMap[cd] / d.total * 100)
    }));

    this.tiendasState.tableRutas = d.topRegiones.map((rg: any) => {
      const info = d.regionMap[rg];
      const pct = info.total ? info.cumple / info.total * 100 : null;
      return {
        name: rg,
        total: info.total,
        pct: this.fmtPct(pct),
        pctClass: this.pillClass(pct)
      };
    });

    const labels = Object.keys(d.dist);
    this.makeDoughnut('chartTiendas', this.chartTiendasRef.nativeElement, labels, labels.map(l => d.dist[l]), labels.map((l, i) => {
      if (l === 'CUMPLE') return '#2dd4bf';
      if (l === 'CUMPLE PARCIAL') return '#f2a93c';
      if (l === 'NO CUMPLE') return '#e2564f';
      return this.PALETTE[i % this.PALETTE.length];
    }));
  }

  // ---------- PARSER 3: AGRICULTORES (MENSUAL) ----------
  parseAgri(wb: any) {
    const sheetName = this.findSheet(wb, ['TIEMPO AGRICULTORES']);
    if (!sheetName) throw new Error('No encontré la hoja "TIEMPO AGRICULTORES".');

    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iMes = this.colIndex(header, 'MES');
    const iAgricultor = this.colIndex(header, 'AGRICULTOR');
    const iLlegada = this.colIndex(header, 'CUMPLIMIENTO LLEGADA AGRICULTOR');
    const iTiempo = this.colIndex(header, 'CUMPLIMIENTO TIEMPO EN AGRICULTOR');
    const iPlanta = this.colIndex(header, 'CUMPLIMIENTO LLEGADA A PLANTA');

    if (iMes < 0 || iAgricultor < 0) throw new Error('No encontré las columnas MES/AGRICULTOR esperadas en "TIEMPO AGRICULTORES".');

    const EXCLUIR = ['LECHUGAS DEL DIA', 'LECHUGAS DEL DÍA'];
    const normName = (s: any) => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    const displayName = (raw: any): string => {
      const n = normName(raw);
      if (n.startsWith('FERRUCAS')) return 'FERRUCAS';
      return String(raw).trim();
    };

    const data = rows.slice(1).filter(r => r && r[iMes] && r[iAgricultor] && !EXCLUIR.includes(normName(r[iAgricultor])));

    const counts: any = {};
    data.forEach(r => {
      const m = String(r[iMes]).trim().toUpperCase();
      counts[m] = (counts[m] || 0) + 1;
    });
    const month = this.pickLatestCompleteMonth(counts) || 'JUNIO';
    const rowsMonth = data.filter(r => String(r[iMes]).trim().toUpperCase() === month);

    // ---- por agricultor (agrupando FERRUCAS N -> FERRUCAS) ----
    const agriMap: any = {};
    rowsMonth.forEach(r => {
      const name = displayName(r[iAgricultor]);
      if (!agriMap[name]) agriMap[name] = { viajes: 0, llegadaOk: 0, llegadaTot: 0, tiempoOk: 0, tiempoTot: 0, plantaOk: 0, plantaTot: 0 };
      agriMap[name].viajes++;
      const vL = iLlegada >= 0 ? String(r[iLlegada] || '').trim().toUpperCase() : '';
      if (vL === 'CUMPLE' || vL === 'NO CUMPLE') { agriMap[name].llegadaTot++; if (vL === 'CUMPLE') agriMap[name].llegadaOk++; }
      const vT = iTiempo >= 0 ? String(r[iTiempo] || '').trim().toUpperCase() : '';
      if (vT === 'CUMPLE' || vT === 'NO CUMPLE') { agriMap[name].tiempoTot++; if (vT === 'CUMPLE') agriMap[name].tiempoOk++; }
      const vP = iPlanta >= 0 ? String(r[iPlanta] || '').trim().toUpperCase() : '';
      if (vP === 'CUMPLE' || vP === 'NO CUMPLE') { agriMap[name].plantaTot++; if (vP === 'CUMPLE') agriMap[name].plantaOk++; }
    });

    const agricultores = Object.keys(agriMap)
      .sort((a, b) => agriMap[b].viajes - agriMap[a].viajes)
      .map(name => {
        const m = agriMap[name];
        return {
          name,
          viajes: m.viajes,
          pctLlegada: m.llegadaTot ? m.llegadaOk / m.llegadaTot * 100 : null,
          pctTiempo: m.tiempoTot ? m.tiempoOk / m.tiempoTot * 100 : null,
          pctPlanta: m.plantaTot ? m.plantaOk / m.plantaTot * 100 : null
        };
      });

    const totalViajes = rowsMonth.length;

    const statFrom = (field: number) => {
      let ok = 0, tot = 0;
      rowsMonth.forEach(r => {
        const v = String(r[field] || '').trim().toUpperCase();
        if (v === 'CUMPLE' || v === 'NO CUMPLE') { tot++; if (v === 'CUMPLE') ok++; }
      });
      return { avg: tot ? ok / tot * 100 : null, count: ok };
    };

    return {
      month,
      agricultores,
      totalViajes,
      llegada: statFrom(iLlegada),
      tiempo: statFrom(iTiempo),
      planta: statFrom(iPlanta)
    };
  }

  renderAgri(d: any, sourceLabel: string) {
    this.agriState.src = sourceLabel + ' · ' + d.month + ' (mes más reciente completo)';
    this.agriState.note = `Cumplimiento de ${d.totalViajes} registros durante ${d.month.toLowerCase()}, en tres momentos del proceso: llegada del transporte, tiempo en finca y llegada a planta. Los agricultores "FERRUCAS 1" a "FERRUCAS 10" se agrupan en uno solo (FERRUCAS).`;
    this.agriState.kpiSemanaLabel = d.month.toLowerCase();
    this.agriState.tableLabel = d.month.toLowerCase();
    this.agriState.chartLabel = '(' + d.month.toLowerCase() + ')';
    this.agriState.kpiTotal = d.totalViajes;
    this.agriState.kpiTotalSub = d.agricultores.length + ' agricultores';
    this.agriState.kpiLlegada = this.fmtPct(d.llegada.avg);
    this.agriState.kpiTiempo = this.fmtPct(d.tiempo.avg);
    this.agriState.kpiPlanta = this.fmtPct(d.planta.avg);

    this.agriState.table = d.agricultores.map((a: any) => ({
      name: a.name,
      total: a.viajes,
      llegada: this.fmtPct(a.pctLlegada),
      llegadaClass: this.pillClass(a.pctLlegada),
      tiempo: this.fmtPct(a.pctTiempo),
      tiempoClass: this.pillClass(a.pctTiempo),
      planta: this.fmtPct(a.pctPlanta),
      plantaClass: this.pillClass(a.pctPlanta)
    }));

    const labels = d.agricultores.map((a: any) => a.name);
    const chartData = d.agricultores.map((a: any) => a.pctLlegada || 0);
    const colors = d.agricultores.map((a: any) => {
      const cls = this.pillClass(a.pctLlegada);
      return cls === 'good' ? '#2dd4bf' : cls === 'mid' ? '#f2a93c' : '#e2564f';
    });

    this.makeHBar('chartAgri', this.chartAgriRef.nativeElement, labels, chartData, colors);
  }
}