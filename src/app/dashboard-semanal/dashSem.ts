import { Component, AfterViewInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';
import { Router } from '@angular/router'; 

// Constantes globales de configuración de la paleta y orden cronológico
const PALETTE: string[] = ['#2F6F4E', '#BE8A2E', '#B5453A', '#3E6C99', '#7A5C99', '#5C8A6E'];
const MESES_ORDER: string[] = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const CURRENT_MONTH_NAME: string = MESES_ORDER[new Date().getMonth()];

@Component({
  selector: 'app-dash-semanal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashSem.html',
  styleUrl: './dashSem.css'
})
export class DashSemanalComponent implements AfterViewInit {
  // Almacenamiento privado de las instancias de los gráficos activos
  private charts: { [key: string]: any } = {};

  constructor(private router: Router) {
    // Configuración de los estilos por defecto de Chart.js
    Chart.defaults.color = '#8B7E6C';
    Chart.defaults.font.family = "'Space Mono', monospace";
    Chart.defaults.font.size = 11;

    // Vinculación explícita al objeto global window para mantener compatibilidad con onchange="handleUpload(...)" del HTML
    (window as any).handleUpload = (evt: any, kind: string) => this.handleUpload(evt, kind);
  }

  ngAfterViewInit(): void {
    // Ejecución inicial de los gráficos y tablas estáticas con el histórico base
    this.inicializarCuadrosControl();
  }

  // ==========================================
  // CONSTRUCTORES DE GRÁFICOS (CHART UTILS)
  // ==========================================

  public makeDoughnut(id: string, labels: string[], data: number[], colors: string[]): void {
    if (this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(document.getElementById(id) as HTMLCanvasElement, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{ data, backgroundColor: colors, borderColor: '#FFFEFB', borderWidth: 3 }]
      },
      options: {
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, padding: 16 } }
        },
        cutout: '62%'
      }
    });
  }

  public makeStackedBar(id: string, labels: string[], datasets: any[]): void {
    if (this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(document.getElementById(id) as HTMLCanvasElement, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, grid: { color: '#E4DAC3' } }
        },
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, padding: 14 } }
        }
      }
    });
  }

  public makeHBar(id: string, labels: string[], data: number[], colors: string[], max?: number): void {
    if (this.charts[id]) this.charts[id].destroy();
    this.charts[id] = new Chart(document.getElementById(id) as HTMLCanvasElement, {
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
          x: { max: max || 100, grid: { color: '#E4DAC3' } },
          y: { grid: { display: false } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }

  public makeLine(id: string, labels: any[], datasets: any[], dualAxis?: boolean): void {
    if (this.charts[id]) this.charts[id].destroy();
    const scales = dualAxis ? {
      x: { grid: { display: false } },
      y: { position: 'left' as const, grid: { color: '#E4DAC3' }, title: { display: true, text: 'Viajes / registros', font: { size: 10 } } },
      y1: { position: 'right' as const, min: 0, max: 100, grid: { display: false }, title: { display: true, text: '% cumplimiento', font: { size: 10 } } }
    } : {
      x: { grid: { display: false } },
      y: { min: 0, max: 100, grid: { color: '#E4DAC3' } }
    };

    this.charts[id] = new Chart(document.getElementById(id) as HTMLCanvasElement, {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales,
        plugins: {
          legend: { position: 'bottom' as const, labels: { boxWidth: 10, padding: 14 } }
        }
      }
    });
  }

  // ==========================================
  // HELPERS DE FORMATO E INDICADORES LOGÍSTICOS
  // ==========================================

  public pillClass(pct: number | null | undefined): string {
    if (pct === null || pct === undefined || isNaN(pct)) return 'mid';
    if (pct >= 70) return 'good';
    if (pct >= 40) return 'mid';
    return 'bad';
  }

  public fmtPct(pct: number | null | undefined): string {
    return (pct === null || pct === undefined || isNaN(pct)) ? '—' : pct.toFixed(1) + '%';
  }

  public calcularEfectividadSustraccion(total: number, fallas: number): number {
    if (!total || total <= 0) return 0;
    const porcentajeDescuento = (fallas / total) * 100;
    return Math.max(0, 100 - porcentajeDescuento);
  }

  // ==========================================
  // HELPERS DE EXTRACCIÓN Y LECTURA EXCEL (XLSX)
  // ==========================================

  private findSheet(wb: any, patterns: string[]): string | null {
    const names = wb.SheetNames;
    for (const p of patterns) {
      const hit = names.find((n: string) => n.toUpperCase().replace(/\s+/g, ' ').trim().includes(p));
      if (hit) return hit;
    }
    return null;
  }

  public sheetRows(wb: any, name: string): any[] {
    const ws = wb.Sheets[name];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  public colIndex(header: any[], names: string | string[]): number {
    const candidates = Array.isArray(names) ? names : [names];
    for (const name of candidates) {
      const idx = header.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
      if (idx >= 0) return idx;
    }
    return -1;
  }

  public pickLatestWeek(counts: any): number | null {
    const weeks = Object.keys(counts).map(Number).filter(w => !isNaN(w)).sort((a, b) => a - b);
    if (!weeks.length) return null;
    return weeks[weeks.length - 1];
  }

  // ==========================================
  // PARSER 1: VIAJEROS Y TERCEROS MAKAND
  // ==========================================

  public parseMakand(wb: any): any {
    const sheetName = this.findSheet(wb, ['DT VIAJEROS']);
    if (!sheetName) throw new Error('No encontré la hoja "DT VIAJEROS." en este archivo.');
    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iSem = this.colIndex(header, ['SEMANAS', 'SEMANA']),
          iTransp = this.colIndex(header, 'TRANSPORTE'),
          iAlmacen = this.colIndex(header, 'ALMACEN'),
          iLlegada = this.colIndex(header, 'CUMPLIMIENTO LLEGADA VEHICULO'),
          iCargue = this.colIndex(header, 'CUMPLIMIENTO CARGUE VEHICULO'),
          iCajas = this.colIndex(header, 'TOTAL CAJAS');

    if (iSem < 0 || iTransp < 0) throw new Error('No encontré las columnas SEMANA/TRANSPORTE esperadas.');

    const data = rows.slice(1).filter(r => r && r[iSem] !== null && r[iTransp]);
    const counts: any = {};
    data.forEach(r => {
      const w = Number(r[iSem]);
      counts[w] = (counts[w] || 0) + 1;
    });

    const week = this.pickLatestWeek(counts);
    const rowsWeek = data.filter(r => Number(r[iSem]) === week);

    const transpMap: any = {};
    const almacenSet = new Set<string>();

    rowsWeek.forEach(r => {
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

    const total = rowsWeek.length;
    const transportadoras = Object.keys(transpMap).sort((a, b) => transpMap[b].total - transpMap[a].total);
    const overallLlegada = total ? (Object.values(transpMap).reduce((s: number, t: any) => s + t.llegadaOk, 0) / total * 100) : 0;
    const totalCajas = Object.values(transpMap).reduce((s: number, t: any) => s + t.cajas, 0);
    const lider = transportadoras[0];

    const allWeeks = [...new Set(data.map(r => Number(r[iSem])))].sort((a, b) => a - b);
    const allTransportadoras = [...new Set(data.map(r => String(r[iTransp]).trim().toUpperCase()))];
    const trendTotal: number[] = [];
    const trendLlegada: (number | null)[] = [];
    const trendPorTransp: any = {};
    allTransportadoras.forEach(t => trendPorTransp[t] = []);

    allWeeks.forEach(w => {
      const rw = data.filter(r => Number(r[iSem]) === w);
      trendTotal.push(rw.length);
      const ok = iLlegada >= 0 ? rw.filter(r => String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE').length : 0;
      trendLlegada.push(rw.length ? ok / rw.length * 100 : null);
      allTransportadoras.forEach(t => {
        trendPorTransp[t].push(rw.filter(r => String(r[iTransp]).trim().toUpperCase() === t).length);
      });
    });

    const trend = { weeks: allWeeks, total: trendTotal, llegada: trendLlegada, porTransportadora: trendPorTransp, transportadoras: allTransportadoras };
    const iMes = this.colIndex(header, 'MES');
    let monthly = null;
    if (iMes >= 0) {
      const mesesPresentes = [...new Set(data.map(r => String(r[iMes]).trim().toUpperCase()))]
        .filter(m => MESES_ORDER.includes(m))
        .sort((a, b) => MESES_ORDER.indexOf(a) - MESES_ORDER.indexOf(b));
      const totalPorMes: number[] = [], llegadaPorMes: (number | null)[] = [];
      mesesPresentes.forEach(m => {
        const rm = data.filter(r => String(r[iMes]).trim().toUpperCase() === m);
        totalPorMes.push(rm.length);
        const ok = iLlegada >= 0 ? rm.filter(r => String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE').length : 0;
        llegadaPorMes.push(rm.length ? ok / rm.length * 100 : null);
      });
      monthly = { meses: mesesPresentes, total: totalPorMes, llegada: llegadaPorMes };
    }

    return { week, total, transportadoras, transpMap, almacenes: [...almacenSet], overallLlegada, totalCajas, lider, liderPct: total ? (transpMap[lider].total / total * 100) : 0, trend, monthly };
  }

  public renderMakand(d: any, sourceLabel: string): void {
    document.getElementById('tabMakand')!.textContent = 'SEM ' + d.week;
    document.getElementById('srcMakand')!.textContent = sourceLabel;
    document.getElementById('noteMakand')!.textContent = `${d.total} viajes registrados en la semana ${d.week}, repartidos entre ${d.transportadoras.length} transportadora(s). Los % de cumplimiento de llegada corresponden a la llegada del vehículo a planta según la hora citada.`;
    
    document.getElementById('kpiMakandTotal')!.textContent = d.total;
    document.getElementById('kpiMakandPeriodo')!.textContent = 'semana ' + d.week;
    document.getElementById('kpiMakandCumple')!.textContent = this.fmtPct(d.overallLlegada);
    document.getElementById('kpiMakandCajas')!.textContent = Math.round(d.totalCajas).toLocaleString('es-CO');
    document.getElementById('kpiMakandCajasSub')!.textContent = d.transportadoras.join('+');
    document.getElementById('kpiMakandLider')!.textContent = d.lider;
    document.getElementById('kpiMakandLiderSub')!.textContent = this.fmtPct(d.liderPct) + ' de los viajes';

    const tbody = document.getElementById('tableMakand')!;
    tbody.innerHTML = d.transportadoras.map((t: string, i: number) => {
      const m = d.transpMap[t];
      const pctTotal = m.total / d.total * 100;
      const pctLlegada = m.total ? m.llegadaOk / m.total * 100 : 0;
      const pctCargue = m.total ? m.cargueOk / m.total * 100 : 0;
      return `<tr>
        <td class="name"><span class="legend-dot" style="background:${PALETTE[i % PALETTE.length]}"></span>${t}</td>
        <td>${m.total}</td><td>${this.fmtPct(pctTotal)}</td>
        <td><span class="pill ${this.pillClass(pctLlegada)}">${this.fmtPct(pctLlegada)}</span></td>
        <td><span class="pill ${this.pillClass(pctCargue)}">${this.fmtPct(pctCargue)}</span></td>
      </tr>`;
    }).join('');

    this.makeDoughnut('chartTransp', d.transportadoras, d.transportadoras.map((t: string) => d.transpMap[t].total), d.transportadoras.map((_: any, i: number) => PALETTE[i % PALETTE.length]));
    document.getElementById('makandLatestLabel')!.textContent = '(semana ' + d.week + ')';
    document.getElementById('makandLatestLabel2')!.textContent = '(semana ' + d.week + ')';

    const trendLabels = d.trend.weeks.map((w: number) => 'S' + w);
    this.makeLine('chartMakandTrend', trendLabels, [
      { label: 'Viajes totales', data: d.trend.total, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3, yAxisID: 'y' },
      { label: '% Cumplimiento llegada', data: d.trend.llegada, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3, borderDash: [4, 3], yAxisID: 'y1' }
    ], true);

    if (d.monthly) {
      this.renderMonthly('chartMakandMes', 'tableMakandMes', d.monthly.meses, d.monthly.total, d.monthly.llegada, 'Viajes', '% Cumple llegada');
    }
  }

  // ==========================================
  // PARSER 2: ON TIME EN TIENDAS
  // ==========================================

  public parseTiendas(wb: any): any {
    const sheetName = this.findSheet(wb, ['DETALLE REGIONES']);
    if (!sheetName) throw new Error('No encontré la hoja "Detalle regiones" en este archivo.');
    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iSem = this.colIndex(header, ['SEMANA', 'SEMANAS']),
          iCump = this.colIndex(header, 'CUMPLIMIENTO'),
          iRegion = this.colIndex(header, 'REGIÓN'),
          iCedi = this.colIndex(header, 'CEDI');

    if (iSem < 0 || iCump < 0) throw new Error('No encontré las columnas SEMANA/CUMPLIMIENTO esperadas.');

    const data = rows.slice(1).filter(r => r && r[iSem] !== null);
    const counts: any = {};
    data.forEach(r => {
      const w = Number(r[iSem]);
      counts[w] = (counts[w] || 0) + 1;
    });
    const week = this.pickLatestWeek(counts);
    const rowsWeek = data.filter(r => Number(r[iSem]) === week);

    const dist: any = {};
    const cediMap: any = {};
    const regionMap: any = {};
    rowsWeek.forEach(r => {
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

    const total = rowsWeek.length;
    const topRegiones = Object.keys(regionMap).sort((a, b) => regionMap[b].total - regionMap[a].total).slice(0, 5);
    const cedis = Object.keys(cediMap).sort((a, b) => cediMap[b] - cediMap[a]);

    const allWeeks = [...new Set(data.map(r => Number(r[iSem])))].sort((a, b) => a - b);
    const trendCumple: (number | null)[] = [], trendParcial: (number | null)[] = [], trendNoCumple: (number | null)[] = [], trendTotal: number[] = [];

    allWeeks.forEach(w => {
      const rw = data.filter(r => Number(r[iSem]) === w);
      const t = rw.length;
      trendTotal.push(t);
      const c = rw.filter(r => String(r[iCump] || '').trim().toUpperCase() === 'CUMPLE').length;
      const p = rw.filter(r => String(r[iCump] || '').trim().toUpperCase() === 'CUMPLE PARCIAL').length;
      const n = rw.filter(r => String(r[iCump] || '').trim().toUpperCase() === 'NO CUMPLE').length;
      trendCumple.push(t ? c / t * 100 : null);
      trendParcial.push(t ? p / t * 100 : null);
      trendNoCumple.push(t ? n / t * 100 : null);
    });

    return {
      week, total, dist, cediMap, cedis, regionMap, topRegiones,
      trend: { weeks: allWeeks, cumple: trendCumple, parcial: trendParcial, noCumple: trendNoCumple, total: trendTotal },
      monthly: (() => {
        const iMes = this.colIndex(header, 'MES');
        if (iMes < 0) return null;
        const mesesPresentes = [...new Set(data.map(r => String(r[iMes]).trim().toUpperCase()))]
          .filter(m => MESES_ORDER.includes(m))
          .sort((a, b) => MESES_ORDER.indexOf(a) - MESES_ORDER.indexOf(b));
        const totalPorMes: number[] = [], cumplePorMes: (number | null)[] = [];
        mesesPresentes.forEach(m => {
          const rm = data.filter(r => String(r[iMes]).trim().toUpperCase() === m);
          totalPorMes.push(rm.length);
          const ok = rm.filter(r => String(r[iCump] || '').trim().toUpperCase() === 'CUMPLE').length;
          cumplePorMes.push(rm.length ? ok / rm.length * 100 : null);
        });
        return { meses: mesesPresentes, total: totalPorMes, cumple: cumplePorMes };
      })()
    };
  }

  public renderTiendas(d: any, sourceLabel: string): void {
    document.getElementById('tabTiendas')!.textContent = 'SEM ' + d.week;
    document.getElementById('srcTiendas')!.textContent = sourceLabel;
    document.getElementById('noteTiendas')!.textContent = `${d.total} registros de llegada/salida en ruta durante la semana ${d.week}. El cumplimiento mide si el vehículo permaneció en el punto dentro del tiempo estándar asignado.`;
    document.getElementById('rutasMesLabel')!.textContent = '(semana ' + d.week + ')';

    const cumple = d.dist['CUMPLE'] || 0, parcial = d.dist['CUMPLE PARCIAL'] || 0, noCumple = d.dist['NO CUMPLE'] || 0;
    
    document.getElementById('kpiTiendasTotal')!.textContent = d.total;
    document.getElementById('kpiTiendasPeriodo')!.textContent = 'semana ' + d.week;
    document.getElementById('kpiTiendasCumple')!.textContent = this.fmtPct(cumple / d.total * 100);
    document.getElementById('kpiTiendasCumpleSub')!.textContent = cumple + ' registros';
    document.getElementById('kpiTiendasParcial')!.textContent = this.fmtPct(parcial / d.total * 100);
    document.getElementById('kpiTiendasParcialSub')!.textContent = parcial + ' registros';
    document.getElementById('kpiTiendasNoCumple')!.textContent = this.fmtPct(noCumple / d.total * 100);
    document.getElementById('kpiTiendasNoCumpleSub')!.textContent = noCumple + ' registros';

    document.getElementById('tableCedi')!.innerHTML = d.cedis.map((cd: string) => {
      const pct = d.cediMap[cd] / d.total * 100;
      return `<tr><td class="name">${cd}</td><td>${d.cediMap[cd]}</td><td>${this.fmtPct(pct)}</td></tr>`;
    }).join('');

    document.getElementById('tableRutas')!.innerHTML = d.topRegiones.map((rg: string) => {
      const info = d.regionMap[rg];
      const pct = info.total ? info.cumple / info.total * 100 : null;
      return `<tr><td class="name">${rg}</td><td>${info.total}</td><td><span class="pill ${this.pillClass(pct)}">${this.fmtPct(pct)}</span></td></tr>`;
    }).join('');

    const labels = Object.keys(d.dist);
    this.makeDoughnut('chartTiendas', labels, labels.map(l => d.dist[l]), labels.map((l, i) => {
      if (l === 'CUMPLE') return '#2F6F4E';
      if (l === 'CUMPLE PARCIAL') return '#BE8A2E';
      if (l === 'NO CUMPLE') return '#B5453A';
      return PALETTE[i % PALETTE.length];
    }));

    document.getElementById('tiendasLatestLabel')!.textContent = '(semana ' + d.week + ')';
    document.getElementById('tiendasLatestLabel2')!.textContent = '(semana ' + d.week + ')';

    const trendLabels = d.trend.weeks.map((w: number) => 'S' + w);
    this.makeLine('chartTiendasTrend', trendLabels, [
      { label: '% Cumple', data: d.trend.cumple, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3 },
      { label: '% Cumple parcial', data: d.trend.parcial, borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% No cumple', data: d.trend.noCumple, borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);

    if (d.monthly) {
      this.renderMonthly('chartTiendasMes', 'tableTiendasMes', d.monthly.meses, d.monthly.total, d.monthly.cumple, 'Registros', '% Cumple');
    }
  }

  // ==========================================
  // RENDER GENÉRICO DE RESUMEN MENSUAL
  // ==========================================

  public renderMonthly(chartId: string, tableId: string, meses: string[], totales: number[], pcts: (number | null)[], labelTotal: string, labelPct: string): void {
    this.makeLine(chartId, meses.map(m => m.charAt(0) + m.slice(1).toLowerCase()), [
      { label: labelTotal, data: totales, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3, yAxisID: 'y', type: 'bar' },
      { label: labelPct, data: pcts, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3, borderDash: [4, 3], yAxisID: 'y1' }
    ], true);

    document.getElementById(tableId)!.innerHTML = meses.map((m, i) => {
      const isCurrent = m === CURRENT_MONTH_NAME;
      const estado = isCurrent ? '<span class="pill mid">En curso</span>' : '<span class="pill good">Cerrado</span>';
      const label = m.charAt(0) + m.slice(1).toLowerCase();
      return `<tr><td class="name">${label}</td><td>${totales[i]}</td><td>${this.fmtPct(pcts[i])}</td><td>${estado}</td></tr>`;
    }).join('');
  }

  // ==========================================
  // PARSER 3: ON TIME AGRICULTORES
  // ==========================================

  public parseAgri(wb: any): any {
    const sheetName = this.findSheet(wb, ['RESUMEN']);
    if (!sheetName) throw new Error('No encontré la hoja "RESUMEN" en este archivo.');
    const rows = this.sheetRows(wb, sheetName);

    let headerRow = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i] && String(rows[i][1] || '').trim().toUpperCase() === 'AGRICULTOR') {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) throw new Error('No encontré la tabla "Agricultor" en la hoja RESUMEN.');

    const weekRow = rows[headerRow + 1];
    const weekCols: number[] = [];
    for (let c = 3; c < weekRow.length; c++) {
      if (weekRow[c] !== null && weekRow[c] !== undefined && weekRow[c] !== '') weekCols.push(c);
    }
    if (!weekCols.length) throw new Error('No encontré columnas de semana (S..) en RESUMEN.');

    const nWeeks = (weekCols.length / 3) | 0;
    const llegadaCols = weekCols.slice(0, nWeeks);
    const lastWeekCol = llegadaCols[llegadaCols.length - 1];
    const tiempoCols = weekCols.slice(nWeeks, 2 * nWeeks);
    const plantaCols = weekCols.slice(2 * nWeeks, 3 * nWeeks);
    const lastTiempoCol = tiempoCols[tiempoCols.length - 1];
    const lastPlantaCol = plantaCols[plantaCols.length - 1];
    const weekLabel = weekRow[lastWeekCol];

    function toPct(v: any): number | null {
      if (v === null || v === undefined || v === '-') return null;
      if (typeof v === 'string' && v.trim().endsWith('%')) return parseFloat(v);
      if (typeof v === 'number') return v <= 1 ? v * 100 : v;
      return null;
    }

    const agricultores: any[] = [];
    let r = headerRow + 2;
    const blockStarts: number[] = [];

    while (r < rows.length) {
      const name = rows[r] && rows[r][1];
      if (!name) break;
      blockStarts.push(r);
      const totalViajes = rows[r][lastWeekCol];
      const pctLlegada = toPct(rows[r + 3] ? rows[r + 3][lastWeekCol] : null);
      const pctTiempo = toPct(rows[r + 3] ? rows[r + 3][lastTiempoCol] : null);
      const pctPlanta = toPct(rows[r + 3] ? rows[r + 3][lastPlantaCol] : null);
      
      if (typeof totalViajes === 'number') {
        agricultores.push({ name: String(name).trim(), viajes: totalViajes, pctLlegada, pctTiempo, pctPlanta });
      }
      r += 4;
    }

    const totalViajes = agricultores.reduce((s, a) => s + a.viajes, 0);

    function weighted(key: string): number | null {
      let num = 0, den = 0;
      agricultores.forEach(a => {
        if (a[key] !== null) { num += a[key] * a.viajes; den += a.viajes; }
      });
      return den ? num / den : null;
    }

    const weekSeries = (cols: number[]): (number | null)[] => {
      return cols.map(col => {
        let num = 0, den = 0;
        blockStarts.forEach(r0 => {
          const total = rows[r0][col];
          if (typeof total !== 'number') return;
          let cumple = rows[r0 + 1] ? rows[r0 + 1][col] : null;
          if (typeof cumple !== 'number') cumple = 0;
          num += cumple; den += total;
        });
        return den ? (num / den) * 100 : null;
      });
    };

    const weekLabels = llegadaCols.map(c => String(weekRow[c]).replace(/[^0-9]/g, ''));
    const trend = { weeks: weekLabels, llegada: weekSeries(llegadaCols), tiempo: weekSeries(tiempoCols), planta: weekSeries(plantaCols) };

    return { weekLabel: weekLabel || '', agricultores, totalViajes, avgLlegada: weighted('pctLlegada'), avgTiempo: weighted('pctTiempo'), avgPlanta: weighted('pctPlanta'), trend };
  }

  public renderAgri(d: any, sourceLabel: string): void {
    const weekNum = String(d.weekLabel).replace(/[^0-9]/g, '') || d.weekLabel;
    const wLabel = 'Semana ' + weekNum;
    
    document.getElementById('tabAgri')!.textContent = 'SEM ' + weekNum;
    document.getElementById('srcAgri')!.textContent = sourceLabel;
    document.getElementById('kpiAgriSemanaLabel')!.textContent = 'semana ' + weekNum;
    document.getElementById('chartAgriLabel')!.textContent = '(S' + weekNum + ')';
    document.getElementById('tableAgriLabel')!.textContent = wLabel.toLowerCase();
    
    document.getElementById('kpiAgriTotal')!.textContent = d.totalViajes;
    document.getElementById('kpiAgriTotalSub')!.textContent = d.agricultores.length + ' agricultores';
    document.getElementById('kpiAgriLlegada')!.textContent = this.fmtPct(d.avgLlegada);
    document.getElementById('kpiAgriTiempo')!.textContent = this.fmtPct(d.avgTiempo);
    document.getElementById('kpiAgriPlanta')!.textContent = this.fmtPct(d.avgPlanta);

    document.getElementById('tableAgri')!.innerHTML = d.agricultores.map((a: any) => `
      <tr>
        <td class="name">${a.name}</td><td>${a.viajes}</td>
        <td><span class="pill ${this.pillClass(a.pctLlegada)}">${this.fmtPct(a.pctLlegada)}</span></td>
        <td><span class="pill ${this.pillClass(a.pctTiempo)}">${this.fmtPct(a.pctTiempo)}</span></td>
        <td><span class="pill ${this.pillClass(a.pctPlanta)}">${this.fmtPct(a.pctPlanta)}</span></td>
      </tr>
    `).join('');

    this.makeHBar('chartAgri', d.agricultores.map((a: any) => a.name), d.agricultores.map((a: any) => a.pctLlegada || 0), d.agricultores.map((a: any) => this.pillClass(a.pctLlegada) === 'good' ? '#2F6F4E' : this.pillClass(a.pctLlegada) === 'mid' ? '#BE8A2E' : '#B5453A'));

    const trendLabels = d.trend.weeks.map((w: string) => 'S' + w);
    this.makeLine('chartAgriTrend', trendLabels, [
      { label: '% Cumple llegada', data: d.trend.llegada, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3 },
      { label: '% Cumple tiempo en finca', data: d.trend.tiempo, borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% Cumple llegada a planta', data: d.trend.planta, borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);
  }

  // ==========================================
  // MANEJADORES DE CARGA (UPLOAD HANDLING)
  // ==========================================

  public setStatus(key: string, msg: string, cls?: string): void {
    const el = document.getElementById('status' + key.charAt(0).toUpperCase() + key.slice(1))!;
    el.textContent = msg;
    el.className = 'upload-status' + (cls ? (' ' + cls) : '');
  }

  public handleUpload(evt: any, kind: string): void {
    const file = evt.target.files[0];
    if (!file) return;

    const keyName = kind === 'makand' ? 'Makand' : kind === 'tiendas' ? 'Tiendas' : 'Agri';
    this.setStatus(keyName.toLowerCase(), 'Leyendo ' + file.name + '…');

    const reader = new FileReader();
    reader.onload = (e: any) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
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
        this.setStatus(keyName.toLowerCase(), '✓ actualizado ' + new Date().toLocaleString('es-CO'), 'ok');
      } catch (err: any) {
        console.error(err);
        this.setStatus(keyName.toLowerCase(), 'Error: ' + err.message, 'err');
      }
    };
    reader.onerror = () => {
      this.setStatus(keyName.toLowerCase(), 'No se pudo leer el archivo.', 'err');
    };
    reader.readAsArrayBuffer(file);
  }

  // ==========================================
  // CARGA INICIAL DE DATOS HISTÓRICOS (ESTÁTICOS)
  // ==========================================

  private inicializarCuadrosControl(): void {
    this.makeDoughnut('chartTransp', ['MAKAND', 'ARSITRANS', 'POLAR'], [23, 17, 11], ['#2F6F4E', '#BE8A2E', '#B5453A']);
    this.makeDoughnut('chartTiendas', ['No cumple', 'Cumple', 'Cumple parcial'], [69, 54, 16], ['#B5453A', '#2F6F4E', '#BE8A2E']);
    this.makeHBar('chartAgri', ['Diego', 'Ferrucas', 'Gabriel', 'Georgeth', 'José T.', 'Lechugas'], [16.7, 65.9, 0, 0, 0, 100], ['#B5453A', '#2F6F4E', '#B5453A', '#B5453A', '#B5453A', '#2F6F4E']);

    const MAKAND_WEEKS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28];
    const MAKAND_TOTAL = [24, 53, 49, 49, 51, 45, 56, 53, 51, 57, 57, 58, 58, 49, 62, 67, 55, 60, 63, 61, 60, 64, 57, 58, 58, 65, 56, 51];
    const MAKAND_LLEGADA = [100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 100.0, 93.3, 73.0, 75.4, 85.0, 79.7, 82.5, 81.0, 79.3, 76.9, 91.1, 74.5];
    this.makeLine('chartMakandTrend', MAKAND_WEEKS.map(w => 'S' + w), [
      { label: 'Viajes totales', data: MAKAND_TOTAL, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3, yAxisID: 'y' },
      { label: '% Cumplimiento llegada', data: MAKAND_LLEGADA, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3, borderDash: [4, 3], yAxisID: 'y1' }
    ], true);

    const TIENDAS_WEEKS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28];
    const TIENDAS_CUMPLE = [32.5, 38.9, 37.1, 32.0, 39.6, 40.1, 42.6, 40.0, 43.3, 35.1, 39.5, 36.8, 38.6, 29.7, 40.3, 40.3, 40.1, 27.0, 33.5, 39.6, 32.2, 31.5, 34.3, 38.2, 31.0, 31.0, 34.3, 38.8];
    const TIENDAS_PARCIAL = [25.0, 13.0, 16.9, 13.6, 11.9, 11.2, 13.2, 15.2, 7.8, 16.0, 12.0, 10.4, 11.1, 12.3, 13.1, 10.0, 7.8, 14.3, 11.3, 10.7, 15.6, 13.2, 14.3, 9.8, 10.3, 12.5, 14.8, 11.5];
    const TIENDAS_NOCUMPLE = [42.5, 48.1, 46.0, 54.4, 48.5, 48.7, 44.1, 44.8, 48.9, 48.9, 48.5, 52.7, 50.3, 58.0, 46.6, 49.8, 50.0, 56.6, 55.2, 49.7, 52.2, 55.3, 51.4, 52.0, 58.7, 56.5, 51.0, 49.6];
    this.makeLine('chartTiendasTrend', TIENDAS_WEEKS.map(w => 'S' + w), [
      { label: '% Cumple', data: TIENDAS_CUMPLE, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3 },
      { label: '% Cumple parcial', data: TIENDAS_PARCIAL, borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% No cumple', data: TIENDAS_NOCUMPLE, borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);

    this.makeLine('chartAgriTrend', ['S22', 'S23'], [
      { label: '% Cumple llegada', data: [71.6, 54.3], borderColor: '#223A57', backgroundColor: '#223A57', tension: .3 },
      { label: '% Cumple tiempo en finca', data: [43.2, 23.4], borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% Cumple llegada a planta', data: [6.9, 0.0], borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);

    const MAKAND_MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO'];
    const MAKAND_MES_TOTAL = [226, 205, 253, 255, 266, 258, 84];
    const MAKAND_MES_LLEGADA = [100.0, 100.0, 100.0, 100.0, 77.8, 80.6, 82.1];
    this.renderMonthly('chartMakandMes', 'tableMakandMes', MAKAND_MESES, MAKAND_MES_TOTAL, MAKAND_MES_LLEGADA, 'Viajes', '% Cumple llegada');

    const TIENDAS_MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO'];
    const TIENDAS_MES_TOTAL = [877, 813, 854, 803, 864, 857, 284];
    const TIENDAS_MES_CUMPLE = [36.7, 41.9, 37.2, 36.9, 33.6, 33.0, 39.4];
    this.renderMonthly('chartTiendasMes', 'tableTiendasMes', TIENDAS_MESES, TIENDAS_MES_TOTAL, TIENDAS_MES_CUMPLE, 'Registros', '% Cumple');
  }

  /**
   * Método de navegación por rutas utilizando el Router inyectado de Angular
   * @param ruta Destino dentro de la aplicación (ej: 'registro-operaciones')
   */
  public navegarA(ruta: string): void {
    if (ruta) {
      this.router.navigate([`/${ruta}`]);
    }
  }
}

// ==========================================
// EL EXPORT PARA TU ENRUTADO
// ==========================================
export const dashSem = 'dash-semanal';