import { Component, ElementRef, AfterViewInit, ViewChild, signal, computed } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

/*
 * Dependencias necesarias (instalar antes de usar este componente):
 *   npm install xlsx chart.js
 *
 * Uso:
 *   1. Copia los 3 archivos (reporte-semanal.component.ts / .html / .css) a tu proyecto Angular,
 *      por ejemplo en src/app/reporte-semanal/
 *   2. Es un componente STANDALONE: solo impórtalo donde lo necesites:
 *        import { ReporteSemanalComponent } from './reporte-semanal/reporte-semanal.component';
 *      y agrégalo a los "imports" de tu componente/módulo, o úsalo directo en un componente standalone padre.
 *   3. En el HTML del padre: <app-reporte-semanal></app-reporte-semanal>
 */

// ---------- Tipos ----------
interface TransportadoraStat {
  total: number;
  llegadaOk: number;
  cargueOk: number;
  cajas: number;
  almacenes: Record<string, number>;
}

interface TrendMakand {
  weeks: number[];
  total: number[];
  llegada: (number | null)[];
}

interface MonthlyData {
  meses: string[];
  total: number[];
  pct: (number | null)[];
}

interface MakandData {
  week: number | null;
  total: number;
  transportadoras: string[];
  transpMap: Record<string, TransportadoraStat>;
  almacenes: string[];
  overallLlegada: number;
  totalCajas: number;
  lider: string;
  liderPct: number;
  trend: TrendMakand;
  monthly: MonthlyData | null;
}

interface RegionStat { total: number; cumple: number; }

interface TrendTiendas {
  weeks: number[];
  cumple: (number | null)[];
  parcial: (number | null)[];
  noCumple: (number | null)[];
  total: number[];
}

interface TiendasData {
  week: number | null;
  total: number;
  dist: Record<string, number>;
  cediMap: Record<string, number>;
  cedis: string[];
  regionMap: Record<string, RegionStat>;
  topRegiones: string[];
  trend: TrendTiendas;
  monthly: MonthlyData | null;
}

interface AgricultorRow {
  name: string;
  viajes: number;
  pctLlegada: number | null;
  pctTiempo: number | null;
  pctPlanta: number | null;
}

interface TrendAgri {
  weeks: string[];
  llegada: (number | null)[];
  tiempo: (number | null)[];
  planta: (number | null)[];
}

interface MetricStat { avg: number | null; count: number; }

interface AgriData {
  weekLabel: string;
  agricultores: AgricultorRow[];
  totalViajes: number;
  llegada: MetricStat;
  tiempo: MetricStat;
  planta: MetricStat;
  trend: TrendAgri;
  monthly: MonthlyData | null;
}

@Component({
  selector: 'app-reporte-semanal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './reporte_semanal_operaciones.html',
  styleUrl: './reporte-semanal.component.css'
})
export class ReporteSemanalComponent implements AfterViewInit {

  // ---------- refs a los <canvas> ----------
  @ViewChild('chartTransp') chartTranspRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartMakandTrend') chartMakandTrendRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartMakandMes') chartMakandMesRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartTiendas') chartTiendasRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartTiendasTrend') chartTiendasTrendRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartTiendasMes') chartTiendasMesRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAgri') chartAgriRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAgriTrend') chartAgriTrendRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAgriMes') chartAgriMesRef!: ElementRef<HTMLCanvasElement>;

  private charts: Record<string, Chart> = {};

  readonly PALETTE = ['#2F6F4E', '#BE8A2E', '#B5453A', '#3E6C99', '#7A5C99', '#5C8A6E'];
  readonly MESES_ORDER = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  readonly CURRENT_MONTH_NAME = this.MESES_ORDER[new Date().getMonth()];

  // ---------- estado reactivo (signals) ----------
  makand = signal<MakandData>(this.datosIniciales_makand());
  tiendas = signal<TiendasData>(this.datosIniciales_tiendas());
  agri = signal<AgriData>(this.datosIniciales_agri());

  fuenteMakand = signal<string>('VIAJEROS_Y_TERCEROS_MAKAND.xlsx');
  fuenteTiendas = signal<string>('ON_TIME_EN_TIENDAS.xlsx');
  fuenteAgri = signal<string>('ON_TIME_AGRICULTORES_2026_ACT.xlsx');

  statusMakand = signal<{msg: string, cls: string}>({msg: '', cls: ''});
  statusTiendas = signal<{msg: string, cls: string}>({msg: '', cls: ''});
  statusAgri = signal<{msg: string, cls: string}>({msg: '', cls: ''});

  transportadorasMakand = computed(() => this.makand().transportadoras);

  ngAfterViewInit(): void {
    // pequeño delay para asegurar que los <canvas> ya midieron su tamaño en el DOM
    setTimeout(() => {
      this.renderMakandCharts(this.makand());
      this.renderTiendasCharts(this.tiendas());
      this.renderAgriCharts(this.agri());
    });
  }

  // =========================================================
  //  HELPERS GENERALES
  // =========================================================
  private fmtPct(pct: number | null | undefined): string {
    return (pct === null || pct === undefined || isNaN(pct)) ? '—' : pct.toFixed(1) + '%';
  }
  pillClass(pct: number | null | undefined): string {
    if (pct === null || pct === undefined || isNaN(pct)) return 'mid';
    if (pct >= 70) return 'good';
    if (pct >= 40) return 'mid';
    return 'bad';
  }
  fmt(pct: number | null | undefined): string { return this.fmtPct(pct); }

  private findSheet(wb: XLSX.WorkBook, patterns: string[]): string | null {
    for (const p of patterns) {
      const hit = wb.SheetNames.find(n => n.toUpperCase().replace(/\s+/g, ' ').trim().includes(p));
      if (hit) return hit;
    }
    return null;
  }
  private sheetRows(wb: XLSX.WorkBook, name: string): any[][] {
    const ws = wb.Sheets[name];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as any[][];
  }
  private colIndex(header: any[], names: string | string[]): number {
    const candidates = Array.isArray(names) ? names : [names];
    for (const name of candidates) {
      const idx = header.findIndex(h => h && String(h).trim().toUpperCase() === name.toUpperCase());
      if (idx >= 0) return idx;
    }
    return -1;
  }
  private pickLatestWeek(counts: Record<number, number>): number | null {
    const weeks = Object.keys(counts).map(Number).filter(w => !isNaN(w)).sort((a, b) => a - b);
    if (!weeks.length) return null;
    return weeks[weeks.length - 1];
  }

  // =========================================================
  //  SUBIDA DE ARCHIVOS
  // =========================================================
  onFileSelected(evt: Event, kind: 'makand' | 'tiendas' | 'agri'): void {
    const input = evt.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    const statusSignal = kind === 'makand' ? this.statusMakand : kind === 'tiendas' ? this.statusTiendas : this.statusAgri;
    statusSignal.set({ msg: 'Leyendo ' + file.name + '…', cls: '' });

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: false });

        if (kind === 'makand') {
          const d = this.parseMakand(wb);
          this.makand.set(d);
          this.fuenteMakand.set(file.name);
          this.renderMakandCharts(d);
        } else if (kind === 'tiendas') {
          const d = this.parseTiendas(wb);
          this.tiendas.set(d);
          this.fuenteTiendas.set(file.name);
          this.renderTiendasCharts(d);
        } else {
          const d = this.parseAgri(wb);
          this.agri.set(d);
          this.fuenteAgri.set(file.name);
          this.renderAgriCharts(d);
        }
        statusSignal.set({ msg: '✓ actualizado ' + new Date().toLocaleString('es-CO'), cls: 'ok' });
      } catch (err: any) {
        console.error(err);
        statusSignal.set({ msg: 'Error: ' + err.message, cls: 'err' });
      }
      input.value = '';
    };
    reader.onerror = () => statusSignal.set({ msg: 'No se pudo leer el archivo.', cls: 'err' });
    reader.readAsArrayBuffer(file);
  }

  // =========================================================
  //  PARSER 1: VIAJEROS Y TERCEROS MAKAND
  // =========================================================
  private parseMakand(wb: XLSX.WorkBook): MakandData {
    const sheetName = this.findSheet(wb, ['DT VIAJEROS']);
    if (!sheetName) throw new Error('No encontré la hoja "DT VIAJEROS." en este archivo.');
    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];
    const iSem = this.colIndex(header, ['SEMANAS', 'SEMANA']);
    const iTransp = this.colIndex(header, 'TRANSPORTE');
    const iAlmacen = this.colIndex(header, 'ALMACEN');
    const iLlegada = this.colIndex(header, 'CUMPLIMIENTO LLEGADA VEHICULO');
    const iCargue = this.colIndex(header, 'CUMPLIMIENTO CARGUE VEHICULO');
    const iCajas = this.colIndex(header, 'TOTAL CAJAS');
    const iMes = this.colIndex(header, 'MES');
    if (iSem < 0 || iTransp < 0) throw new Error('No encontré las columnas SEMANA/TRANSPORTE esperadas.');

    const data = rows.slice(1).filter(r => r && r[iSem] !== null && r[iTransp]);

    const counts: Record<number, number> = {};
    data.forEach(r => { const w = Number(r[iSem]); counts[w] = (counts[w] || 0) + 1; });
    const week = this.pickLatestWeek(counts);
    const rowsWeek = data.filter(r => Number(r[iSem]) === week);

    const transpMap: Record<string, TransportadoraStat> = {};
    const almacenSet = new Set<string>();
    rowsWeek.forEach(r => {
      const t = String(r[iTransp]).trim().toUpperCase();
      if (!transpMap[t]) transpMap[t] = { total: 0, llegadaOk: 0, cargueOk: 0, cajas: 0, almacenes: {} };
      transpMap[t].total++;
      if (iLlegada >= 0 && String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE') transpMap[t].llegadaOk++;
      if (iCargue >= 0 && String(r[iCargue]).trim().toUpperCase() === 'CUMPLE') transpMap[t].cargueOk++;
      if (iCajas >= 0) { const v = parseFloat(r[iCajas]); if (!isNaN(v)) transpMap[t].cajas += v; }
      const alm = iAlmacen >= 0 ? String(r[iAlmacen] || '').trim() : 'N/D';
      almacenSet.add(alm);
      transpMap[t].almacenes[alm] = (transpMap[t].almacenes[alm] || 0) + 1;
    });

    const total = rowsWeek.length;
    const transportadoras = Object.keys(transpMap).sort((a, b) => transpMap[b].total - transpMap[a].total);
    const overallLlegada = total ? (Object.values(transpMap).reduce((s, t) => s + t.llegadaOk, 0) / total * 100) : 0;
    const totalCajas = Object.values(transpMap).reduce((s, t) => s + t.cajas, 0);
    const lider = transportadoras[0];

    // ---- tendencia: TODAS las semanas ----
    const allWeeks = [...new Set(data.map(r => Number(r[iSem])))].sort((a, b) => a - b);
    const trendTotal: number[] = [];
    const trendLlegada: (number | null)[] = [];
    allWeeks.forEach(w => {
      const rw = data.filter(r => Number(r[iSem]) === w);
      trendTotal.push(rw.length);
      const ok = iLlegada >= 0 ? rw.filter(r => String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE').length : 0;
      trendLlegada.push(rw.length ? ok / rw.length * 100 : null);
    });

    // ---- resumen mensual ----
    let monthly: MonthlyData | null = null;
    if (iMes >= 0) {
      const mesesPresentes = [...new Set(data.map(r => String(r[iMes]).trim().toUpperCase()))]
        .filter(m => this.MESES_ORDER.includes(m))
        .sort((a, b) => this.MESES_ORDER.indexOf(a) - this.MESES_ORDER.indexOf(b));
      const totalPorMes: number[] = [];
      const llegadaPorMes: (number | null)[] = [];
      mesesPresentes.forEach(m => {
        const rm = data.filter(r => String(r[iMes]).trim().toUpperCase() === m);
        totalPorMes.push(rm.length);
        const ok = iLlegada >= 0 ? rm.filter(r => String(r[iLlegada]).trim().toUpperCase() === 'CUMPLE').length : 0;
        llegadaPorMes.push(rm.length ? ok / rm.length * 100 : null);
      });
      monthly = { meses: mesesPresentes, total: totalPorMes, pct: llegadaPorMes };
    }

    return {
      week, total, transportadoras, transpMap, almacenes: [...almacenSet],
      overallLlegada, totalCajas, lider,
      liderPct: total ? (transpMap[lider].total / total * 100) : 0,
      trend: { weeks: allWeeks, total: trendTotal, llegada: trendLlegada },
      monthly
    };
  }

  // =========================================================
  //  PARSER 2: ON TIME EN TIENDAS
  // =========================================================
  private parseTiendas(wb: XLSX.WorkBook): TiendasData {
    const sheetName = this.findSheet(wb, ['DETALLE REGIONES']);
    if (!sheetName) throw new Error('No encontré la hoja "Detalle regiones" en este archivo.');
    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];
    const iSem = this.colIndex(header, ['SEMANA', 'SEMANAS']);
    const iCump = this.colIndex(header, 'CUMPLIMIENTO');
    const iRegion = this.colIndex(header, 'REGIÓN');
    const iCedi = this.colIndex(header, 'CEDI');
    const iMes = this.colIndex(header, 'MES');
    if (iSem < 0 || iCump < 0) throw new Error('No encontré las columnas SEMANA/CUMPLIMIENTO esperadas.');

    const data = rows.slice(1).filter(r => r && r[iSem] !== null);
    const counts: Record<number, number> = {};
    data.forEach(r => { const w = Number(r[iSem]); counts[w] = (counts[w] || 0) + 1; });
    const week = this.pickLatestWeek(counts);
    const rowsWeek = data.filter(r => Number(r[iSem]) === week);

    const dist: Record<string, number> = {};
    const cediMap: Record<string, number> = {};
    const regionMap: Record<string, RegionStat> = {};
    rowsWeek.forEach(r => {
      const c = iCump >= 0 ? String(r[iCump] || '').trim().toUpperCase() : 'N/D';
      dist[c] = (dist[c] || 0) + 1;
      if (iCedi >= 0) { const cd = String(r[iCedi] || '').trim() || 'N/D'; cediMap[cd] = (cediMap[cd] || 0) + 1; }
      if (iRegion >= 0) {
        const rg = String(r[iRegion] || '').trim() || 'N/D';
        if (!regionMap[rg]) regionMap[rg] = { total: 0, cumple: 0 };
        regionMap[rg].total++;
        if (c === 'CUMPLE') regionMap[rg].cumple++;
      }
    });

    const total = rowsWeek.length;
    const topRegiones = Object.keys(regionMap).sort((a, b) => regionMap[b].total - regionMap[a].total).slice(0, 10);
    const cedis = Object.keys(cediMap).sort((a, b) => cediMap[b] - cediMap[a]);

    // ---- tendencia: TODAS las semanas ----
    const allWeeks = [...new Set(data.map(r => Number(r[iSem])))].sort((a, b) => a - b);
    const trendCumple: (number | null)[] = [], trendParcial: (number | null)[] = [],
          trendNoCumple: (number | null)[] = [], trendTotal: number[] = [];
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

    // ---- resumen mensual ----
    let monthly: MonthlyData | null = null;
    if (iMes >= 0) {
      const mesesPresentes = [...new Set(data.map(r => String(r[iMes]).trim().toUpperCase()))]
        .filter(m => this.MESES_ORDER.includes(m))
        .sort((a, b) => this.MESES_ORDER.indexOf(a) - this.MESES_ORDER.indexOf(b));
      const totalPorMes: number[] = [];
      const cumplePorMes: (number | null)[] = [];
      mesesPresentes.forEach(m => {
        const rm = data.filter(r => String(r[iMes]).trim().toUpperCase() === m);
        totalPorMes.push(rm.length);
        const ok = rm.filter(r => String(r[iCump] || '').trim().toUpperCase() === 'CUMPLE').length;
        cumplePorMes.push(rm.length ? ok / rm.length * 100 : null);
      });
      monthly = { meses: mesesPresentes, total: totalPorMes, pct: cumplePorMes };
    }

    return {
      week, total, dist, cediMap, cedis, regionMap, topRegiones,
      trend: { weeks: allWeeks, cumple: trendCumple, parcial: trendParcial, noCumple: trendNoCumple, total: trendTotal },
      monthly
    };
  }

  // =========================================================
  //  PARSER 3: ON TIME AGRICULTORES
  // =========================================================
  private parseAgri(wb: XLSX.WorkBook): AgriData {
    const sheetName = this.findSheet(wb, ['TIEMPO AGRICULTORES']);
    if (!sheetName) throw new Error('No encontré la hoja "TIEMPO AGRICULTORES" en este archivo.');
    const rows = this.sheetRows(wb, sheetName);
    const header = rows[0];

    const iSem = this.colIndex(header, ['SEMANA', 'SEMANAS']);
    const iAgricultor = this.colIndex(header, 'AGRICULTOR');
    const iLlegada = this.colIndex(header, 'CUMPLIMIENTO LLEGADA AGRICULTOR');
    const iTiempo = this.colIndex(header, 'CUMPLIMIENTO TIEMPO EN AGRICULTOR');
    const iPlanta = this.colIndex(header, 'CUMPLIMIENTO LLEGADA A PLANTA');
    const iMes = this.colIndex(header, 'MES');
    if (iSem < 0 || iAgricultor < 0) throw new Error('No encontré las columnas SEMANA/AGRICULTOR esperadas en "TIEMPO AGRICULTORES".');

    const EXCLUIR = ['LECHUGAS DEL DIA', 'LECHUGAS DEL DÍA'];
    const normName = (s: string) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
    const displayName = (raw: string): string => {
      const n = normName(raw);
      if (n.startsWith('FERRUCAS')) return 'FERRUCAS';
      return String(raw).trim();
    };

    const data = rows.slice(1).filter(r => r && r[iSem] !== null && r[iAgricultor] && !EXCLUIR.includes(normName(r[iAgricultor])));

    const counts: Record<number, number> = {};
    data.forEach(r => { const w = Number(r[iSem]); counts[w] = (counts[w] || 0) + 1; });
    const week = this.pickLatestWeek(counts);
    const rowsWeek = data.filter(r => Number(r[iSem]) === week);

    // ---- por agricultor (agrupando FERRUCAS N -> FERRUCAS) ----
    interface AgriAgg { viajes: number; llegadaOk: number; llegadaTot: number; tiempoOk: number; tiempoTot: number; plantaOk: number; plantaTot: number; }
    const agriMap: Record<string, AgriAgg> = {};
    rowsWeek.forEach(r => {
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

    const agricultores: AgricultorRow[] = Object.keys(agriMap)
      .sort((a, b) => agriMap[b].viajes - agriMap[a].viajes)
      .map(name => {
        const m = agriMap[name];
        return {
          name, viajes: m.viajes,
          pctLlegada: m.llegadaTot ? m.llegadaOk / m.llegadaTot * 100 : null,
          pctTiempo: m.tiempoTot ? m.tiempoOk / m.tiempoTot * 100 : null,
          pctPlanta: m.plantaTot ? m.plantaOk / m.plantaTot * 100 : null
        };
      });

    const totalViajes = rowsWeek.length;

    const statFrom = (field: number): MetricStat => {
      let ok = 0, tot = 0;
      rowsWeek.forEach(r => {
        const v = String(r[field] || '').trim().toUpperCase();
        if (v === 'CUMPLE' || v === 'NO CUMPLE') { tot++; if (v === 'CUMPLE') ok++; }
      });
      return { avg: tot ? ok / tot * 100 : null, count: ok };
    };

    // ---- tendencia: TODAS las semanas ----
    const allWeeks = [...new Set(data.map(r => Number(r[iSem])))].sort((a, b) => a - b);
    const trendSeries = (field: number): (number | null)[] => {
      return allWeeks.map(w => {
        const rw = data.filter(r => Number(r[iSem]) === w);
        let ok = 0, tot = 0;
        rw.forEach(r => {
          const v = String(r[field] || '').trim().toUpperCase();
          if (v === 'CUMPLE' || v === 'NO CUMPLE') { tot++; if (v === 'CUMPLE') ok++; }
        });
        return tot ? ok / tot * 100 : null;
      });
    };
    const trend: TrendAgri = {
      weeks: allWeeks.map(String),
      llegada: trendSeries(iLlegada),
      tiempo: trendSeries(iTiempo),
      planta: trendSeries(iPlanta)
    };

    // ---- resumen mensual ----
    let monthly: MonthlyData | null = null;
    if (iMes >= 0) {
      const mesesPresentes = [...new Set(data.map(r => String(r[iMes]).trim().toUpperCase()))]
        .filter(m => this.MESES_ORDER.includes(m))
        .sort((a, b) => this.MESES_ORDER.indexOf(a) - this.MESES_ORDER.indexOf(b));
      const totalPorMes: number[] = [];
      const llegadaPorMes: (number | null)[] = [];
      mesesPresentes.forEach(m => {
        const rm = data.filter(r => String(r[iMes]).trim().toUpperCase() === m);
        totalPorMes.push(rm.length);
        let ok = 0, tot = 0;
        rm.forEach(r => {
          const v = String(r[iLlegada] || '').trim().toUpperCase();
          if (v === 'CUMPLE' || v === 'NO CUMPLE') { tot++; if (v === 'CUMPLE') ok++; }
        });
        llegadaPorMes.push(tot ? ok / tot * 100 : null);
      });
      monthly = { meses: mesesPresentes, total: totalPorMes, pct: llegadaPorMes };
    }

    return {
      weekLabel: String(week ?? ''), agricultores, totalViajes,
      llegada: statFrom(iLlegada), tiempo: statFrom(iTiempo), planta: statFrom(iPlanta),
      trend, monthly
    };
  }

  // =========================================================
  //  RENDER DE GRÁFICAS (Chart.js)
  // =========================================================
  private makeDoughnut(canvas: HTMLCanvasElement, id: string, labels: string[], data: number[], colors: string[]): void {
    this.charts[id]?.destroy();
    this.charts[id] = new Chart(canvas, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#FFFEFB', borderWidth: 3 }] },
      options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, padding: 16 } } }, cutout: '62%' }
    });
  }

  private makeHBar(canvas: HTMLCanvasElement, id: string, labels: string[], data: number[], colors: string[]): void {
    this.charts[id]?.destroy();
    this.charts[id] = new Chart(canvas, {
      type: 'bar',
      data: { labels, datasets: [{ data, backgroundColor: colors }] },
      options: {
        responsive: true, maintainAspectRatio: false, indexAxis: 'y',
        scales: { x: { max: 100, grid: { color: '#E4DAC3' } }, y: { grid: { display: false } } },
        plugins: { legend: { display: false } }
      }
    });
  }

  private makeLine(canvas: HTMLCanvasElement, id: string, labels: string[], datasets: any[], dualAxis: boolean): void {
    this.charts[id]?.destroy();
    const scales: any = dualAxis ? {
      x: { grid: { display: false } },
      y: { position: 'left', grid: { color: '#E4DAC3' }, title: { display: true, text: 'Viajes / registros', font: { size: 10 } } },
      y1: { position: 'right', min: 0, max: 100, grid: { display: false }, title: { display: true, text: '% cumplimiento', font: { size: 10 } } }
    } : {
      x: { grid: { display: false } },
      y: { min: 0, max: 100, grid: { color: '#E4DAC3' } }
    };
    this.charts[id] = new Chart(canvas, {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales,
        plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, padding: 14 } } }
      }
    });
  }

  private renderMonthly(canvas: HTMLCanvasElement, id: string, m: MonthlyData, labelTotal: string, labelPct: string): void {
    const labels = m.meses.map(mes => mes.charAt(0) + mes.slice(1).toLowerCase());
    this.makeLine(canvas, id, labels, [
      { label: labelTotal, data: m.total, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3, yAxisID: 'y', type: 'bar' },
      { label: labelPct, data: m.pct, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3, borderDash: [4, 3], yAxisID: 'y1' }
    ], true);
  }

  esMesActual(mes: string): boolean { return mes === this.CURRENT_MONTH_NAME; }
  nombreMes(mes: string): string { return mes.charAt(0) + mes.slice(1).toLowerCase(); }

  private renderMakandCharts(d: MakandData): void {
    if (!this.chartTranspRef) return;
    this.makeDoughnut(this.chartTranspRef.nativeElement, 'chartTransp', d.transportadoras,
      d.transportadoras.map(t => d.transpMap[t].total),
      d.transportadoras.map((_, i) => this.PALETTE[i % this.PALETTE.length]));

    this.makeLine(this.chartMakandTrendRef.nativeElement, 'chartMakandTrend', d.trend.weeks.map(w => 'S' + w), [
      { label: 'Viajes totales', data: d.trend.total, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3, yAxisID: 'y' },
      { label: '% Cumplimiento llegada', data: d.trend.llegada, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3, borderDash: [4, 3], yAxisID: 'y1' }
    ], true);

    if (d.monthly) this.renderMonthly(this.chartMakandMesRef.nativeElement, 'chartMakandMes', d.monthly, 'Viajes', '% Cumple llegada');
  }

  private renderTiendasCharts(d: TiendasData): void {
    if (!this.chartTiendasRef) return;
    const labels = Object.keys(d.dist);
    this.makeDoughnut(this.chartTiendasRef.nativeElement, 'chartTiendas', labels, labels.map(l => d.dist[l]),
      labels.map((l, i) => l === 'CUMPLE' ? '#2F6F4E' : l === 'CUMPLE PARCIAL' ? '#BE8A2E' : l === 'NO CUMPLE' ? '#B5453A' : this.PALETTE[i % this.PALETTE.length]));

    this.makeLine(this.chartTiendasTrendRef.nativeElement, 'chartTiendasTrend', d.trend.weeks.map(w => 'S' + w), [
      { label: '% Cumple', data: d.trend.cumple, borderColor: '#2F6F4E', backgroundColor: '#2F6F4E', tension: .3 },
      { label: '% Cumple parcial', data: d.trend.parcial, borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% No cumple', data: d.trend.noCumple, borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);

    if (d.monthly) this.renderMonthly(this.chartTiendasMesRef.nativeElement, 'chartTiendasMes', d.monthly, 'Registros', '% Cumple');
  }

  private renderAgriCharts(d: AgriData): void {
    if (!this.chartAgriRef) return;
    this.makeHBar(this.chartAgriRef.nativeElement, 'chartAgri', d.agricultores.map(a => a.name), d.agricultores.map(a => a.pctLlegada || 0),
      d.agricultores.map(a => this.pillClass(a.pctLlegada) === 'good' ? '#2F6F4E' : this.pillClass(a.pctLlegada) === 'mid' ? '#BE8A2E' : '#B5453A'));

    this.makeLine(this.chartAgriTrendRef.nativeElement, 'chartAgriTrend', d.trend.weeks.map(w => 'S' + w), [
      { label: '% Cumple llegada', data: d.trend.llegada, borderColor: '#223A57', backgroundColor: '#223A57', tension: .3 },
      { label: '% Cumple tiempo en finca', data: d.trend.tiempo, borderColor: '#BE8A2E', backgroundColor: '#BE8A2E', tension: .3 },
      { label: '% Cumple llegada a planta', data: d.trend.planta, borderColor: '#B5453A', backgroundColor: '#B5453A', tension: .3 }
    ], false);

    if (d.monthly) this.renderMonthly(this.chartAgriMesRef.nativeElement, 'chartAgriMes', d.monthly, 'Registros', '% Cumple llegada');
  }

  // =========================================================
  //  DATOS INICIALES (placeholder antes de subir un archivo)
  //  Reemplaza estos valores por los de tu último reporte real,
  //  o simplemente sube un archivo apenas cargue la página.
  // =========================================================
  private datosIniciales_makand(): MakandData {
    return {
      week: 28, total: 51,
      transportadoras: ['MAKAND', 'ARSITRANS', 'POLAR'],
      transpMap: {
        MAKAND:    { total: 23, llegadaOk: 21, cargueOk: 13, cajas: 10500, almacenes: { D1: 16, ARA: 1, ÉXITO: 5, OLIMPICA: 1 } },
        ARSITRANS: { total: 17, llegadaOk: 11, cargueOk: 8,  cajas: 8552,  almacenes: { D1: 10, ARA: 4, ÉXITO: 3 } },
        POLAR:     { total: 11, llegadaOk: 6,  cargueOk: 9,  cajas: 2939,  almacenes: { D1: 11 } }
      },
      almacenes: ['D1', 'ARA', 'ÉXITO', 'OLIMPICA'],
      overallLlegada: 74.5, totalCajas: 21991, lider: 'MAKAND', liderPct: 45.1,
      trend: {
        weeks: [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28],
        total: [24,53,49,49,51,45,56,53,51,57,57,58,58,49,62,67,55,60,63,61,60,64,57,58,58,65,56,51],
        llegada: [100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,93.3,73.0,75.4,85.0,79.7,82.5,81.0,79.3,76.9,91.1,74.5]
      },
      monthly: {
        meses: ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO'],
        total: [226,205,253,255,266,258,84],
        pct: [100,100,100,100,77.8,80.6,82.1]
      }
    };
  }

  private datosIniciales_tiendas(): TiendasData {
    return {
      week: 28, total: 213,
      dist: { 'CUMPLE': 84, 'CUMPLE PARCIAL': 28, 'NO CUMPLE': 101 },
      cediMap: { 'Tiendas': 140, 'CEDI 2': 46, 'CEDI ARA': 14, 'CEDI 1': 13 },
      cedis: ['Tiendas', 'CEDI 2', 'CEDI ARA', 'CEDI 1'],
      regionMap: {
        'PLATAFORMA SIBERIA': { total: 16, cumple: 16 },
        'OLIMPICA': { total: 8, cumple: 1 },
        'ARA COTA': { total: 8, cumple: 1 },
        'PLATAFORMA CENCOSUD': { total: 7, cumple: 2 },
        'D1 TOCANCIPA': { total: 7, cumple: 4 },
        'JUMBO SANTA ANA': { total: 7, cumple: 0 },
        'D1 SIBATE -': { total: 6, cumple: 0 },
        'ARA GACHANCIPA': { total: 6, cumple: 0 },
        'CARULLA PEPE SIERRA': { total: 6, cumple: 2 },
        'CARULLA SANTA BARBARA': { total: 6, cumple: 0 }
      },
      topRegiones: ['PLATAFORMA SIBERIA', 'OLIMPICA', 'ARA COTA', 'PLATAFORMA CENCOSUD', 'D1 TOCANCIPA', 'JUMBO SANTA ANA', 'D1 SIBATE -', 'ARA GACHANCIPA', 'CARULLA PEPE SIERRA', 'CARULLA SANTA BARBARA'],
      trend: {
        weeks: [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29],
        cumple: [32.5,38.9,37.1,32.0,39.6,40.1,42.6,40.0,43.3,35.1,39.5,36.8,38.6,29.7,40.3,40.3,40.1,27.0,33.5,39.6,32.2,31.5,34.3,38.2,31.0,31.0,34.3,39.4,37.9],
        parcial: [25.0,13.0,16.9,13.6,11.9,11.2,13.2,15.2,7.8,16.0,12.0,10.4,11.1,12.3,13.1,10.0,7.8,14.3,11.3,10.7,15.6,13.2,14.3,9.8,10.3,12.5,14.8,13.1,10.7],
        noCumple: [42.5,48.1,46.0,54.4,48.5,48.7,44.1,44.8,48.9,48.9,48.5,52.7,50.3,58.0,46.6,49.8,50.0,56.6,55.2,49.7,52.2,55.3,51.4,52.0,58.7,56.5,51.0,47.4,51.5],
        total: [138,206,201,192,196,194,197,205,197,210,204,184,200,210,104,197,196,199,205,201,197,204,199,201,203,207,210,213,103]
      },
      monthly: {
        meses: ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO'],
        total: [877,813,854,803,864,857,461],
        pct: [36.7,41.9,37.2,36.9,33.6,33.0,39.3]
      }
    };
  }

  private datosIniciales_agri(): AgriData {
    return {
      weekLabel: '28',
      agricultores: [
        { name: 'Ferrucas', viajes: 37, pctLlegada: 89.2, pctTiempo: 13.5, pctPlanta: 45.9 },
        { name: 'Andres Cadena', viajes: 7, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 16.7 },
        { name: 'Juan Pablo', viajes: 7, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 83.3 },
        { name: 'Jesús', viajes: 7, pctLlegada: 100.0, pctTiempo: 100.0, pctPlanta: 100.0 },
        { name: 'Severo', viajes: 7, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 0.0 },
        { name: 'Wilson', viajes: 7, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 33.3 },
        { name: 'Jose Tibaquicha', viajes: 7, pctLlegada: 71.4, pctTiempo: 16.7, pctPlanta: 33.3 },
        { name: 'Gabriel', viajes: 6, pctLlegada: 16.7, pctTiempo: 60.0, pctPlanta: 0.0 },
        { name: 'Diego', viajes: 6, pctLlegada: 16.7, pctTiempo: 40.0, pctPlanta: 0.0 },
        { name: 'Georgeth', viajes: 2, pctLlegada: 50.0, pctTiempo: 0.0, pctPlanta: 0.0 },
        { name: 'Mario Acevedo', viajes: 7, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 83.3 },
        { name: 'Jorge Baracaldo', viajes: 1, pctLlegada: 100.0, pctTiempo: 0.0, pctPlanta: 0.0 }
      ],
      totalViajes: 101,
      llegada: { avg: 83.2, count: 84 },
      tiempo: { avg: 18.5, count: 17 },
      planta: { avg: 41.3, count: 38 },
      trend: {
        weeks: ['1','2','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29'],
        llegada: [78.6,71.4,74.5,51.3,56.2,40.3,41.3,31.1,33.3,38.1,27.6,44.1,52.5,79.5,73.0,86.0,80.5,83.0,93.1],
        tiempo: [21.4,28.6,25.5,30.8,34.4,29.0,45.7,27.0,18.9,21.2,24.1,23.0,29.5,25.9,32.2,32.2,23.0,21.6,null],
        planta: [37.5,53.8,54.3,35.8,35.6,32.2,36.5,27.6,32.7,37.8,34.2,37.2,32.8,32.1,30.4,34.7,41.4,39.2,null]
      },
      monthly: {
        meses: ['ENERO','MARZO','ABRIL','MAYO','JUNIO','JULIO'],
        total: [40,149,477,543,504,202],
        pct: [76.2,64.9,45.2,35.5,74.0,84.2]
      }
    };
  }
}