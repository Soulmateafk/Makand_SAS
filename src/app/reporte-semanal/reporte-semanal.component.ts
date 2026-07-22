import {
  Component,
  ElementRef,
  AfterViewInit,
  ViewChild,
  signal,
  computed,
  ChangeDetectorRef
} from '@angular/core';

import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

/*
 * Dependencias necesarias:
 *   npm install xlsx chart.js
 */

// =========================================================
// TIPOS
// =========================================================

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

interface RegionStat {
  total: number;
  cumple: number;
}

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

interface MetricStat {
  avg: number | null;
  count: number;
}

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

// =========================================================
// COMPONENTE
// =========================================================

@Component({
  selector: 'app-reporte-semanal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './reporte_semanal_operaciones.html',
  styleUrl: './reporte-semanal.component.css'
})
export class ReporteSemanalComponent implements AfterViewInit {

  constructor(
    private cdr: ChangeDetectorRef
  ) {}

  // =========================================================
  // REFERENCIAS A CANVAS
  // =========================================================

  @ViewChild('chartTransp')
  chartTranspRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartMakandTrend')
  chartMakandTrendRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartMakandMes')
  chartMakandMesRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartTiendas')
  chartTiendasRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartTiendasTrend')
  chartTiendasTrendRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartTiendasMes')
  chartTiendasMesRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartAgri')
  chartAgriRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartAgriTrend')
  chartAgriTrendRef!: ElementRef<HTMLCanvasElement>;

  @ViewChild('chartAgriMes')
  chartAgriMesRef!: ElementRef<HTMLCanvasElement>;

  private charts: Record<string, Chart> = {};

  // =========================================================
  // CONFIGURACIÓN
  // =========================================================

  readonly PALETTE = [
    '#2F6F4E',
    '#BE8A2E',
    '#B5453A',
    '#3E6C99',
    '#7A5C99',
    '#5C8A6E'
  ];

  readonly MESES_ORDER = [
    'ENERO',
    'FEBRERO',
    'MARZO',
    'ABRIL',
    'MAYO',
    'JUNIO',
    'JULIO',
    'AGOSTO',
    'SEPTIEMBRE',
    'OCTUBRE',
    'NOVIEMBRE',
    'DICIEMBRE'
  ];

  readonly CURRENT_MONTH_NAME =
    this.MESES_ORDER[new Date().getMonth()];

  // =========================================================
  // ESTADO REACTIVO
  // =========================================================

  makand = signal<MakandData | null>(null);

  tiendas = signal<TiendasData | null>(null);

  agri = signal<AgriData | null>(null);

  fuenteMakand = signal<string>(
    'VIAJEROS_Y_TERCEROS_MAKAND.xlsx'
  );

  fuenteTiendas = signal<string>(
    'ON_TIME_EN_TIENDAS.xlsx'
  );

  fuenteAgri = signal<string>(
    'ON_TIME_AGRICULTORES_2026_ACT.xlsx'
  );

  statusMakand = signal<{
    msg: string;
    cls: string;
  }>({
    msg: '',
    cls: ''
  });

  statusTiendas = signal<{
    msg: string;
    cls: string;
  }>({
    msg: '',
    cls: ''
  });

  statusAgri = signal<{
    msg: string;
    cls: string;
  }>({
    msg: '',
    cls: ''
  });

  transportadorasMakand = computed(
    () => this.makand()?.transportadoras ?? []
  );

  // =========================================================
  // INIT
  // =========================================================

  ngAfterViewInit(): void {
    // Los datos se cargan cuando el usuario selecciona el Excel.
  }

  // =========================================================
  // HELPERS GENERALES
  // =========================================================

  private fmtPct(
    pct: number | null | undefined
  ): string {

    return (
      pct === null ||
      pct === undefined ||
      isNaN(pct)
    )
      ? '—'
      : pct.toFixed(1) + '%';
  }

  pillClass(
    pct: number | null | undefined
  ): string {

    if (
      pct === null ||
      pct === undefined ||
      isNaN(pct)
    ) {
      return 'mid';
    }

    if (pct >= 70) {
      return 'good';
    }

    if (pct >= 40) {
      return 'mid';
    }

    return 'bad';
  }

  fmt(
    pct: number | null | undefined
  ): string {

    return this.fmtPct(pct);
  }

  // =========================================================
  // BUSCAR HOJA
  // =========================================================

  private findSheet(
    wb: XLSX.WorkBook,
    patterns: string[]
  ): string | null {

    for (const p of patterns) {

      const hit = wb.SheetNames.find(
        n =>
          n
            .toUpperCase()
            .replace(/\s+/g, ' ')
            .trim()
            .includes(p)
      );

      if (hit) {
        return hit;
      }
    }

    return null;
  }

  // =========================================================
  // OBTENER FILAS DE HOJA
  // =========================================================

  private sheetRows(
    wb: XLSX.WorkBook,
    name: string
  ): any[][] {

    const ws = wb.Sheets[name];

    return XLSX.utils.sheet_to_json(
      ws,
      {
        header: 1,
        defval: null,
        raw: true
      }
    ) as any[][];
  }

  // =========================================================
  // BUSCAR ÍNDICE DE COLUMNA
  // =========================================================

  private colIndex(
    header: any[],
    names: string | string[]
  ): number {

    const candidates =
      Array.isArray(names)
        ? names
        : [names];

    for (const name of candidates) {

      const idx = header.findIndex(
        h =>
          h &&
          String(h)
            .trim()
            .toUpperCase() ===
          name.toUpperCase()
      );

      if (idx >= 0) {
        return idx;
      }
    }

    return -1;
  }

  // =========================================================
  // ÚLTIMA SEMANA COMPLETA
  // =========================================================

  private pickLatestWeek(
    rowsData: any[][],
    iSem: number,
    iFecha: number
  ): number | null {

    const weeks = [
      ...new Set(
        rowsData
          .map(r => Number(r[iSem]))
          .filter(w => !isNaN(w))
      )
    ].sort((a, b) => a - b);

    if (!weeks.length) {
      return null;
    }

    if (iFecha < 0) {
      return weeks[weeks.length - 1];
    }

    for (
      let i = weeks.length - 1;
      i >= 0;
      i--
    ) {

      const w = weeks[i];

      const dates = rowsData
        .filter(
          r => Number(r[iSem]) === w
        )
        .map(r => r[iFecha])
        .filter(
          v => typeof v === 'number'
        ) as number[];

      if (dates.length) {

        const span =
          Math.max(...dates) -
          Math.min(...dates);

        if (span >= 6) {
          return w;
        }
      }
    }

    return weeks[weeks.length - 1];
  }

  // =========================================================
  // SUBIDA DE ARCHIVOS
  // =========================================================

  onFileSelected(
    evt: Event,
    kind: 'makand' | 'tiendas' | 'agri'
  ): void {

    const input =
      evt.target as HTMLInputElement;

    const file =
      input.files?.[0];

    if (!file) {
      return;
    }

    const statusSignal =
      kind === 'makand'
        ? this.statusMakand
        : kind === 'tiendas'
          ? this.statusTiendas
          : this.statusAgri;

    statusSignal.set({
      msg:
        'Leyendo ' +
        file.name +
        '…',
      cls: ''
    });

    const reader =
      new FileReader();

    reader.onload = (e) => {

      try {

        const buffer =
          e.target?.result as ArrayBuffer;

        const wb =
          XLSX.read(
            new Uint8Array(buffer),
            {
              type: 'array',
              cellDates: false
            }
          );

        if (kind === 'makand') {

          const d =
            this.parseMakand(wb);

          this.makand.set(d);

          this.fuenteMakand.set(
            file.name
          );

          this.cdr.detectChanges();

          this.renderMakandCharts(d);

        } else if (kind === 'tiendas') {

          const d =
            this.parseTiendas(wb);

          this.tiendas.set(d);

          this.fuenteTiendas.set(
            file.name
          );

          this.cdr.detectChanges();

          this.renderTiendasCharts(d);

        } else {

          const d =
            this.parseAgri(wb);

          this.agri.set(d);

          this.fuenteAgri.set(
            file.name
          );

          this.cdr.detectChanges();

          this.renderAgriCharts(d);
        }

        statusSignal.set({
          msg:
            '✓ actualizado ' +
            new Date().toLocaleString(
              'es-CO'
            ),
          cls: 'ok'
        });

      } catch (err: any) {

        console.error(err);

        statusSignal.set({
          msg:
            'Error: ' +
            err.message,
          cls: 'err'
        });
      }

      input.value = '';
    };

    reader.onerror = () => {

      statusSignal.set({
        msg:
          'No se pudo leer el archivo.',
        cls: 'err'
      });
    };

    reader.readAsArrayBuffer(file);
  }

  // =========================================================
  // PARSER MAKAND
  // =========================================================

  private parseMakand(
    wb: XLSX.WorkBook
  ): MakandData {

    const sheetName =
      this.findSheet(
        wb,
        ['DT VIAJEROS']
      );

    if (!sheetName) {

      throw new Error(
        'No encontré la hoja "DT VIAJEROS." en este archivo.'
      );
    }

    const rows =
      this.sheetRows(
        wb,
        sheetName
      );

    const header =
      rows[0];

    const iSem =
      this.colIndex(
        header,
        ['SEMANAS', 'SEMANA']
      );

    const iTransp =
      this.colIndex(
        header,
        'TRANSPORTE'
      );

    const iAlmacen =
      this.colIndex(
        header,
        'ALMACEN'
      );

    const iLlegada =
      this.colIndex(
        header,
        'CUMPLIMIENTO LLEGADA VEHICULO'
      );

    const iCargue =
      this.colIndex(
        header,
        'CUMPLIMIENTO CARGUE VEHICULO'
      );

    const iCajas =
      this.colIndex(
        header,
        'TOTAL CAJAS'
      );

    const iMes =
      this.colIndex(
        header,
        'MES'
      );

    const iFecha =
      this.colIndex(
        header,
        'FECHA DE CARGUE'
      );

    if (
      iSem < 0 ||
      iTransp < 0
    ) {

      throw new Error(
        'No encontré las columnas SEMANA/TRANSPORTE esperadas.'
      );
    }

    const data =
      rows
        .slice(1)
        .filter(
          r =>
            r &&
            r[iSem] !== null &&
            r[iTransp]
        );

    const week =
      this.pickLatestWeek(
        data,
        iSem,
        iFecha
      );

    const rowsWeek =
      data.filter(
        r =>
          Number(r[iSem]) === week
      );

    const transpMap:
      Record<
        string,
        TransportadoraStat
      > = {};

    const almacenSet =
      new Set<string>();

    rowsWeek.forEach(r => {

      const t =
        String(r[iTransp])
          .trim()
          .toUpperCase();

      if (!transpMap[t]) {

        transpMap[t] = {
          total: 0,
          llegadaOk: 0,
          cargueOk: 0,
          cajas: 0,
          almacenes: {}
        };
      }

      transpMap[t].total++;

      if (
        iLlegada >= 0 &&
        String(r[iLlegada])
          .trim()
          .toUpperCase() ===
        'CUMPLE'
      ) {

        transpMap[t]
          .llegadaOk++;
      }

      if (
        iCargue >= 0 &&
        String(r[iCargue])
          .trim()
          .toUpperCase() ===
        'CUMPLE'
      ) {

        transpMap[t]
          .cargueOk++;
      }

      if (iCajas >= 0) {

        const v =
          parseFloat(
            r[iCajas]
          );

        if (!isNaN(v)) {

          transpMap[t]
            .cajas += v;
        }
      }

      const alm =
        iAlmacen >= 0
          ? String(
              r[iAlmacen] || ''
            ).trim()
          : 'N/D';

      almacenSet.add(alm);

      transpMap[t]
        .almacenes[alm] =
        (
          transpMap[t]
            .almacenes[alm] || 0
        ) + 1;
    });

    const total =
      rowsWeek.length;

    const transportadoras =
      Object.keys(
        transpMap
      ).sort(
        (a, b) =>
          transpMap[b].total -
          transpMap[a].total
      );

    const overallLlegada =
      total
        ? (
            Object.values(
              transpMap
            ).reduce(
              (s, t) =>
                s + t.llegadaOk,
              0
            ) /
            total
          ) * 100
        : 0;

    const totalCajas =
      Object.values(
        transpMap
      ).reduce(
        (s, t) =>
          s + t.cajas,
        0
      );

    const lider =
      transportadoras[0];

    const allWeeks =
      [
        ...new Set(
          data.map(
            r => Number(r[iSem])
          )
        )
      ].sort(
        (a, b) => a - b
      );

    const trendTotal:
      number[] = [];

    const trendLlegada:
      (number | null)[] = [];

    allWeeks.forEach(w => {

      const rw =
        data.filter(
          r =>
            Number(r[iSem]) === w
        );

      trendTotal.push(
        rw.length
      );

      const ok =
        iLlegada >= 0
          ? rw.filter(
              r =>
                String(r[iLlegada])
                  .trim()
                  .toUpperCase() ===
                'CUMPLE'
            ).length
          : 0;

      trendLlegada.push(
        rw.length
          ? (
              ok /
              rw.length
            ) * 100
          : null
      );
    });

    let monthly:
      MonthlyData | null = null;

    if (iMes >= 0) {

      const mesesPresentes =
        [
          ...new Set(
            data.map(
              r =>
                String(
                  r[iMes]
                )
                  .trim()
                  .toUpperCase()
            )
          )
        ]
        .filter(
          m =>
            this.MESES_ORDER
              .includes(m)
        )
        .sort(
          (a, b) =>
            this.MESES_ORDER
              .indexOf(a) -
            this.MESES_ORDER
              .indexOf(b)
        );

      const totalPorMes:
        number[] = [];

      const llegadaPorMes:
        (number | null)[] = [];

      mesesPresentes.forEach(m => {

        const rm =
          data.filter(
            r =>
              String(
                r[iMes]
              )
                .trim()
                .toUpperCase() ===
              m
          );

        totalPorMes.push(
          rm.length
        );

        const ok =
          iLlegada >= 0
            ? rm.filter(
                r =>
                  String(
                    r[iLlegada]
                  )
                    .trim()
                    .toUpperCase() ===
                  'CUMPLE'
              ).length
            : 0;

        llegadaPorMes.push(
          rm.length
            ? (
                ok /
                rm.length
              ) * 100
            : null
        );
      });

      monthly = {
        meses:
          mesesPresentes,
        total:
          totalPorMes,
        pct:
          llegadaPorMes
      };
    }

    return {

      week,

      total,

      transportadoras,

      transpMap,

      almacenes:
        [...almacenSet],

      overallLlegada,

      totalCajas,

      lider,

      liderPct:
        total
          ? (
              transpMap[lider]
                .total /
              total
            ) * 100
          : 0,

      trend: {
        weeks:
          allWeeks,
        total:
          trendTotal,
        llegada:
          trendLlegada
      },

      monthly
    };
  }

  // =========================================================
  // PARSER TIENDAS
  // =========================================================

  private parseTiendas(
    wb: XLSX.WorkBook
  ): TiendasData {

    const sheetName =
      this.findSheet(
        wb,
        ['DETALLE REGIONES']
      );

    if (!sheetName) {

      throw new Error(
        'No encontré la hoja "Detalle regiones" en este archivo.'
      );
    }

    const rows =
      this.sheetRows(
        wb,
        sheetName
      );

    const header =
      rows[0];

    const iSem =
      this.colIndex(
        header,
        [
          'SEMANA',
          'SEMANAS'
        ]
      );

    const iCump =
      this.colIndex(
        header,
        'CUMPLIMIENTO'
      );

    const iRegion =
      this.colIndex(
        header,
        'REGIÓN'
      );

    const iCedi =
      this.colIndex(
        header,
        'CEDI'
      );

    const iMes =
      this.colIndex(
        header,
        'MES'
      );

    const iFecha =
      this.colIndex(
        header,
        'FECHA'
      );

    if (
      iSem < 0 ||
      iCump < 0
    ) {

      throw new Error(
        'No encontré las columnas SEMANA/CUMPLIMIENTO esperadas.'
      );
    }

    const data =
      rows
        .slice(1)
        .filter(
          r =>
            r &&
            r[iSem] !== null
        );

    const week =
      this.pickLatestWeek(
        data,
        iSem,
        iFecha
      );

    const rowsWeek =
      data.filter(
        r =>
          Number(r[iSem]) === week
      );

    const dist:
      Record<string, number> = {};

    const cediMap:
      Record<string, number> = {};

    const regionMap:
      Record<string, RegionStat> = {};

    rowsWeek.forEach(r => {

      const c =
        iCump >= 0
          ? String(
              r[iCump] || ''
            )
              .trim()
              .toUpperCase()
          : 'N/D';

      dist[c] =
        (dist[c] || 0) + 1;

      if (iCedi >= 0) {

        const cd =
          String(
            r[iCedi] || ''
          ).trim() ||
          'N/D';

        cediMap[cd] =
          (cediMap[cd] || 0) + 1;
      }

      if (iRegion >= 0) {

        const rg =
          String(
            r[iRegion] || ''
          ).trim() ||
          'N/D';

        if (!regionMap[rg]) {

          regionMap[rg] = {
            total: 0,
            cumple: 0
          };
        }

        regionMap[rg].total++;

        if (
          c === 'CUMPLE'
        ) {

          regionMap[rg]
            .cumple++;
        }
      }
    });

    const total =
      rowsWeek.length;

    const topRegiones =
      Object.keys(
        regionMap
      )
      .sort(
        (a, b) =>
          regionMap[b].total -
          regionMap[a].total
      )
      .slice(0, 10);

    const cedis =
      Object.keys(
        cediMap
      ).sort(
        (a, b) =>
          cediMap[b] -
          cediMap[a]
      );

    const allWeeks =
      [
        ...new Set(
          data.map(
            r => Number(r[iSem])
          )
        )
      ].sort(
        (a, b) => a - b
      );

    const trendCumple:
      (number | null)[] = [];

    const trendParcial:
      (number | null)[] = [];

    const trendNoCumple:
      (number | null)[] = [];

    const trendTotal:
      number[] = [];

    allWeeks.forEach(w => {

      const rw =
        data.filter(
          r =>
            Number(r[iSem]) === w
        );

      const t =
        rw.length;

      trendTotal.push(t);

      const c =
        rw.filter(
          r =>
            String(
              r[iCump] || ''
            )
              .trim()
              .toUpperCase() ===
            'CUMPLE'
        ).length;

      const p =
        rw.filter(
          r =>
            String(
              r[iCump] || ''
            )
              .trim()
              .toUpperCase() ===
            'CUMPLE PARCIAL'
        ).length;

      const n =
        rw.filter(
          r =>
            String(
              r[iCump] || ''
            )
              .trim()
              .toUpperCase() ===
            'NO CUMPLE'
        ).length;

      trendCumple.push(
        t
          ? (c / t) * 100
          : null
      );

      trendParcial.push(
        t
          ? (p / t) * 100
          : null
      );

      trendNoCumple.push(
        t
          ? (n / t) * 100
          : null
      );
    });

    let monthly:
      MonthlyData | null = null;

    if (iMes >= 0) {

      const mesesPresentes =
        [
          ...new Set(
            data.map(
              r =>
                String(
                  r[iMes]
                )
                  .trim()
                  .toUpperCase()
            )
          )
        ]
        .filter(
          m =>
            this.MESES_ORDER
              .includes(m)
        )
        .sort(
          (a, b) =>
            this.MESES_ORDER
              .indexOf(a) -
            this.MESES_ORDER
              .indexOf(b)
        );

      const totalPorMes:
        number[] = [];

      const cumplePorMes:
        (number | null)[] = [];

      mesesPresentes.forEach(m => {

        const rm =
          data.filter(
            r =>
              String(
                r[iMes]
              )
                .trim()
                .toUpperCase() ===
              m
          );

        totalPorMes.push(
          rm.length
        );

        const ok =
          rm.filter(
            r =>
              String(
                r[iCump] || ''
              )
                .trim()
                .toUpperCase() ===
              'CUMPLE'
          ).length;

        cumplePorMes.push(
          rm.length
            ? (
                ok /
                rm.length
              ) * 100
            : null
        );
      });

      monthly = {
        meses:
          mesesPresentes,
        total:
          totalPorMes,
        pct:
          cumplePorMes
      };
    }

    return {

      week,

      total,

      dist,

      cediMap,

      cedis,

      regionMap,

      topRegiones,

      trend: {

        weeks:
          allWeeks,

        cumple:
          trendCumple,

        parcial:
          trendParcial,

        noCumple:
          trendNoCumple,

        total:
          trendTotal
      },

      monthly
    };
  }

  // =========================================================
  // PARSER ON TIME AGRICULTORES
  // =========================================================

  private parseAgri(
    wb: XLSX.WorkBook
  ): AgriData {

    const sheetName =
      this.findSheet(
        wb,
        ['TIEMPO AGRICULTORES']
      );

    if (!sheetName) {

      throw new Error(
        'No encontré la hoja "TIEMPO AGRICULTORES" en este archivo.'
      );
    }

    const rows =
      this.sheetRows(
        wb,
        sheetName
      );

    const header =
      rows[0];

    const iSem =
      this.colIndex(
        header,
        [
          'SEMANA',
          'SEMANAS'
        ]
      );

    const iAgricultor =
      this.colIndex(
        header,
        'AGRICULTOR'
      );

    const iLlegada =
      this.colIndex(
        header,
        'CUMPLIMIENTO LLEGADA AGRICULTOR'
      );

    const iTiempo =
      this.colIndex(
        header,
        'CUMPLIMIENTO TIEMPO EN AGRICULTOR'
      );

    const iPlanta =
      this.colIndex(
        header,
        'CUMPLIMIENTO LLEGADA A PLANTA'
      );

    const iMes =
      this.colIndex(
        header,
        'MES'
      );

    const iFecha =
      this.colIndex(
        header,
        'FECHA'
      );

    if (
      iSem < 0 ||
      iAgricultor < 0
    ) {

      throw new Error(
        'No encontré las columnas SEMANA/AGRICULTOR esperadas en "TIEMPO AGRICULTORES".'
      );
    }

    // =========================================================
    // EXCLUIR REGISTROS
    // =========================================================

    const EXCLUIR = [
      'LECHUGAS DEL DIA',
      'LECHUGAS DEL DÍA'
    ];

    const normName = (
      s: string
    ) => {

      return s
        .normalize('NFD')
        .replace(
          /[\u0300-\u036f]/g,
          ''
        )
        .trim()
        .toUpperCase();
    };

    const displayName = (
      raw: string
    ): string => {

      const n =
        normName(raw);

      if (
        n.startsWith(
          'FERRUCAS'
        )
      ) {

        return 'FERRUCAS';
      }

      return String(
        raw
      ).trim();
    };

    const data =
      rows
        .slice(1)
        .filter(
          r =>
            r &&
            r[iSem] !== null &&
            r[iAgricultor] &&
            !EXCLUIR.includes(
              normName(
                r[iAgricultor]
              )
            )
        );

    // =========================================================
    // SEMANA ACTUAL
    // =========================================================

    const week =
      this.pickLatestWeek(
        data,
        iSem,
        iFecha
      );

    const rowsWeek =
      data.filter(
        r =>
          Number(r[iSem]) === week
      );

    // =========================================================
    // AGRUPAR POR AGRICULTOR
    // =========================================================

    interface AgriAgg {

      viajes: number;

      llegadaOk: number;

      llegadaTot: number;

      tiempoOk: number;

      tiempoTot: number;

      plantaOk: number;

      plantaTot: number;
    }

    const agriMap:
      Record<
        string,
        AgriAgg
      > = {};

    rowsWeek.forEach(r => {

      const name =
        displayName(
          r[iAgricultor]
        );

      if (!agriMap[name]) {

        agriMap[name] = {

          viajes: 0,

          llegadaOk: 0,

          llegadaTot: 0,

          tiempoOk: 0,

          tiempoTot: 0,

          plantaOk: 0,

          plantaTot: 0
        };
      }

      agriMap[name]
        .viajes++;

      const vL =
        iLlegada >= 0
          ? String(
              r[iLlegada] || ''
            )
              .trim()
              .toUpperCase()
          : '';

      if (
        vL === 'CUMPLE' ||
        vL === 'NO CUMPLE'
      ) {

        agriMap[name]
          .llegadaTot++;

        if (
          vL === 'CUMPLE'
        ) {

          agriMap[name]
            .llegadaOk++;
        }
      }

      const vT =
        iTiempo >= 0
          ? String(
              r[iTiempo] || ''
            )
              .trim()
              .toUpperCase()
          : '';

      if (
        vT === 'CUMPLE' ||
        vT === 'NO CUMPLE'
      ) {

        agriMap[name]
          .tiempoTot++;

        if (
          vT === 'CUMPLE'
        ) {

          agriMap[name]
            .tiempoOk++;
        }
      }

      const vP =
        iPlanta >= 0
          ? String(
              r[iPlanta] || ''
            )
              .trim()
              .toUpperCase()
          : '';

      if (
        vP === 'CUMPLE' ||
        vP === 'NO CUMPLE'
      ) {

        agriMap[name]
          .plantaTot++;

        if (
          vP === 'CUMPLE'
        ) {

          agriMap[name]
            .plantaOk++;
        }
      }
    });

    const agricultores:
      AgricultorRow[] =
      Object.keys(
        agriMap
      )
      .sort(
        (a, b) =>
          agriMap[b].viajes -
          agriMap[a].viajes
      )
      .map(
        name => {

          const m =
            agriMap[name];

          return {

            name,

            viajes:
              m.viajes,

            pctLlegada:
              m.llegadaTot
                ? (
                    m.llegadaOk /
                    m.llegadaTot
                  ) * 100
                : null,

            pctTiempo:
              m.tiempoTot
                ? (
                    m.tiempoOk /
                    m.tiempoTot
                  ) * 100
                : null,

            pctPlanta:
              m.plantaTot
                ? (
                    m.plantaOk /
                    m.plantaTot
                  ) * 100
                : null
          };
        }
      );

    const totalViajes =
      rowsWeek.length;

    // =========================================================
    // ESTADÍSTICAS GENERALES
    // =========================================================

    const statFrom = (
      field: number
    ): MetricStat => {

      let ok = 0;

      let tot = 0;

      rowsWeek.forEach(r => {

        const v =
          String(
            r[field] || ''
          )
            .trim()
            .toUpperCase();

        if (
          v === 'CUMPLE' ||
          v === 'NO CUMPLE'
        ) {

          tot++;

          if (
            v === 'CUMPLE'
          ) {

            ok++;
          }
        }
      });

      return {

        avg:
          tot
            ? (
                ok /
                tot
              ) * 100
            : null,

        count:
          ok
      };
    };

    // =========================================================
    // TENDENCIA SEMANAL
    // =========================================================

    const allWeeks =
      [
        ...new Set(
          data.map(
            r => Number(
              r[iSem]
            )
          )
        )
      ].sort(
        (a, b) => a - b
      );

    const trendSeries = (
      field: number
    ): (number | null)[] => {

      return allWeeks.map(
        w => {

          const rw =
            data.filter(
              r =>
                Number(
                  r[iSem]
                ) === w
            );

          let ok = 0;

          let tot = 0;

          rw.forEach(r => {

            const v =
              String(
                r[field] || ''
              )
                .trim()
                .toUpperCase();

            if (
              v === 'CUMPLE' ||
              v === 'NO CUMPLE'
            ) {

              tot++;

              if (
                v === 'CUMPLE'
              ) {

                ok++;
              }
            }
          });

          return tot
            ? (
                ok /
                tot
              ) * 100
            : null;
        }
      );
    };

    const trend:
      TrendAgri = {

        weeks:
          allWeeks.map(
            String
          ),

        llegada:
          trendSeries(
            iLlegada
          ),

        tiempo:
          trendSeries(
            iTiempo
          ),

        planta:
          trendSeries(
            iPlanta
          )
      };

    // =========================================================
    // FUNCIÓN PARA OBTENER MES DESDE FECHA
    // =========================================================

    const getMesDesdeFecha = (
      valor: any
    ): string | null => {

      if (
        valor === null ||
        valor === undefined ||
        valor === ''
      ) {

        return null;
      }

      let fecha:
        Date | null = null;

      // -----------------------------------------
      // FECHA SERIAL DE EXCEL
      // -----------------------------------------

      if (
        typeof valor === 'number'
      ) {

        const fechaExcel =
          XLSX.SSF.parse_date_code(
            valor
          );

        if (!fechaExcel) {
          return null;
        }

        fecha =
          new Date(
            fechaExcel.y,
            fechaExcel.m - 1,
            fechaExcel.d
          );
      }

      // -----------------------------------------
      // FECHA COMO DATE
      // -----------------------------------------

      else if (
        valor instanceof Date
      ) {

        fecha = valor;
      }

      // -----------------------------------------
      // FECHA COMO STRING
      // -----------------------------------------

      else {

        const texto =
          String(valor)
            .trim();

        // Formato YYYY-MM-DD
        if (
          /^\d{4}-\d{1,2}-\d{1,2}$/
            .test(texto)
        ) {

          const partes =
            texto.split('-');

          fecha =
            new Date(
              Number(partes[0]),
              Number(partes[1]) - 1,
              Number(partes[2])
            );
        }

        // Formato DD/MM/YYYY
        else if (
          /^\d{1,2}\/\d{1,2}\/\d{4}$/
            .test(texto)
        ) {

          const partes =
            texto.split('/');

          fecha =
            new Date(
              Number(partes[2]),
              Number(partes[1]) - 1,
              Number(partes[0])
            );
        }

        // Intento general
        else {

          const parsed =
            new Date(texto);

          if (
            !isNaN(
              parsed.getTime()
            )
          ) {

            fecha = parsed;
          }
        }
      }

      if (
        !fecha ||
        isNaN(
          fecha.getTime()
        )
      ) {

        return null;
      }

      return this.MESES_ORDER[
        fecha.getMonth()
      ];
    };

    // =========================================================
    // RESUMEN MENSUAL AGRICULTORES
    // =========================================================

    let monthly:
      MonthlyData | null = null;

    /*
     * IMPORTANTE:
     *
     * Primero intenta usar la columna MES.
     *
     * Si MES está vacía o tiene un valor que no coincide
     * con ENERO, FEBRERO, MARZO, etc., entonces intenta
     * obtener el mes directamente desde FECHA.
     */

    const mesesMap:
      Record<
        string,
        any[] 
      > = {};

    data.forEach(r => {

      let mes:
        string | null = null;

      // -----------------------------------------
      // OPCIÓN 1: COLUMNA MES
      // -----------------------------------------

      if (
        iMes >= 0 &&
        r[iMes] !== null &&
        r[iMes] !== undefined &&
        r[iMes] !== ''
      ) {

        const valorMes =
          String(
            r[iMes]
          )
            .trim()
            .toUpperCase();

        // Si viene directamente como ENERO
        if (
          this.MESES_ORDER
            .includes(valorMes)
        ) {

          mes =
            valorMes;
        }

        // Si viene como número 1-12
        else if (
          !isNaN(
            Number(valorMes)
          )
        ) {

          const numeroMes =
            Number(valorMes);

          if (
            numeroMes >= 1 &&
            numeroMes <= 12
          ) {

            mes =
              this.MESES_ORDER[
                numeroMes - 1
              ];
          }
        }
      }

      // -----------------------------------------
      // OPCIÓN 2: OBTENER MES DESDE FECHA
      // -----------------------------------------

      if (
        !mes &&
        iFecha >= 0
      ) {

        mes =
          getMesDesdeFecha(
            r[iFecha]
          );
      }

      // -----------------------------------------
      // GUARDAR REGISTRO EN SU MES
      // -----------------------------------------

      if (mes) {

        if (
          !mesesMap[mes]
        ) {

          mesesMap[mes] = [];
        }

        mesesMap[mes]
          .push(r);
      }
    });

    // =========================================================
    // ORDENAR MESES
    // =========================================================

    const mesesPresentes =
      Object.keys(
        mesesMap
      )
      .sort(
        (a, b) =>
          this.MESES_ORDER
            .indexOf(a) -
          this.MESES_ORDER
            .indexOf(b)
      );

    const totalPorMes:
      number[] = [];

    const llegadaPorMes:
      (number | null)[] = [];

    // =========================================================
    // CALCULAR REGISTROS Y % LLEGADA POR MES
    // =========================================================

    mesesPresentes.forEach(
      mes => {

        const rm =
          mesesMap[mes];

        // -----------------------------
        // TOTAL DE REGISTROS
        // -----------------------------

        totalPorMes.push(
          rm.length
        );

        // -----------------------------
        // CUMPLIMIENTO DE LLEGADA
        // -----------------------------

        let ok = 0;

        let tot = 0;

        rm.forEach(r => {

          if (
            iLlegada < 0
          ) {

            return;
          }

          const valor =
            String(
              r[iLlegada] || ''
            )
              .trim()
              .toUpperCase();

          if (
            valor === 'CUMPLE' ||
            valor === 'NO CUMPLE'
          ) {

            tot++;

            if (
              valor === 'CUMPLE'
            ) {

              ok++;
            }
          }
        });

        llegadaPorMes.push(

          tot > 0

            ? (
                ok /
                tot
              ) * 100

            : null

        );
      }
    );

    // =========================================================
    // CREAR OBJETO MONTHLY
    // =========================================================

    if (
      mesesPresentes.length > 0
    ) {

      monthly = {

        meses:
          mesesPresentes,

        total:
          totalPorMes,

        pct:
          llegadaPorMes
      };
    }

    // =========================================================
    // DEBUG TEMPORAL
    // =========================================================

    console.log(
      'ON TIME AGRICULTORES - RESUMEN MENSUAL:',
      monthly
    );

    console.log(
      'ON TIME AGRICULTORES - MESES:',
      mesesPresentes
    );

    console.log(
      'ON TIME AGRICULTORES - REGISTROS:',
      totalPorMes
    );

    console.log(
      'ON TIME AGRICULTORES - % LLEGADA:',
      llegadaPorMes
    );

    // =========================================================
    // RETORNO
    // =========================================================

    return {

      weekLabel:
        String(
          week ?? ''
        ),

      agricultores,

      totalViajes,

      llegada:
        statFrom(
          iLlegada
        ),

      tiempo:
        statFrom(
          iTiempo
        ),

      planta:
        statFrom(
          iPlanta
        ),

      trend,

      monthly
    };
  }

  // =========================================================
  // OCULTAR GRÁFICA VACÍA
  // =========================================================

  private hideEmpty(
    canvas: HTMLCanvasElement
  ): void {

    const empty =
      canvas.parentElement
        ?.querySelector(
          '.chart-empty'
        ) as HTMLElement | null;

    if (empty) {

      empty.style.display =
        'none';
    }
  }

  // =========================================================
  // GRÁFICA DOUGHNUT
  // =========================================================

  private makeDoughnut(
    canvas: HTMLCanvasElement,
    id: string,
    labels: string[],
    data: number[],
    colors: string[]
  ): void {

    this.hideEmpty(canvas);

    this.charts[id]
      ?.destroy();

    this.charts[id] =
      new Chart(
        canvas,
        {

          type: 'doughnut',

          data: {

            labels,

            datasets: [

              {

                data,

                backgroundColor:
                  colors,

                borderColor:
                  '#FFFEFB',

                borderWidth:
                  3
              }
            ]
          },

          options: {

            plugins: {

              legend: {

                position:
                  'bottom',

                labels: {

                  boxWidth:
                    10,

                  padding:
                    16
                }
              }
            },

            cutout:
              '62%'
          }
        }
      );
  }

  // =========================================================
  // GRÁFICA HORIZONTAL
  // =========================================================

  private makeHBar(
    canvas: HTMLCanvasElement,
    id: string,
    labels: string[],
    data: number[],
    colors: string[]
  ): void {

    this.hideEmpty(canvas);

    this.charts[id]
      ?.destroy();

    this.charts[id] =
      new Chart(
        canvas,
        {

          type: 'bar',

          data: {

            labels,

            datasets: [

              {

                data,

                backgroundColor:
                  colors
              }
            ]
          },

          options: {

            responsive:
              true,

            maintainAspectRatio:
              false,

            indexAxis:
              'y',

            scales: {

              x: {

                max:
                  100,

                grid: {

                  color:
                    '#E4DAC3'
                }
              },

              y: {

                grid: {

                  display:
                    false
                }
              }
            },

            plugins: {

              legend: {

                display:
                  false
              }
            }
          }
        }
      );
  }

  // =========================================================
  // GRÁFICA LINEAL
  // =========================================================

  private makeLine(
    canvas: HTMLCanvasElement,
    id: string,
    labels: string[],
    datasets: any[],
    dualAxis: boolean
  ): void {

    this.hideEmpty(canvas);

    this.charts[id]
      ?.destroy();

    const scales: any =
      dualAxis

        ? {

            x: {

              grid: {

                display:
                  false
              }
            },

            y: {

              position:
                'left',

              grid: {

                color:
                  '#E4DAC3'
              },

              title: {

                display:
                  true,

                text:
                  'Viajes / registros',

                font: {

                  size:
                    10
                }
              }
            },

            y1: {

              position:
                'right',

              min:
                0,

              max:
                100,

              grid: {

                display:
                  false
              },

              title: {

                display:
                  true,

                text:
                  '% cumplimiento',

                font: {

                  size:
                    10
                }
              }
            }
          }

        : {

            x: {

              grid: {

                display:
                  false
              }
            },

            y: {

              min:
                0,

              max:
                100,

              grid: {

                color:
                  '#E4DAC3'
              }
            }
          };

    this.charts[id] =
      new Chart(
        canvas,
        {

          type: 'line',

          data: {

            labels,

            datasets
          },

          options: {

            responsive:
              true,

            maintainAspectRatio:
              false,

            interaction: {

              mode:
                'index',

              intersect:
                false
            },

            scales,

            plugins: {

              legend: {

                position:
                  'bottom',

                labels: {

                  boxWidth:
                    10,

                  padding:
                    14
                }
              }
            }
          }
        }
      );
  }

  // =========================================================
  // GRÁFICA RESUMEN MENSUAL
  // =========================================================

  private renderMonthly(
    canvas: HTMLCanvasElement,
    id: string,
    m: MonthlyData,
    labelTotal: string,
    labelPct: string
  ): void {

    const labels =
      m.meses.map(
        mes =>
          mes.charAt(0) +
          mes
            .slice(1)
            .toLowerCase()
      );

    this.makeLine(
      canvas,
      id,
      labels,
      [

        {

          label:
            labelTotal,

          data:
            m.total,

          borderColor:
            '#223A57',

          backgroundColor:
            '#223A57',

          tension:
            .3,

          yAxisID:
            'y',

          type:
            'bar'
        },

        {

          label:
            labelPct,

          data:
            m.pct,

          borderColor:
            '#2F6F4E',

          backgroundColor:
            '#2F6F4E',

          tension:
            .3,

          borderDash:
            [4, 3],

          yAxisID:
            'y1'
        }
      ],

      true
    );
  }

  // =========================================================
  // UTILIDADES DE MESES
  // =========================================================

  esMesActual(
    mes: string
  ): boolean {

    return (
      mes ===
      this.CURRENT_MONTH_NAME
    );
  }

  nombreMes(
    mes: string
  ): string {

    return (
      mes.charAt(0) +
      mes
        .slice(1)
        .toLowerCase()
    );
  }

  // =========================================================
  // RENDER MAKAND
  // =========================================================

  private renderMakandCharts(
    d: MakandData
  ): void {

    if (
      !this.chartTranspRef
    ) {

      return;
    }

    this.makeDoughnut(

      this.chartTranspRef
        .nativeElement,

      'chartTransp',

      d.transportadoras,

      d.transportadoras.map(
        t =>
          d.transpMap[t]
            .total
      ),

      d.transportadoras.map(
        (_, i) =>
          this.PALETTE[
            i %
            this.PALETTE.length
          ]
      )
    );

    this.makeLine(

      this.chartMakandTrendRef
        .nativeElement,

      'chartMakandTrend',

      d.trend.weeks.map(
        w =>
          'S' + w
      ),

      [

        {

          label:
            'Viajes totales',

          data:
            d.trend.total,

          borderColor:
            '#223A57',

          backgroundColor:
            '#223A57',

          tension:
            .3,

          yAxisID:
            'y'
        },

        {

          label:
            '% Cumplimiento llegada',

          data:
            d.trend.llegada,

          borderColor:
            '#2F6F4E',

          backgroundColor:
            '#2F6F4E',

          tension:
            .3,

          borderDash:
            [4, 3],

          yAxisID:
            'y1'
        }
      ],

      true
    );

    if (
      d.monthly
    ) {

      this.renderMonthly(

        this.chartMakandMesRef.nativeElement,

        'chartMakandMes',

        d.monthly,

        'Viajes',

        '% Cumple llegada'
      );
    }
  }

  // =========================================================
  // RENDER TIENDAS
  // =========================================================

  private renderTiendasCharts(
    d: TiendasData
  ): void {

    if (
      !this.chartTiendasRef
    ) {

      return;
    }

    const labels =
      Object.keys(
        d.dist
      );

    this.makeDoughnut(

      this.chartTiendasRef
        .nativeElement,

      'chartTiendas',

      labels,

      labels.map(
        l =>
          d.dist[l]
      ),

      labels.map(
        (l, i) =>

          l === 'CUMPLE'

            ? '#2F6F4E'

            : l ===
                'CUMPLE PARCIAL'

              ? '#BE8A2E'

              : l ===
                  'NO CUMPLE'

                ? '#B5453A'

                : this.PALETTE[
                    i %
                    this.PALETTE.length
                  ]
      )
    );

    this.makeLine(

      this.chartTiendasTrendRef
        .nativeElement,

      'chartTiendasTrend',

      d.trend.weeks.map(
        w =>
          'S' + w
      ),

      [

        {

          label:
            '% Cumple',

          data:
            d.trend.cumple,

          borderColor:
            '#2F6F4E',

          backgroundColor:
            '#2F6F4E',

          tension:
            .3
        },

        {

          label:
            '% Cumple parcial',

          data:
            d.trend.parcial,

          borderColor:
            '#BE8A2E',

          backgroundColor:
            '#BE8A2E',

          tension:
            .3
        },

        {

          label:
            '% No cumple',

          data:
            d.trend.noCumple,

          borderColor:
            '#B5453A',

          backgroundColor:
            '#B5453A',

          tension:
            .3
        }
      ],

      false
    );

    if (
      d.monthly
    ) {

      this.renderMonthly(

        this.chartTiendasMesRef.nativeElement,

        'chartTiendasMes',

        d.monthly,

        'Registros',

        '% Cumple'
      );
    }
  }

  // =========================================================
  // RENDER ON TIME AGRICULTORES
  // =========================================================

  private renderAgriCharts(
    d: AgriData
  ): void {

    if (
      !this.chartAgriRef
    ) {

      return;
    }

    this.makeHBar(

      this.chartAgriRef
        .nativeElement,

      'chartAgri',

      d.agricultores.map(
        a =>
          a.name
      ),

      d.agricultores.map(
        a =>
          a.pctLlegada || 0
      ),

      d.agricultores.map(
        a =>

          this.pillClass(
            a.pctLlegada
          ) === 'good'

            ? '#2F6F4E'

            : this.pillClass(
                a.pctLlegada
              ) === 'mid'

              ? '#BE8A2E'

              : '#B5453A'
      )
    );

    this.makeLine(

      this.chartAgriTrendRef
        .nativeElement,

      'chartAgriTrend',

      d.trend.weeks.map(
        w =>
          'S' + w
      ),

      [

        {

          label:
            '% Cumple llegada',

          data:
            d.trend.llegada,

          borderColor:
            '#223A57',

          backgroundColor:
            '#223A57',

          tension:
            .3
        },

        {

          label:
            '% Cumple tiempo en finca',

          data:
            d.trend.tiempo,

          borderColor:
            '#BE8A2E',

          backgroundColor:
            '#BE8A2E',

          tension:
            .3
        },

        {

          label:
            '% Cumple llegada a planta',

          data:
            d.trend.planta,

          borderColor:
            '#B5453A',

          backgroundColor:
            '#B5453A',

          tension:
            .3
        }
      ],

      false
    );

    // =========================================================
    // AQUÍ SE GENERA EL RESUMEN MENSUAL DE AGRICULTORES
    // =========================================================

    if (
      d.monthly &&
      d.monthly.meses.length > 0
    ) {

      this.renderMonthly(

        this.chartAgriMesRef
          .nativeElement,

        'chartAgriMes',

        d.monthly,

        'Registros',

        '% Cumple llegada'
      );

    } else {

      console.warn(
        'No hay datos mensuales para ON TIME AGRICULTORES'
      );
    }
  }
}