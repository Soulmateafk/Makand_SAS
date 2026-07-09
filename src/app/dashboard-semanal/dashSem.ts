import { Component, ElementRef, ViewChild, AfterViewInit, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

// Interfaces internas para mantener el tipado estricto
interface MakandTableRow {
  name: string;
  total: number;
  pct: string;
  llegada: string;
  llegadaClass: string;
  cargue: string;
  cargueClass: string;
  color: string;
}

interface TiendasCediRow {
  name: string;
  total: number;
  pct: string;
}

interface TiendasRutaRow {
  name: string;
  total: number;
  pct: string;
  pctClass: string;
}

interface AgriTableRow {
  name: string;
  total: number;
  llegada: string;
  llegadaClass: string;
  tiempo: string;
  tiempoClass: string;
  planta: string;
  plantaClass: string;
}

@Component({
  selector: 'app-dash-semanal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashSem.html',
  styleUrl: './dashSem.css'
})
export class DashSemanalComponent implements AfterViewInit {
  @ViewChild('chartTransp') chartTranspRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAlmacen') chartAlmacenRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartTiendas') chartTiendasRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAgri') chartAgriRef!: ElementRef<HTMLCanvasElement>;

  private chartTransp?: Chart;
  private chartAlmacen?: Chart;
  private chartTiendas?: Chart;
  private chartAgri?: Chart;

  constructor(private cdr: ChangeDetectorRef) {}

  // --- MÓDULOS DE ESTADO ---
  makandState = {
    src: 'Ningún archivo seleccionado',
    status: 'Pendiente de carga',
    statusClass: 'badge-pendiente',
    note: 'Sube el reporte de "VIAJEROS Y TERCEROS" para calcular viajes e ingresos a planta.',
    kpiTotal: 0,
    kpiPeriodo: 'Sin datos',
    kpiCumple: '0%',
    kpiCajas: 0 as any,
    kpiCajasSub: '0 cajas totales',
    kpiLider: 'Ninguna',
    kpiLiderSub: 'Esperando registros',
    table: [] as MakandTableRow[]
  };

  tiendasState = {
    src: 'Ningún archivo seleccionado',
    status: 'Pendiente de carga',
    statusClass: 'badge-pendiente',
    note: 'Sube el reporte "ON TIME EN TIENDAS" para validar las ventanas de entrega por CEDI.',
    kpiTotal: 0,
    kpiPeriodo: 'Sin datos',
    kpiCumple: '0',
    kpiCumpleSub: '0% de efectividad',
    kpiParcial: '0',
    kpiParcialSub: 'Entregas toleradas',
    kpiNoCumple: '0',
    kpiNoCumpleSub: '0% de fallas',
    tableCedi: [] as TiendasCediRow[],
    rutasMesLabel: '',
    tableRutas: [] as TiendasRutaRow[]
  };

  agriState = {
    src: 'Ningún archivo seleccionado',
    status: 'Pendiente de carga',
    statusClass: 'badge-pendiente',
    note: 'Sube el reporte "ON TIME AGRICULTORES" para evaluar los tres frentes de recolección.',
    kpiSemanaLabel: '',
    kpiTotal: 0,
    kpiTotalSub: 'Viajes en la semana',
    kpiLlegada: 0 as number | string, 
    kpiTiempo: 0 as number | string,
    kpiPlanta: 0 as number | string,
    chartLabel: '',
    tableLabel: 'Pendiente',
    table: [] as AgriTableRow[]
  };

  ngAfterViewInit() {
    this.inicializarGraficosVacios();
  }

  handleUpload(event: Event, tipo: 'makand' | 'tiendas' | 'agri') {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    const file = input.files[0];
    const reader = new FileReader();

    if (tipo === 'makand') this.makandState.src = file.name;
    if (tipo === 'tiendas') this.tiendasState.src = file.name;
    if (tipo === 'agri') this.agriState.src = file.name;

    reader.onload = (e: ProgressEvent<FileReader>) => {
      const target = e.target;
      if (!target || !target.result) return;

      const data = new Uint8Array(target.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      // Lógica ESTRICTA de búsqueda de pestañas para evitar cargar hojas vacías
      let sheetName = workbook.SheetNames[0]; 
      
      if (tipo === 'makand') {
        sheetName = workbook.SheetNames.find(n => 
          n.toLowerCase().includes('tercero') || n.toLowerCase().includes('viajero')
        ) || workbook.SheetNames[0];
      } else if (tipo === 'tiendas') {
        sheetName = workbook.SheetNames.find(n => 
          n.toLowerCase().includes('dt') || n.toLowerCase().includes('detalle') || n.toLowerCase().includes('tiendas')
        ) || workbook.SheetNames[0];
      } else if (tipo === 'agri') {
        sheetName = workbook.SheetNames.find(n => 
          (n.toLowerCase().includes('agri') || n.toLowerCase().includes('finca') || n.toLowerCase().includes('campo')) && 
          !n.toLowerCase().includes('canastilla')
        ) || workbook.SheetNames[0];
      }

      const worksheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json<any>(worksheet, { defval: '' });

      if (jsonRows.length > 0) {
        if (tipo === 'makand') this.procesarReporteMakand(jsonRows);
        if (tipo === 'tiendas') this.procesarReporteTiendas(jsonRows);
        if (tipo === 'agri') this.procesarReporteAgricultores(jsonRows);

        this.cdr.detectChanges();
      }
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * PROCESAMIENTO 01: MAKAND
   */
  private procesarReporteMakand(rows: any[]) {
    // Normalización de cabeceras
    const datosLimpios = rows.reduce((acc: any[], row: any) => {
      const cleanRow: any = {};
      Object.keys(row).forEach(key => { if (key) cleanRow[key.trim().toUpperCase()] = row[key]; });
      
      const transp = cleanRow['TRANSPORTE'] || cleanRow['TRANSPORTADORA'];
      if (transp) {
        cleanRow['_transpNormalizado'] = transp;
        acc.push(cleanRow);
      }
      return acc;
    }, []);

    if (datosLimpios.length === 0) return;

    this.makandState.status = 'Cargado';
    this.makandState.statusClass = 'text-success fw-bold';
    this.makandState.note = 'Datos consolidados en base al histórico de transportadoras externas y flota propia.';

    let totalViajes = 0;
    let totalCajas = 0;
    let viajesCumpleLlegada = 0;
    const transportadoras: Record<string, { total: number; cumpleLlegada: number; cumpleCargue: number }> = {};
    const almacenes: Record<string, number> = {};
    let mesDetectado = 'Julio 2026';

    datosLimpios.forEach((cleanRow: any) => {
      const transp = cleanRow['_transpNormalizado'];
      totalViajes++;
      if (cleanRow['MES']) mesDetectado = cleanRow['MES'];

      const cajas = parseFloat(cleanRow['TOTAL CAJAS'] || 0);
      totalCajas += isNaN(cajas) ? 0 : cajas;

      if (!transportadoras[transp]) transportadoras[transp] = { total: 0, cumpleLlegada: 0, cumpleCargue: 0 };
      transportadoras[transp].total++;

      const cumpleLlegadaStr = String(cleanRow['CUMPLIMIENTO LLEGADA VEHICULO'] || '').toUpperCase();
      if (cumpleLlegadaStr.includes('CUMPLE') && !cumpleLlegadaStr.includes('NO')) {
        transportadoras[transp].cumpleLlegada++;
        viajesCumpleLlegada++;
      }

      const cumpleCargueStr = String(cleanRow['CUMPLIMIENTO CARGUE VEHICULO'] || '').toUpperCase();
      if (cumpleCargueStr.includes('CUMPLE') && !cumpleCargueStr.includes('NO')) {
        transportadoras[transp].cumpleCargue++;
      }

      const almacen = cleanRow['ALMACEN'] || cleanRow['ALMACÉN'] || 'Otros';
      almacenes[almacen] = (almacenes[almacen] || 0) + 1;
    });

    this.makandState.kpiTotal = totalViajes;
    this.makandState.kpiPeriodo = `Periodo operativo: ${mesDetectado}`;
    this.makandState.kpiCumple = totalViajes > 0 ? `${Math.round((viajesCumpleLlegada / totalViajes) * 100)}%` : '0%';
    this.makandState.kpiCajas = totalCajas.toLocaleString('de-DE') as any;

    const colores = ['#2a5298', '#11998e', '#ff416c', '#f5af19', '#8e44ad'];
    let idx = 0; let maxViajes = -1; let liderTransp = 'Ninguna';

    this.makandState.table = Object.keys(transportadoras).map(name => {
      const t = transportadoras[name];
      const pctValue = ((t.total / totalViajes) * 100).toFixed(1);
      if (t.total > maxViajes) { maxViajes = t.total; liderTransp = name; }
      
      const pctLlegada = Math.round((t.cumpleLlegada / t.total) * 100);
      const pctCargue = Math.round((t.cumpleCargue / t.total) * 100);

      return {
        name, total: t.total, pct: `${pctValue}%`,
        llegada: `${pctLlegada}%`, llegadaClass: pctLlegada >= 80 ? 'pill-success' : 'pill-danger',
        cargue: `${pctCargue}%`, cargueClass: pctCargue >= 80 ? 'pill-success' : 'pill-danger',
        color: colores[idx++ % colores.length]
      };
    });

    this.makandState.kpiLider = liderTransp;
    this.makandState.kpiLiderSub = `Lidera con ${maxViajes} despachos`;
    this.actualizarGraficoMakand(almacenes);
  }

  /**
   * PROCESAMIENTO 02: TIENDAS
   */
  private procesarReporteTiendas(rows: any[]) {
    // Normalización de cabeceras
    const datosLimpios = rows.reduce((acc: any[], row: any) => {
      const cleanRow: any = {};
      Object.keys(row).forEach(key => { if (key) cleanRow[key.trim().toUpperCase()] = row[key]; });
      
      const region = cleanRow['REGIÓN'] || cleanRow['REGION'] || cleanRow['RUTA'];
      if (region) {
        cleanRow['_regionNormalizada'] = region;
        acc.push(cleanRow);
      }
      return acc;
    }, []);

    if (datosLimpios.length === 0) return;

    this.tiendasState.status = 'Cargado';
    this.tiendasState.statusClass = 'text-success fw-bold';
    this.tiendasState.note = 'Análisis de ventanas de descarga con base en el tiempo en región capturado.';

    let totalRegistros = 0; let cumple = 0; let parcial = 0; let noCumple = 0;
    const cedis: Record<string, number> = {};
    const rutasMap: Record<string, { total: number; cumple: number }> = {};
    let semanaNum = '';

    datosLimpios.forEach((cleanRow: any) => {
      totalRegistros++;
      const region = cleanRow['_regionNormalizada'];
      if (cleanRow['SEMANA']) semanaNum = `Semana ${cleanRow['SEMANA']}`;

      const cumplimiento = String(cleanRow['CUMPLIMIENTO'] || '').toUpperCase().trim();
      if (cumplimiento.includes('NO CUMPLE')) noCumple++;
      else if (cumplimiento.includes('PARCIAL')) { parcial++; cumple++; }
      else cumple++;

      const cedi = cleanRow['CEDI'] || 'Por clasificar';
      cedis[cedi] = (cedis[cedi] || 0) + 1;

      if (!rutasMap[region]) rutasMap[region] = { total: 0, cumple: 0 };
      rutasMap[region].total++;
      if (!cumplimiento.includes('NO CUMPLE')) rutasMap[region].cumple++;
    });

    this.tiendasState.kpiTotal = totalRegistros;
    this.tiendasState.kpiPeriodo = semanaNum || 'Mes operativo';
    this.tiendasState.kpiCumple = cumple.toString();
    this.tiendasState.kpiParcial = parcial.toString();
    this.tiendasState.kpiNoCumple = noCumple.toString();

    // Lógica Sustractiva (Caso A aplicativo)
    if (totalRegistros > 0) {
      const porcentajeFallas = (noCumple / totalRegistros) * 100;
      const porcentajeEfectividad = (100 - porcentajeFallas).toFixed(1);
      this.tiendasState.kpiCumpleSub = `${porcentajeEfectividad}% de efectividad`;
      this.tiendasState.kpiNoCumpleSub = `${porcentajeFallas.toFixed(1)}% de fallas`;
    } else {
      this.tiendasState.kpiCumpleSub = '0% de efectividad';
      this.tiendasState.kpiNoCumpleSub = '0% de fallas';
    }

    this.tiendasState.tableCedi = Object.keys(cedis).map(name => ({
      name, total: cedis[name],
      pct: totalRegistros > 0 ? `${((cedis[name] / totalRegistros) * 100).toFixed(1)}%` : '0%'
    }));

    this.tiendasState.rutasMesLabel = semanaNum;
    this.tiendasState.tableRutas = Object.keys(rutasMap)
      .map(name => {
        const r = rutasMap[name];
        const pctCumple = r.total > 0 ? Math.round((r.cumple / r.total) * 100) : 0;
        return {
          name, total: r.total, pct: `${pctCumple}%`,
          pctClass: pctCumple >= 85 ? 'pill-success' : pctCumple >= 70 ? 'pill-warning' : 'pill-danger'
        };
      }).sort((a, b) => b.total - a.total).slice(0, 5);

    try {
      if (this.chartTiendas && this.chartTiendas.data.datasets) {
        this.chartTiendas.data.datasets[0].data = [cumple - parcial, parcial, noCumple];
        this.chartTiendas.update();
      }
    } catch (e) { }

    this.cdr.detectChanges();
  }

  /**
   * PROCESAMIENTO 03: AGRICULTORES
   */
  private procesarReporteAgricultores(rows: any[]) {
    // Normalización de cabeceras estricta para evitar undefined por espacios ("AGRICULTOR ")
    const datosLimpios = rows.reduce((acc: any[], row: any) => {
      const cleanRow: any = {};
      Object.keys(row).forEach(key => { if (key) cleanRow[key.trim().toUpperCase()] = row[key]; });
      
      const agri = cleanRow['AGRICULTOR'] || cleanRow['NOMBRE AGRICULTOR'];
      if (agri) {
        cleanRow['_agriNormalizado'] = agri;
        acc.push(cleanRow);
      }
      return acc;
    }, []);

    if (datosLimpios.length === 0) {
      this.agriState.status = 'Error de lectura';
      this.agriState.statusClass = 'text-danger fw-bold';
      this.agriState.note = 'No se encontró la columna AGRICULTOR en la pestaña. Verifica el formato.';
      this.cdr.detectChanges();
      return;
    }

    this.agriState.status = 'Cargado';
    this.agriState.statusClass = 'text-success fw-bold';
    this.agriState.note = 'Indicadores ponderados de recolección en fincas y tiempos de traslado a central.';

    let totalViajes = 0; let cumpleLlegada = 0; let cumpleTiempo = 0; let cumplePlanta = 0;
    let semanaStr = 'S/N';
    const agricultores: Record<string, { total: number; llegada: number; tiempo: number; planta: number }> = {};

    datosLimpios.forEach((cleanRow: any) => {
      const agri = cleanRow['_agriNormalizado'];
      totalViajes++;
      if (cleanRow['SEMANA']) semanaStr = `S${cleanRow['SEMANA']}`;

      if (!agricultores[agri]) agricultores[agri] = { total: 0, llegada: 0, tiempo: 0, planta: 0 };
      agricultores[agri].total++;

      const cLlegada = String(cleanRow['CUMPLIMIENTO LLEGADA AGRICULTOR'] || cleanRow['CUMPLIMIENTO LLEGADA ACRICULTOR'] || '').toUpperCase();
      if (cLlegada.includes('CUMPLE') && !cLlegada.includes('NO')) { cumpleLlegada++; agricultores[agri].llegada++; }

      const cTiempo = String(cleanRow['CUMPLIMIENTO TIEMPO EN AGRICULTOR'] || cleanRow['CUMPLIMIENTO TIEMPO CARGUE'] || '').toUpperCase();
      if (cTiempo.includes('CUMPLE') && !cTiempo.includes('NO')) { cumpleTiempo++; agricultores[agri].tiempo++; }

      const cPlanta = String(cleanRow['CUMPLIMIENTO LLEGADA A PLANTA'] || cleanRow['CUMPLIMIENTO LLEGADA PLANTA'] || '').toUpperCase();
      if (cPlanta.includes('CUMPLE') && !cPlanta.includes('NO')) { cumplePlanta++; agricultores[agri].planta++; }
    });

    this.agriState.kpiSemanaLabel = `Semana: ${semanaStr}`;
    this.agriState.kpiTotal = totalViajes;
    
    // Asignación Numérica Pura (Esto es lo que evita el NaN en el HTML al usar el símbolo % manual)
    this.agriState.kpiLlegada = totalViajes > 0 ? Math.round((cumpleLlegada / totalViajes) * 100) : 0;
    this.agriState.kpiTiempo = totalViajes > 0 ? Math.round((cumpleTiempo / totalViajes) * 100) : 0;
    this.agriState.kpiPlanta = totalViajes > 0 ? Math.round((cumplePlanta / totalViajes) * 100) : 0;
    this.agriState.chartLabel = `Semana ${semanaStr}`;
    this.agriState.tableLabel = `Semana ${semanaStr}`;

    const labelsAgri: string[] = [];
    const datasetLlegada: number[] = [];

    this.agriState.table = Object.keys(agricultores).map(name => {
      const a = agricultores[name];
      const pLlegada = Math.round((a.llegada / a.total) * 100);
      const pTiempo = Math.round((a.tiempo / a.total) * 100);
      const pPlanta = Math.round((a.planta / a.total) * 100);

      labelsAgri.push(name);
      datasetLlegada.push(pLlegada);

      return {
        name, total: a.total,
        llegada: `${pLlegada}%`, llegadaClass: pLlegada >= 85 ? 'pill-success' : pLlegada >= 70 ? 'pill-warning' : 'pill-danger',
        tiempo: `${pTiempo}%`, tiempoClass: pTiempo >= 85 ? 'pill-success' : pTiempo >= 70 ? 'pill-warning' : 'pill-danger',
        planta: `${pPlanta}%`, plantaClass: pPlanta >= 85 ? 'pill-success' : pPlanta >= 70 ? 'pill-warning' : 'pill-danger'
      };
    });

    try {
      if (this.chartAgri && this.chartAgri.data.datasets.length > 0) {
        this.chartAgri.data.labels = labelsAgri;
        this.chartAgri.data.datasets[0].data = datasetLlegada;
        this.chartAgri.update();
      }
    } catch (e) { }
    
    this.cdr.detectChanges();
  }

  // --- INICIALIZADOR DE GRÁFICOS ---
  private inicializarGraficosVacios() {
    this.chartTransp = new Chart(this.chartTranspRef.nativeElement, {
      type: 'bar',
      data: { labels: [], datasets: [{ label: 'Viajes', data: [], backgroundColor: '#2a5298' }] },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false }
    });

    this.chartAlmacen = new Chart(this.chartAlmacenRef.nativeElement, {
      type: 'bar',
      data: { labels: [], datasets: [{ label: 'Despachos', data: [], backgroundColor: '#f5af19' }] },
      options: { responsive: true, maintainAspectRatio: false }
    });

    this.chartTiendas = new Chart(this.chartTiendasRef.nativeElement, {
      type: 'doughnut',
      data: { labels: ['Cumple', 'Cumple Parcial', 'No Cumple'], datasets: [{ data: [0, 0, 0], backgroundColor: ['#11998e', '#f5af19', '#ff416c'] }] },
      options: { responsive: true, maintainAspectRatio: false }
    });

    this.chartAgri = new Chart(this.chartAgriRef.nativeElement, {
      type: 'line',
      data: { labels: [], datasets: [{ label: '% Cumplimiento Llegada', data: [], borderColor: '#11998e', tension: 0.2, fill: true, backgroundColor: 'rgba(17, 153, 142, 0.1)' }] },
      options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 0, max: 100 } } }
    });
  }

  private actualizarGraficoMakand(almacenes: Record<string, number>) {
    if (!this.chartTransp || !this.chartAlmacen) return;
    this.chartTransp.data.labels = this.makandState.table.map(t => t.name);
    this.chartTransp.data.datasets[0].data = this.makandState.table.map(t => t.total);
    this.chartTransp.data.datasets[0].backgroundColor = this.makandState.table.map(t => t.color);
    this.chartTransp.update();
    this.chartAlmacen.data.labels = Object.keys(almacenes);
    this.chartAlmacen.data.datasets[0].data = Object.values(almacenes);
    this.chartAlmacen.update();
  }
}

export const dashSem = 'dashboardsemanal';