import { Component, ElementRef, ViewChild, AfterViewInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

// Interfaces internas para mantener el tipado estricto en las tablas
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
  // Referencias nativas a los canvas del HTML
  @ViewChild('chartTransp') chartTranspRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAlmacen') chartAlmacenRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartTiendas') chartTiendasRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('chartAgri') chartAgriRef!: ElementRef<HTMLCanvasElement>;

  // Instancias de Chart.js para su control de refresco
  private chartTransp?: Chart;
  private chartAlmacen?: Chart;
  private chartTiendas?: Chart;
  private chartAgri?: Chart;

  // --- Módulos de Estado vinculados directamente con tu Plantilla HTML ---

  makandState = {
    src: 'Ningún archivo seleccionado',
    status: 'Pendiente de carga',
    statusClass: 'badge-pendiente',
    note: 'Sube el reporte de "VIAJEROS Y TERCEROS" para calcular viajes e ingresos a planta.',
    kpiTotal: 0,
    kpiPeriodo: 'Sin datos',
    kpiCumple: '0%',
    kpiCajas: 0,
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
    kpiCumpleSub: '0% del total',
    kpiParcial: '0',
    kpiParcialSub: '0% del total',
    kpiNoCumple: '0',
    kpiNoCumpleSub: '0% del total',
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
    kpiLlegada: '0%',
    kpiTiempo: '0%',
    kpiPlanta: '0%',
    chartLabel: '',
    tableLabel: 'Pendiente',
    table: [] as AgriTableRow[]
  };

  ngAfterViewInit() {
    this.inicializarGraficosVacios();
  }

  /**
   * Gestor de eventos unificado para la carga de los tres archivos Excel
   */
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

      // Buscador dinámico de pestañas operativas comunes en tus archivos
      const sheetName = workbook.SheetNames.find(name => {
        const n = name.toLowerCase();
        return n.includes('dt') || n.includes('detalle') || n.includes('tiempo') || n === 'hoja1';
      }) || workbook.SheetNames[0];

      const worksheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json<any>(worksheet, { defval: '' });

      if (jsonRows.length > 0) {
        if (tipo === 'makand') this.procesarReporteMakand(jsonRows);
        if (tipo === 'tiendas') this.procesarReporteTiendas(jsonRows);
        if (tipo === 'agri') this.procesarReporteAgricultores(jsonRows);
      }
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * PROCESAMIENTO REPORTE 01: VIAJEROS Y TERCEROS (PLANTA MAKAND)
   */
  private procesarReporteMakand(rows: any[]) {
    this.makandState.status = 'Cargado';
    this.makandState.statusClass = 'text-success fw-bold';
    this.makandState.note = 'Datos consolidados en base al histórico de transportadoras externas y flota propia.';

    let totalViajes = 0;
    let totalCajas = 0;
    let viajesCumpleLlegada = 0;
    
    const transportadoras: Record<string, { total: number; cumpleLlegada: number; cumpleCargue: number }> = {};
    const almacenes: Record<string, number> = {};
    let mesDetectado = 'Julio 2026';

    rows.forEach(row => {
      const transp = row['TRANSPORTE'] || row['Transportadora'];
      if (!transp) return; // Salta filas vacías o de decoración en el Excel

      totalViajes++;
      if (row['MES']) mesDetectado = row['MES'];

      // Conteo de cajas movilizadas
      const cajas = parseFloat(row['TOTAL CAJAS'] || row['Total Cajas'] || 0);
      totalCajas += isNaN(cajas) ? 0 : cajas;

      // Agrupación por Transportadora
      if (!transportadoras[transp]) {
        transportadoras[transp] = { total: 0, cumpleLlegada: 0, cumpleCargue: 0 };
      }
      transportadoras[transp].total++;

      // Evaluar cumplimientos de llegada y cargue
      const cumpleLlegadaStr = String(row['CUMPLIMIENTO LLEGADA VEHICULO'] || '').toUpperCase();
      if (cumpleLlegadaStr.includes('CUMPLE') && !cumpleLlegadaStr.includes('NO')) {
        transportadoras[transp].cumpleLlegada++;
        viajesCumpleLlegada++;
      }

      const cumpleCargueStr = String(row['CUMPLIMIENTO CARGUE VEHICULO'] || '').toUpperCase();
      if (cumpleCargueStr.includes('CUMPLE') && !cumpleCargueStr.includes('NO')) {
        transportadoras[transp].cumpleCargue++;
      }

      // Agrupación por almacén destino
      const almacen = row['ALMACEN'] || row['Almacen'] || 'Otros';
      almacenes[almacen] = (almacenes[almacen] || 0) + 1;
    });

    // Rellenado de KPIs de la sección 01
    this.makandState.kpiTotal = totalViajes;
    this.makandState.kpiPeriodo = `Periodo operativo: ${mesDetectado}`;
    this.makandState.kpiCumple = totalViajes > 0 ? `${Math.round((viajesCumpleLlegada / totalViajes) * 100)}%` : '0%';
    this.makandState.kpiCajas = totalCajas.toLocaleString('de-DE') as any;
    this.makandState.kpiCajasSub = 'unidades distribuidas';

    // Generar la tabla de transportadoras para el HTML
    const colores = ['#2a5298', '#11998e', '#ff416c', '#f5af19', '#8e44ad'];
    let idx = 0;
    let maxViajes = -1;
    let liderTransp = 'Ninguna';

    this.makandState.table = Object.keys(transportadoras).map(name => {
      const t = transportadoras[name];
      const pctValue = ((t.total / totalViajes) * 100).toFixed(1);
      
      if (t.total > maxViajes) {
        maxViajes = t.total;
        liderTransp = name;
      }

      const pctLlegada = Math.round((t.cumpleLlegada / t.total) * 100);
      const pctCargue = Math.round((t.cumpleCargue / t.total) * 100);

      return {
        name,
        total: t.total,
        pct: `${pctValue}%`,
        llegada: `${pctLlegada}%`,
        llegadaClass: pctLlegada >= 80 ? 'pill-success' : 'pill-danger',
        cargue: `${pctCargue}%`,
        cargueClass: pctCargue >= 80 ? 'pill-success' : 'pill-danger',
        color: colores[idx++ % colores.length]
      };
    });

    this.makandState.kpiLider = liderTransp;
    this.makandState.kpiLiderSub = `Lidera con ${maxViajes} despachos`;

    // Renderizado de gráficos de la sección 01
    this.actualizarGraficoMakand(almacenes);
  }

  /**
   * PROCESAMIENTO REPORTE 02: HORARIOS EN TIENDAS (D1, ARA, EXITO)
   */
  private procesarReporteTiendas(rows: any[]) {
    this.tiendasState.status = 'Cargado';
    this.tiendasState.statusClass = 'text-success fw-bold';
    this.tiendasState.note = 'Análisis de ventanas de descarga con base en el tiempo en región capturado.';

    let totalRegistros = 0;
    let cumple = 0;
    let parcial = 0;
    let noCumple = 0;

    const cedis: Record<string, number> = {};
    const rutasMap: Record<string, { total: number; cumple: number }> = {};
    let semanaNum = '';

    rows.forEach(row => {
      const region = row['Región'] || row['Región'] || row['RUTA '];
      if (!region) return;

      totalRegistros++;
      if (row['SEMANA']) semanaNum = `Semana ${row['SEMANA']}`;

      // Clasificación de cumplimiento de horarios
      const cumplimiento = String(row['CUMPLIMIENTO'] || '').toUpperCase();
      if (cumplimiento.includes('NO CUMPLE')) {
        noCumple++;
      } else if (cumplimiento.includes('PARCIAL')) {
        parcial++;
        cumple++; // Si aplica como cumplimiento intermedio
      } else {
        cumple++;
      }

      // Conteo por CEDI de despacho
      const cedi = row['CEDI'] || 'Por clasificar';
      cedis[cedi] = (cedis[cedi] || 0) + 1;

      // Desempeño de rutas críticas
      if (!rutasMap[region]) rutasMap[region] = { total: 0, cumple: 0 };
      rutasMap[region].total++;
      if (!cumplimiento.includes('NO CUMPLE')) rutasMap[region].cumple++;
    });

    this.tiendasState.kpiTotal = totalRegistros;
    this.tiendasState.kpiPeriodo = semanaNum || 'Mes operativo';
    this.tiendasState.kpiCumple = cumple.toString();
    this.tiendasState.kpiCumpleSub = `${((cumple / totalRegistros) * 100).toFixed(1)}% de efectividad`;
    this.tiendasState.kpiParcial = parcial.toString();
    this.tiendasState.kpiParcialSub = 'Entregas toleradas';
    this.tiendasState.kpiNoCumple = noCumple.toString();
    this.tiendasState.kpiNoCumpleSub = `${((noCumple / totalRegistros) * 100).toFixed(1)}% de fallas`;

    // Armar tabla CEDI
    this.tiendasState.tableCedi = Object.keys(cedis).map(name => ({
      name,
      total: cedis[name],
      pct: `${((cedis[name] / totalRegistros) * 100).toFixed(1)}%`
    }));

    // Armar tabla Rutas principales
    this.tiendasState.rutasMesLabel = semanaNum;
    this.tiendasState.tableRutas = Object.keys(rutasMap)
      .map(name => {
        const r = rutasMap[name];
        const pctCumple = Math.round((r.cumple / r.total) * 100);
        return {
          name,
          total: r.total,
          pct: `${pctCumple}%`,
          pctClass: pctCumple >= 85 ? 'pill-success' : pctCumple >= 70 ? 'pill-warning' : 'pill-danger'
        };
      })
      .sort((a, b) => b.total - a.total)
      .slice(0, 5); // Tomamos el Top 5 rutas con más movimientos

    // Refrescar gráfico circular de tiendas
    if (this.chartTiendas) {
      this.chartTiendas.data.datasets[0].data = [cumple - parcial, parcial, noCumple];
      this.chartTiendas.update();
    }
  }

  /**
   * PROCESAMIENTO REPORTE 03: RECOLECCIÓN AGRICULTORES (CAMPO Y TRANSPORTE)
   */
  private procesarReporteAgricultores(rows: any[]) {
    this.agriState.status = 'Cargado';
    this.agriState.statusClass = 'text-success fw-bold';
    this.agriState.note = 'Indicadores ponderados de recolección en fincas y tiempos de traslado a central.';

    let totalViajes = 0;
    let cumpleLlegada = 0;
    let cumpleTiempo = 0;
    let cumplePlanta = 0;
    let semanaStr = 'S/N';

    const agricultores: Record<string, { total: number; llegada: number; tiempo: number; planta: number }> = {};

    rows.forEach(row => {
      const agri = row['AGRICULTOR'] || row['Agricultor'];
      if (!agri) return;

      totalViajes++;
      if (row['semana']) semanaStr = `S${row['semana']}`;

      if (!agricultores[agri]) {
        agricultores[agri] = { total: 0, llegada: 0, tiempo: 0, planta: 0 };
      }
      agricultores[agri].total++;

      // 1. Cumplimiento de llegada del carro al agricultor
      const cLlegada = String(row['CUMPLIMIENTO LLEGADA AGRICULTOR'] || row['CUMPLIMIENTO LLEGADA ACRICULTOR'] || '').toUpperCase();
      if (cLlegada.includes('CUMPLE') && !cLlegada.includes('NO')) {
        cumpleLlegada++;
        agricultores[agri].llegada++;
      }

      // 2. Cumplimiento del tiempo de permanencia en la finca
      const cTiempo = String(row['CUMPLIMIENTO TIEMPO EN AGRICULTOR'] || '').toUpperCase();
      if (cTiempo.includes('CUMPLE') && !cTiempo.includes('NO')) {
        cumpleTiempo++;
        agricultores[agri].tiempo++;
      }

      // 3. Cumplimiento del transporte hacia la planta
      const cPlanta = String(row['CUMPLIMIENTO LLEGADA A PLANTA'] || '').toUpperCase();
      if (cPlanta.includes('CUMPLE') && !cPlanta.includes('NO')) {
        cumplePlanta++;
        agricultores[agri].planta++;
      }
    });

    this.agriState.kpiSemanaLabel = `Semana: ${semanaStr}`;
    this.agriState.kpiTotal = totalViajes;
    this.agriState.kpiLlegada = totalViajes > 0 ? `${Math.round((cumpleLlegada / totalViajes) * 100)}%` : '0%';
    this.agriState.kpiTiempo = totalViajes > 0 ? `${Math.round((cumpleTiempo / totalViajes) * 100)}%` : '0%';
    this.agriState.kpiPlanta = totalViajes > 0 ? `${Math.round((cumplePlanta / totalViajes) * 100)}%` : '0%';
    
    this.agriState.chartLabel = `Semana ${semanaStr}`;
    this.agriState.tableLabel = `Semana ${semanaStr}`;

    // Construcción de la tabla estructurada para agricultores
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
        name,
        total: a.total,
        llegada: `${pLlegada}%`,
        llegadaClass: pLlegada >= 85 ? 'pill-success' : pLlegada >= 70 ? 'pill-warning' : 'pill-danger',
        tiempo: `${pTiempo}%`,
        tiempoClass: pTiempo >= 85 ? 'pill-success' : pTiempo >= 70 ? 'pill-warning' : 'pill-danger',
        planta: `${pPlanta}%`,
        plantaClass: pPlanta >= 85 ? 'pill-success' : pPlanta >= 70 ? 'pill-warning' : 'pill-danger'
      };
    });

    // Actualizar gráfico de barras de agricultores
    if (this.chartAgri) {
      this.chartAgri.data.labels = labelsAgri;
      this.chartAgri.data.datasets[0].data = datasetLlegada;
      this.chartAgri.update();
    }
  }

  /**
   * INICIALIZADOR DE GRÁFICOS: Inicializa las vistas de Chart.js con configuraciones limpias
   */
  private inicializarGraficosVacios() {
    // 01. Gráfico Barras Horizontal - Viajes por Transportadora
    this.chartTransp = new Chart(this.chartTranspRef.nativeElement, {
      type: 'bar',
      data: { labels: [], datasets: [{ label: 'Viajes', data: [], backgroundColor: '#2a5298' }] },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false }
    });

    // 01. Gráfico de Barras Vertical - Destino de Almacenes
    this.chartAlmacen = new Chart(this.chartAlmacenRef.nativeElement, {
      type: 'bar',
      data: { labels: [], datasets: [{ label: 'Despachos', data: [], backgroundColor: '#f5af19' }] },
      options: { responsive: true, maintainAspectRatio: false }
    });

    // 02. Gráfico de Dona - Distribución de cumplimiento Tiendas
    this.chartTiendas = new Chart(this.chartTiendasRef.nativeElement, {
      type: 'doughnut',
      data: {
        labels: ['Cumple', 'Cumple Parcial', 'No Cumple'],
        datasets: [{ data: [0, 0, 0], backgroundColor: ['#11998e', '#f5af19', '#ff416c'] }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });

    // 03. Gráfico de Línea - Cumplimiento Llegada por Agricultor
    this.chartAgri = new Chart(this.chartAgriRef.nativeElement, {
      type: 'line',
      data: { labels: [], datasets: [{ label: '% Cumplimiento Llegada', data: [], borderColor: '#11998e', tension: 0.2, fill: true, backgroundColor: 'rgba(17, 153, 142, 0.1)' }] },
      options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 0, max: 100 } } }
    });
  }

  /**
   * Auxiliar para refrescar y poblar los gráficos cruzados del módulo Makand
   */
  private actualizarGraficoMakand(almacenes: Record<string, number>) {
    if (!this.chartTransp || !this.chartAlmacen) return;

    // Actualizar Gráfico 1: Viajes por transportadora
    this.chartTransp.data.labels = this.makandState.table.map(t => t.name);
    this.chartTransp.data.datasets[0].data = this.makandState.table.map(t => t.total);
    this.chartTransp.data.datasets[0].backgroundColor = this.makandState.table.map(t => t.color);
    this.chartTransp.update();

    // Actualizar Gráfico 2: Despachos por almacén de destino
    this.chartAlmacen.data.labels = Object.keys(almacenes);
    this.chartAlmacen.data.datasets[0].data = Object.values(almacenes);
    this.chartAlmacen.update();
  }
}

// --- SOLUCIÓN ERROR COMPILADOR ROUTING ANGULAR ---
// Token exportado requerido por tu 'app.routes.ts' para mapear la ruta sin fallas de carga limpia.
export const dashSem = 'dashboardsemanal';