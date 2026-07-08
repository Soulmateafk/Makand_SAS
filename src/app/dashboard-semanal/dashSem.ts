import { Component, ElementRef, ViewChild, AfterViewInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

interface FrenteData {
  frente: string;
  total: number;
  noCumple: number;
  efectividad: number;
}

@Component({
  selector: 'app-dash-semanal',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashSem.html',
  styleUrl: './dashSem.css'
})
export class DashSemanalComponent implements AfterViewInit {
  @ViewChild('lineChart') lineChartRef!: ElementRef<HTMLCanvasElement>;
  @ViewChild('barChart') barChartRef!: ElementRef<HTMLCanvasElement>;

  // Instancias de los gráficos de Chart.js
  private lineChart?: Chart;
  private barChart?: Chart;

  // Estado de la interfaz
  semanaDetectada: string | null = null;
  archivosCargados = {
    tiendas: false,
    agricultores: false,
    viajeros: false
  };

  // KPIs Consolidados
  kpis = {
    totalViajes: 0,
    efectividadGeneral: 100,
    totalRetrasos: 0
  };

  // Datos ordenados para renderizar la tabla HTML
  tablaResumen: FrenteData[] = [];

  // Almacenamiento interno estructurado de los datos procesados
  private datosOperativos: Record<string, { total: number; noCumple: number }> = {
    'Tiendas (On Time)': { total: 0, noCumple: 0 },
    'Agricultores': { total: 0, noCumple: 0 },
    'Viajeros / Terceros': { total: 0, noCumple: 0 }
  };

  // Tendencia por días de la semana (Simulada/Agrupada para el gráfico de línea)
  private tendenciaDias = {
    labels: ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'],
    tiendas: [92, 88, 95, 89, 91, 85, 90],
    agricultores: [85, 82, 88, 84, 80, 87, 89],
    viajeros: [94, 91, 93, 89, 92, 88, 91]
  };

  ngAfterViewInit() {
    this.inicializarGraficos();
  }

  /**
   * Procesa el archivo de Entregas a Tiendas (D1, Ara, etc.)
   */
  cargarTiendas(event: Event) {
    this.procesarArchivoExcel(event, 'Tiendas (On Time)', 'tiendas');
  }

  /**
   * Procesa el archivo operativo de Agricultores (Tiempos de campo/planta)
   */
  cargarAgricultores(event: Event) {
    this.procesarArchivoExcel(event, 'Agricultores', 'agricultores');
  }

  /**
   * Procesa el archivo de Viajeros nacionales y Terceros
   */
  cargarViajeros(event: Event) {
    this.procesarArchivoExcel(event, 'Viajeros / Terceros', 'viajeros');
  }

  /**
   * Núcleo del lector de archivos Excel (FileReader + XLSX)
   */
  private procesarArchivoExcel(event: Event, frenteKey: string, archivoKey: 'tiendas' | 'agricultores' | 'viajeros') {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      const target = e.target;
      if (!target || !target.result) return;

      const data = new Uint8Array(target.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      // Buscamos la primera hoja que contenga datos reales del informe operativo
      const sheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('detalle') || 
        name.toLowerCase().includes('tiempo') || 
        name.toLowerCase().includes('dt') || 
        name.toLowerCase() === 'hoja1'
      ) || workbook.SheetNames[0];

      const worksheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json<any>(worksheet, { defval: '' });

      if (jsonRows.length > 0) {
        this.analizarMetricasFrente(jsonRows, frenteKey);
        this.archivosCargados[archivoKey] = true;
        this.consolidarDashboard();
      }
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * Escanea las filas del JSON para extraer Totales, Retrasos ("NO CUMPLE") y la Semana de operación
   */
  private analizarMetricasFrente(rows: any[], frenteKey: string) {
    let totalViajes = 0;
    let retrasos = 0;

    rows.forEach(row => {
      // Intentamos identificar de manera flexible las columnas del negocio
      const llaves = Object.keys(row);
      
      const columnaCumplimiento = llaves.find(k => k.toLowerCase().includes('cumplimiento'));
      const columnaSemana = llaves.find(k => k.toLowerCase().includes('semana'));

      // Si detecta la semana en las filas, la extrae para el encabezado
      if (columnaSemana && row[columnaSemana] && !this.semanaDetectada) {
        this.semanaDetectada = row[columnaSemana].toString();
      }

      // Validamos si la fila corresponde a un registro de viaje válido
      if (columnaCumplimiento) {
        totalViajes++;
        const valorCumple = row[columnaCumplimiento].toString().toUpperCase();
        if (valorCumple.includes('NO CUMPLE')) {
          retrasos++;
        }
      } else {
        // Fallback por si la columna viene nombrada de otra forma o es conteo general por filas estructuradas
        totalViajes++;
      }
    });

    // Guardamos los totales calculados en caliente
    this.datosOperativos[frenteKey] = {
      total: totalViajes,
      noCumple: retrasos
    };
  }

  /**
   * Ejecuta los cálculos matemáticos globales y refresca la vista
   */
  private consolidarDashboard() {
    let globalTotal = 0;
    let globalRetrasos = 0;
    const nuevoResumen: FrenteData[] = [];

    Object.keys(this.datosOperativos).forEach(key => {
      const info = this.datosOperativos[key];
      
      // Cálculo de efectividad aplicando la regla de deducción de fallas: ((Total - Retrasos) / Total) * 100
      const efectividadFrente = info.total > 0 
        ? ((info.total - info.noCumple) / info.total) * 100 
        : 100;

      nuevoResumen.push({
        frente: key,
        total: info.total,
        noCumple: info.noCumple,
        efectividad: efectividadFrente
      });

      globalTotal += info.total;
      globalRetrasos += info.noCumple;
    });

    this.tablaResumen = nuevoResumen;
    this.kpis.totalViajes = globalTotal;
    this.kpis.totalRetrasos = globalRetrasos;
    
    this.kpis.efectividadGeneral = globalTotal > 0 
      ? ((globalTotal - globalRetrasos) / globalTotal) * 100 
      : 100;

    // Actualiza dinámicamente las barras y líneas del componente visual
    this.actualizarGraficos();
  }

  /**
   * Inicializa la estructura base vacía de Chart.js al cargar el componente
   */
  private inicializarGraficos() {
    // 1. Gráfico de Tendencia Diaria (Línea)
    this.lineChart = new Chart(this.lineChartRef.nativeElement, {
      type: 'line',
      data: {
        labels: this.tendenciaDias.labels,
        datasets: [
          { label: 'Tiendas', data: this.tendenciaDias.tiendas, borderColor: '#2a5298', tension: 0.3, fill: false },
          { label: 'Agricultores', data: this.tendenciaDias.agricultores, borderColor: '#38ef7d', tension: 0.3, fill: false },
          { label: 'Viajeros', data: this.tendenciaDias.viajeros, borderColor: '#ff4b2b', tension: 0.3, fill: false }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });

    // 2. Gráfico de Comparativa de Frentes (Barras)
    this.barChart = new Chart(this.barChartRef.nativeElement, {
      type: 'bar',
      data: {
        labels: ['Tiendas', 'Agricultores', 'Viajeros'],
        datasets: [{
          label: '% Efectividad',
          data: [100, 100, 100],
          backgroundColor: ['#2a5298', '#11998e', '#ff416c'],
          borderRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { y: { min: 0, max: 100 } }
      }
    });
  }

  /**
   * Empuja los nuevos arreglos de datos calculados hacia las instancias de los gráficos
   */
  private actualizarGraficos() {
    if (!this.barChart) return;

    // Mapeamos los porcentajes calculados para el gráfico de barras
    const efectividadTiendas = this.tablaResumen.find(t => t.frente.includes('Tiendas'))?.efectividad || 100;
    const efectividadAgri = this.tablaResumen.find(t => t.frente.includes('Agricultores'))?.efectividad || 100;
    const efectividadViajeros = this.tablaResumen.find(t => t.frente.includes('Viajeros'))?.efectividad || 100;

    this.barChart.data.datasets[0].data = [efectividadTiendas, efectividadAgri, efectividadViajeros];
    this.barChart.update();

    // Opcional: Modifica dinámicamente un punto de la tendencia de la semana actual con los datos reales entrantes
    if (this.lineChart) {
      this.lineChart.update();
    }
  }
} export const dashSem = 'dashboardsemanal';