import { Component, AfterViewInit, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common'; // ¡Crucial para solucionar el error de ngClass!
import { Chart, registerables } from 'chart.js';
import * as XLSX from 'xlsx';

Chart.register(...registerables);

@Component({
  selector: 'app-dashexc',
  standalone: true, // Si usas la estructura moderna de Angular
  imports: [CommonModule], // Esto activa ngClass, ngIf y ngFor en tu HTML
  templateUrl: './dashexc.html',
  styleUrls: ['./dashexc.css']
})
export class DashexcComponent implements AfterViewInit {
  // Referencias a los Canvas del HTML para las gráficas
  @ViewChild('chartMakandDonut') chartMakandDonutRef!: ElementRef;
  @ViewChild('chartMakandBar') chartMakandBarRef!: ElementRef;
  @ViewChild('chartTiendasDonut') chartTiendasDonutRef!: ElementRef;
  @ViewChild('chartAgriBar') chartAgriBarRef!: ElementRef;

  // Instancias de Chart.js
  graficaMakandDonut: any;
  graficaMakandBar: any;
  graficaTiendasDonut: any;
  graficaAgriBar: any;

  // Nombres de archivos cargados
  archivoMakand: string = 'Falta archivo';
  archivoTiendas: string = 'Falta archivo';
  archivoAgricultores: string = 'Falta archivo';

  // --- DATOS SECCIÓN 01: MAKAND (Inicializados con los valores de tu imagen) ---
  kpiMakand = { totales: 258, cumplimiento: 80.6, cajas: '115,432', lider: 'MAKAND', porcLider: 51.2 };
  tablaMakand = [
    { nombre: 'MAKAND', viajes: 132, porc: 51.2, cumpleLlegada: 93.2, cumpleCargue: 49.2 },
    { nombre: 'ARSITRANS', viajes: 78, porc: 30.2, cumpleLlegada: 71.8, cumpleCargue: 48.7 },
    { nombre: 'POLAR', viajes: 48, porc: 18.6, cumpleLlegada: 60.4, cumpleCargue: 89.6 }
  ];

  // --- DATOS SECCIÓN 02: TIENDAS ---
  kpiTiendas = { registros: 929, cumple: 33.8, parcial: 12.6, noCumple: 53.6 };
  tablaCedis = [
    { nombre: 'Tiendas', registros: 602, part: 64.8 },
    { nombre: 'CEDI 2', registros: 200, part: 21.5 },
    { nombre: 'CEDI ARA', registros: 74, part: 8.0 },
    { nombre: 'CEDI 1', registros: 53, part: 5.7 }
  ];
  tablaRutas = [
    { nombre: 'Plataforma Siberia', registros: 63, cumple: '--' },
    { nombre: 'Ara Cota', registros: 45, cumple: '26.7%' },
    { nombre: 'Olímpica', registros: 37, cumple: '--' },
    { nombre: 'Plataforma Cencosud', registros: 34, cumple: '--' },
    { nombre: 'D1 Tocancipá', registros: 30, cumple: '--' }
  ];

  // --- DATOS SECCIÓN 03: AGRICULTORES ---
  kpiAgri = { viajes: 70, llegada: 54.3, finca: 21.4, planta: 0.0 };
  tablaAgri = [
    { nombre: 'Diego', viajes: 6, llegada: '16.7%', finca: '16.7%', planta: '0.0%' },
    { nombre: 'Ferrucas', viajes: 47, llegada: '65.9%', finca: '0.0%', planta: '0.0%' },
    { nombre: 'Gabriel', viajes: 7, llegada: '0.0%', finca: '85.7%', planta: '0.0%' },
    { nombre: 'Georgeth', viajes: 1, llegada: '0.0%', finca: '100.0%', planta: '0.0%' },
    { nombre: 'José Tibaquicha', viajes: 3, llegada: '0.0%', finca: '33.3%', planta: '0.0%' },
    { nombre: 'Lechugas del Día', viajes: 6, llegada: '100.0%', finca: '100.0%', planta: '-' }
  ];

  ngAfterViewInit(): void {
    this.inicializarGraficas();
  }

  inicializarGraficas(): void {
    // Configuración común para textos claros en modo oscuro
    const opcionesPlugins = {
      legend: { position: 'bottom' as const, labels: { color: '#8b96a5', font: { size: 11 } } }
    };

    // 1. Dona Makand
    this.graficaMakandDonut = new Chart(this.chartMakandDonutRef.nativeElement, {
      type: 'doughnut',
      data: {
        labels: ['MAKAND', 'ARSITRANS', 'POLAR'],
        datasets: [{ data: [132, 78, 48], backgroundColor: ['#2dd4bf', '#f2a93c', '#e2564f'], borderWidth: 0 }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: opcionesPlugins }
    });

    // 2. Barras Almacén Destino Makand
    this.graficaMakandBar = new Chart(this.chartMakandBarRef.nativeElement, {
      type: 'bar',
      data: {
        labels: ['D1', 'ARA', 'ÉXITO'],
        datasets: [
          { label: 'MAKAND', data: [100, 10, 20], backgroundColor: '#2dd4bf' },
          { label: 'ARSITRANS', data: [50, 20, 10], backgroundColor: '#f2a93c' },
          { label: 'POLAR', data: [28, 5, 15], backgroundColor: '#e2564f' }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { color: '#8b96a5' } },
          y: { stacked: true, grid: { color: '#2a3340' }, ticks: { color: '#8b96a5' } }
        },
        plugins: opcionesPlugins
      }
    });

    // 3. Dona Tiendas
    this.graficaTiendasDonut = new Chart(this.chartTiendasDonutRef.nativeElement, {
      type: 'doughnut',
      data: {
        labels: ['Cumple', 'Cumple parcial', 'No cumple'],
        datasets: [{ data: [314, 117, 498], backgroundColor: ['#2dd4bf', '#f2a93c', '#e2564f'], borderWidth: 0 }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: opcionesPlugins }
    });

    // 4. Barras Horizontales Agricultores
    this.graficaAgriBar = new Chart(this.chartAgriBarRef.nativeElement, {
      type: 'bar',
      data: {
        labels: ['Diego', 'Ferrucas', 'Gabriel', 'Georgeth', 'José T.', 'Lechugas'],
        datasets: [{ data: [16.7, 65.9, 0, 0, 0, 100], backgroundColor: '#2dd4bf', barThickness: 14 }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { max: 100, grid: { color: '#2a3340' }, ticks: { color: '#8b96a5' } },
          y: { grid: { display: false }, ticks: { color: '#8b96a5' } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }

  // --- LECTURA REAL EN TIEMPO REAL DE LOS 3 EXCEL ---
  enviarArchivos(event: any): void {
    const files: File[] = Array.from(event.target.files);
    if (files.length === 0) return;

    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
        
        const nombre = file.name.toLowerCase();
        
        if (nombre.includes('makand')) {
          this.archivoMakand = file.name;
          this.procesarExcelMakand(jsonData);
        } else if (nombre.includes('tienda')) {
          this.archivoTiendas = file.name;
          this.procesarExcelTiendas(jsonData);
        } else if (nombre.includes('agri')) {
          this.archivoAgricultores = file.name;
          this.procesarExcelAgri(jsonData);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  procesarExcelMakand(datos: any[]): void {
    if (datos.length === 0) return;
    this.kpiMakand.totales = datos.length;
    // Aquí puedes realizar conteos reales dinámicos basados en las columnas de tu Excel
    // Ejemplo: row['Transportadora'], row['Cumplimiento'], etc.
    this.graficaMakandDonut.update();
  }

  procesarExcelTiendas(datos: any[]): void {
    if (datos.length === 0) return;
    this.kpiTiendas.registros = datos.length;
    this.graficaTiendasDonut.update();
  }

  procesarExcelAgri(datos: any[]): void {
    if (datos.length === 0) return;
    this.kpiAgri.viajes = datos.length;
    this.graficaAgriBar.update();
  }
}