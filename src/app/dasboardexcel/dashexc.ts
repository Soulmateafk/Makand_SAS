import { Component, AfterViewInit, ViewChild, ElementRef } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Chart, registerables } from 'chart.js';
import * as XLSX from 'xlsx';

Chart.register(...registerables);

@Component({
  selector: 'app-dashexc',
  templateUrl: './dashexc.html',
  styleUrls: ['./dashexc.css']
})
export class DashexcComponent implements AfterViewInit {
  @ViewChild('chartTranspCanvas') chartTranspCanvas!: ElementRef<HTMLCanvasElement>;
  
  chartTransp: any;
  nombreArchivo: string = 'Esperando archivo...';
  viajesTotales: number | string = '--';

  constructor(private http: HttpClient) {}

  ngAfterViewInit(): void {
    this.inicializarGrafica();
  }

  inicializarGrafica(): void {
    this.chartTransp = new Chart(this.chartTranspCanvas.nativeElement, {
      type: 'doughnut',
      data: { 
        labels: ['Cargando...'], 
        datasets: [{ data: [1], backgroundColor: ['#2a3340'] }] 
      }
    });
  }

  // Esta función procesa el Excel, lo envía al backend y actualiza la vista
  enviarArchivos(event: any): void {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append('excelFiles', files[i]);
    }

    // 1. Envías al Backend
    this.http.post('http://localhost:3000/upload', formData).subscribe({
      next: (res: any) => {
        console.log('Backend procesó:', res);
        // 2. Aquí llamamos a la función con los datos devueltos por el servidor
        this.actualizarGrafica(res.resumen[0].datos); 
      }
    });
  }

  // Esta es la función que te faltaba
  actualizarGrafica(datos: any[]): void {
    this.viajesTotales = datos.length;
    
    // Ejemplo de conteo básico
    const conteo = { 'MAKAND': 0, 'ARSITRANS': 0, 'POLAR': 0 };
    datos.forEach(f => {
      const t = String(f['Transportadora'] || '').toUpperCase();
      if (t.includes('MAKAND')) conteo['MAKAND']++;
      else if (t.includes('ARSI')) conteo['ARSITRANS']++;
      else if (t.includes('POLAR')) conteo['POLAR']++;
    });

    // Actualizar Chart.js
    this.chartTransp.data.labels = Object.keys(conteo);
    this.chartTransp.data.datasets[0].data = Object.values(conteo);
    this.chartTransp.update();
  }
}