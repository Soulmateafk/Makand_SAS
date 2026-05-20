import { Component, ViewChild, ElementRef, AfterViewInit, OnInit, Inject, PLATFORM_ID, ChangeDetectorRef } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http'; 
import { Chart, registerables } from 'chart.js';
import { AlertasComponent } from '../pages/alertas/alertas'; 

Chart.register(...registerables);

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [
    CommonModule, 
    FormsModule,  
    AlertasComponent 
  ],
  templateUrl: './dashboard.html',
  styleUrls: ['./dashboard.css']
})
export class dashboard implements AfterViewInit, OnInit {
  @ViewChild('lineChartCanvas') lineChartCanvas!: ElementRef;
  @ViewChild('doughnutChartCanvas') doughnutChartCanvas!: ElementRef;

  public editId: any | null = null; 
  public searchTerm: string = '';
  public vehiculosFiltrados: any[] = [];
  public fallasData: any[] = []; 

  public totalVehiculos: number = 0;
  public totalConductores: number = 0;
  
  public vBueno: number = 0;
  public vRegular: number = 0;
  public vMalo: number = 0;
  public vTaller: number = 0;

  public statsMantenimiento: any = {
    total: 0,
    enTaller: 0,
    operativos: 0,
    vencidosFecha: 0,
    alertaKilometraje: 0
  };

  public errorPlaca: string = ''; 
  private baseApiUrl = 'http://localhost:3000/api'; 

  private lineChart: any;
  private doughnutChart: any;

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cdr: ChangeDetectorRef 
  ) {}

  get totalEntregas() { return this.fallasData.reduce((acc, f) => acc + (Number(f.entregas) || 0), 0); }
  get totalDevoluciones() { return this.fallasData.reduce((acc, f) => acc + (Number(f.devoluciones) || 0), 0); }
  get totalViajes() { return this.fallasData.length; }
  get tasaRechazo() { 
    if (this.totalEntregas === 0) return 0;
    return ((this.totalDevoluciones / this.totalEntregas) * 100).toFixed(1);
  }

  // Dejamos el ngOnInit vacío o preparado para el renderizado del cliente
  ngOnInit() {}

  // 🔒 Cambiamos la carga aquí para asegurar que el DOM, los Canvas y el localStorage estén listos en el navegador
  ngAfterViewInit() {
    if (isPlatformBrowser(this.platformId)) {
      this.cargarDatos();
    }
  }

  cargarDatos() {
    if (isPlatformBrowser(this.platformId)) {
      
      // 1. CARGAR CONDUCTORES
      this.http.get<any[]>(`${this.baseApiUrl}/conductores`).subscribe({
        next: (cond) => {
          this.totalConductores = cond.length;
          this.cdr.detectChanges(); 
        }
      });

      // 2. CARGAR MANTENIMIENTO
      this.http.get<any>(`${this.baseApiUrl}/mantenimiento/stats`).subscribe({
        next: (res) => {
          this.statsMantenimiento = res;
          this.vTaller = res.enTaller;
          this.cdr.detectChanges();
          this.actualizarGraficas();
        }
      });

      // 3. VEHÍCULOS Y ENTREGAS (Con filtro de seguridad y agrupación)
      this.http.get<any[]>(`${this.baseApiUrl}/vehiculos`).subscribe({
        next: (flota) => {
          this.totalVehiculos = flota.length;
          this.vBueno = flota.filter((v: any) => v.estado === 'Bueno' || v.estado === 'Óptimo').length;
          this.vRegular = flota.filter((v: any) => v.estado === 'Regular').length;
          this.vMalo = flota.filter((v: any) => v.estado === 'Malo').length;

          this.http.get<any[]>(`${this.baseApiUrl}/entregas`).subscribe({
            next: (entregas) => {
              
              // FILTRO DE SEGURIDAD: Solo procesar registros que tengan datos reales
              const entregasValidas = entregas.filter(e => 
                (Number(e.entregas) > 0 || Number(e.devoluciones) > 0)
              );

              const agrupado = entregasValidas.reduce((acc: any, curr: any) => {
                const placaKey = curr.placa?.toString().toUpperCase();
                if (!placaKey) return acc;

                if (!acc[placaKey]) {
                  acc[placaKey] = { ...curr, entregas: 0, devoluciones: 0 };
                }
                
                acc[placaKey].entregas += Number(curr.entregas) || 0;
                acc[placaKey].devoluciones += Number(curr.devoluciones) || 0;
                
                // Priorizar destinos reales sobre etiquetas por defecto
                if (curr.destino && !['Sin destino', 'Sin destino asignado'].includes(curr.destino)) {
                  acc[placaKey].destino = curr.destino;
                }
                return acc;
              }, {});

              this.fallasData = Object.values(agrupado).map((ent: any) => {
                const vInfo = flota.find(v => v.placa?.toUpperCase() === ent.placa?.toUpperCase());
                
                // --- NUEVA LÓGICA DE EFECTIVIDAD LOGÍSTICA REAL ---
                const totalRuta = Number(ent.entregas) || 0;
                const devueltas = Number(ent.devoluciones) || 0;
                let porcentajeEfectividad = 0;

                if (totalRuta > 0) {
                  const exitosas = totalRuta - devueltas;
                  porcentajeEfectividad = (exitosas / totalRuta) * 100;

                  // Control de seguridad por si las devoluciones superan las entregas cargadas
                  if (porcentajeEfectividad < 0) {
                    porcentajeEfectividad = 0;
                  }
                }

                return {
                  ...ent,
                  kilometraje: vInfo ? vInfo.kilometraje : 0, 
                  kmActual: vInfo ? vInfo.km : 0,
                  kilometrajeProximo: vInfo ? vInfo.kilometrajeProximo : 0,
                  efectividad: Number(porcentajeEfectividad.toFixed(1))
                };
              });

              this.vehiculosFiltrados = [...this.fallasData];
              this.cdr.detectChanges();
              this.actualizarGraficas();
            }
          });
        }
      });
    }
  }

  private actualizarGraficas() {
    if (isPlatformBrowser(this.platformId) && this.lineChartCanvas && this.doughnutChartCanvas) {
      this.renderCharts();
    }
  }

  filtrarVehiculos() {
    const busqueda = this.searchTerm.toLowerCase().trim();
    this.vehiculosFiltrados = this.fallasData.filter(f => 
      (f.placa && f.placa.toLowerCase().includes(busqueda)) || 
      (f.destino && f.destino.toLowerCase().includes(busqueda))
    );
  }

  startEdit(item: any) { this.editId = item._id || item.id; }

  saveEdit() {
    const item = this.fallasData.find(f => (f._id === this.editId || f.id === this.editId));
    if (item) {
      this.http.put(`${this.baseApiUrl}/entregas/${item._id || item.id}`, item).subscribe({
        next: () => {
          this.editId = null;
          this.cargarDatos(); 
        }
      });
    }
  }

  cancelEdit() { this.editId = null; this.cargarDatos(); }

  eliminarPunto(id: string) {
    if (id && confirm('¿Desea eliminar este registro?')) {
      this.http.delete(`${this.baseApiUrl}/entregas/${id}`).subscribe(() => this.cargarDatos());
    }
  }

  private renderCharts() {
    if (!isPlatformBrowser(this.platformId) || !this.lineChartCanvas || !this.doughnutChartCanvas) return;
    if (this.lineChart) this.lineChart.destroy();
    if (this.doughnutChart) this.doughnutChart.destroy();

    const labels = this.fallasData.map(f => f.destino || f.placa);
    
    this.lineChart = new Chart(this.lineChartCanvas.nativeElement, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          { label: 'Entregas', data: this.fallasData.map(f => f.entregas), borderColor: '#481165', fill: true },
          { label: 'Devoluciones', data: this.fallasData.map(f => f.devoluciones), borderColor: '#ff4d4d', fill: true }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });

    this.doughnutChart = new Chart(this.doughnutChartCanvas.nativeElement, {
      type: 'doughnut', 
      data: {
        labels: ['Bueno', 'Regular', 'Malo', 'En Taller'],
        datasets: [{
          data: [this.vBueno, this.vRegular, this.vMalo, this.vTaller],
          backgroundColor: ['#1d6f42', '#ffc107', '#fd7e14', '#dc3545']
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}