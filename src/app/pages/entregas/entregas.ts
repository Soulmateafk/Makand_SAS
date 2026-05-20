import { Component, OnInit, AfterViewInit, Inject, PLATFORM_ID, ChangeDetectorRef } from '@angular/core';
import { CommonModule, isPlatformBrowser, DecimalPipe } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Chart, registerables } from 'chart.js';
import { HttpClient } from '@angular/common/http';
import * as XLSX from 'xlsx';
import { AuthService } from '../../login/auth.service';

@Component({
  selector: 'app-entregas',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './entregas.html',
  styleUrls: ['./entregas.css'],
  providers: [DecimalPipe]
})
export class Entregas implements OnInit, AfterViewInit {
  public listaEntregas: any[] = [];
  private datosOriginales: any[] = []; 
  public mostrarForm = false;
  public editandoId: string | null = null;
  public entregaSeleccionada: any = null;

  private barChart: any;
  private pieChart: any;

  public filtroPlaca: string = ''; 
  public errorPlaca: string = ''; 

  private apiUrl = 'https://makand-sas.onrender.com/api/entregas';

  public anioSeleccionado: number = 2026; 
  public semanaSeleccionada: number = 0;
  public diaSeleccionado: string = '';
  
  public fechaInicio: string = '';
  public fechaFin: string = '';
  
  public listaAnios: number[] = [2024, 2025, 2026, 2027];
  public listaSemanas: any[] = [];
  public listaDiasDisponibles: string[] = [];

  public nuevoPunto = { 
    placa: '', 
    nombre: '', 
    origen: '', 
    entregas: 0, 
    devoluciones: 0,
    conductor: '', 
    estadoProducto: 'Óptimo (Sin novedades)', 
    explicacion: '', 
    foto: null as string | null,
    fecha: '' 
  };

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cdr: ChangeDetectorRef,
    @Inject(AuthService) public authService: AuthService 
  ) {
    Chart.register(...registerables);
    this.generarSemanas();
  }

  ngOnInit() {
    this.cargarDatos();
  }

  ngAfterViewInit() {}

  generarSemanas() {
    this.listaSemanas = [];
    for (let i = 1; i <= 52; i++) {
      this.listaSemanas.push({ num: i, label: `Semana ${i}` });
    }
  }

  buscarPorRango() {
    if (this.fechaInicio && this.fechaFin) {
      const inicio = new Date(this.fechaInicio + 'T00:00:00').getTime();
      const fin = new Date(this.fechaFin + 'T23:59:59').getTime();
      this.listaEntregas = this.datosOriginales.filter(p => {
        const fechaRegistro = new Date(p.fecha).getTime();
        return fechaRegistro >= inicio && fechaRegistro <= fin;
      });
      this.actualizarVistasPostFiltro();
    } else {
      this.listaEntregas = [...this.datosOriginales];
      this.actualizarVistasPostFiltro();
    }
  }

  filtrarPorHistorial() {
    this.fechaInicio = '';
    this.fechaFin = '';
    let url = `${this.apiUrl}/historial?anio=${this.anioSeleccionado}`;
    if (this.semanaSeleccionada > 0) url += `&semana=${this.semanaSeleccionada}`;
    if (this.diaSeleccionado) url += `&fechaExacta=${this.diaSeleccionado}`;
    this.http.get<any[]>(url).subscribe({
      next: (data) => {
        this.procesarLista(data);
        if (this.semanaSeleccionada > 0) {
          this.listaDiasDisponibles = [...new Set(data.map(item => new Date(item.fecha).toISOString().split('T')[0]))];
        }
      },
      error: (err) => console.error('Error filtrando historial:', err)
    });
  }

  limpiarFiltros() {
    this.semanaSeleccionada = 0;
    this.diaSeleccionado = '';
    this.fechaInicio = '';
    this.fechaFin = '';
    this.anioSeleccionado = 2026;
    this.cargarDatos();
  }

  cargarDatos() {
    this.http.get<any[]>(this.apiUrl).subscribe({
      next: (data) => this.procesarLista(data),
      error: (err) => console.error('❌ Error al conectar con MongoDB:', err)
    });
  }

  private procesarLista(data: any[]) {
    const procesados = data.map(item => {
      const ent = Number(item.entregas) || 0;
      const dev = Number(item.devoluciones) || 0;
      let porcentajeEfectividad = ent > 0 ? ((ent - dev) / ent) * 100 : 0;
      if (porcentajeEfectividad < 0) porcentajeEfectividad = 0;
      return { ...item, id: item._id || item.id, entregas: ent, devoluciones: dev, tasa: porcentajeEfectividad.toFixed(1) };
    });
    this.datosOriginales = [...procesados];
    this.listaEntregas = [...procesados];
    this.actualizarVistasPostFiltro();
  }

  private actualizarVistasPostFiltro() {
    this.cdr.detectChanges();
    if (isPlatformBrowser(this.platformId)) this.actualizarGraficas();
  }

  get totalEntregas() { 
    const totalCargadoRuta = this.listaEntregas.reduce((a, b) => a + (Number(b.entregas) || 0), 0);
    const realesExitosas = totalCargadoRuta - this.totalDevoluciones;
    return realesExitosas > 0 ? realesExitosas : 0;
  }
  
  get totalDevoluciones() { return this.listaEntregas.reduce((a, b) => a + (Number(b.devoluciones) || 0), 0); }

  get puntosFiltrados() {
    if (!this.filtroPlaca) return this.listaEntregas;
    const busqueda = this.filtroPlaca.toLowerCase();
    return this.listaEntregas.filter(p => 
      (p.placa?.toLowerCase().includes(busqueda)) || (p.nombre?.toLowerCase().includes(busqueda)) ||
      (p.origen?.toLowerCase().includes(busqueda)) || (p.conductor?.toLowerCase().includes(busqueda))
    );
  }

  agregarPunto() {
    const regexPlaca = /^[A-Z]{3}-?\d{3}$/i;
    if (!regexPlaca.test(this.nuevoPunto.placa)) { this.errorPlaca = 'Formato inválido (Ej: ABC123)'; return; }
    this.errorPlaca = ''; 
    if (this.editandoId) {
      this.http.put(`${this.apiUrl}/${this.editandoId}`, this.nuevoPunto).subscribe({
        next: () => { this.cargarDatos(); this.finalizarAccion(); },
        error: (err) => console.error('Error al actualizar:', err)
      });
    } else {
      this.http.post(this.apiUrl, this.nuevoPunto).subscribe({
        next: () => { this.cargarDatos(); this.finalizarAccion(); },
        error: (err) => console.error('Error al guardar:', err)
      });
    }
  }

  editarPunto(punto: any) {
    this.editandoId = punto._id || punto.id; 
    this.nuevoPunto = { ...punto };
    this.mostrarForm = true;
    if (isPlatformBrowser(this.platformId)) window.scrollTo({ top: 0, behavior: 'smooth' });
    this.cdr.detectChanges();
  }

  eliminarPunto(id: any) {
    if (confirm('¿Estás seguro de eliminar este registro?')) {
      this.http.delete(`${this.apiUrl}/${id}`).subscribe({
        next: () => this.cargarDatos(),
        error: (err) => console.error('Error al eliminar:', err)
      });
    }
  }

  exportarTodoExcel() { this.generarExcel(this.listaEntregas, 'Reporte_General_Logistica'); }
  exportarIndividualExcel(entrega: any) { this.generarExcel([entrega], `Ficha_Entrega_${entrega.placa}_${entrega.nombre}`); }

  private generarExcel(datos: any[], nombreArchivo: string) {
    const dataReporte = datos.map(p => ({
      'Fecha': p.fecha ? new Date(p.fecha).toLocaleDateString() : 'N/A',
      'Vehículo': p.placa, 'Conductor': p.conductor, 'Cliente': p.nombre, 'Destino': p.origen,
      'Entregas': p.entregas, 'Devoluciones': p.devoluciones, 'Estado': p.estadoProducto,
      'Tasa Éxito': p.tasa + '%', 'Observaciones': p.explicacion
    }));
    const worksheet = XLSX.utils.json_to_sheet(dataReporte);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Logística');
    XLSX.writeFile(workbook, `${nombreArchivo}.xlsx`);
  }

  abrirDetalle(entrega: any) { this.entregaSeleccionada = entrega; this.cdr.detectChanges(); }
  cerrarDetalle() { this.entregaSeleccionada = null; }
  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => { this.nuevoPunto.foto = reader.result as string; this.cdr.detectChanges(); };
      reader.readAsDataURL(file);
    }
  }
  finalizarAccion() {
    this.editandoId = null;
    this.nuevoPunto = { placa: '', nombre: '', origen: '', entregas: 0, devoluciones: 0, conductor: '', estadoProducto: 'Óptimo (Sin novedades)', explicacion: '', foto: null, fecha: '' };
    this.mostrarForm = false;
    this.cdr.detectChanges();
  }
  actualizarGraficas() { setTimeout(() => this.initCharts(), 100); }
  initCharts() {
    if (!isPlatformBrowser(this.platformId)) return;
    const ctxBar = document.getElementById('barChart') as HTMLCanvasElement;
    const ctxPie = document.getElementById('pieChart') as HTMLCanvasElement;
    if (this.barChart) this.barChart.destroy();
    if (this.pieChart) this.pieChart.destroy();
    if (ctxBar && this.listaEntregas.length > 0) {
      this.barChart = new Chart(ctxBar, { type: 'bar', data: { labels: this.listaEntregas.map(p => p.nombre || p.placa), datasets: [{ label: 'Entregas', data: this.listaEntregas.map(p => p.entregas), backgroundColor: '#481165' }, { label: 'Devoluciones', data: this.listaEntregas.map(p => p.devoluciones), backgroundColor: '#dc3545' }] }, options: { responsive: true, maintainAspectRatio: false } });
    }
    if (ctxPie && this.listaEntregas.length > 0) {
      this.pieChart = new Chart(ctxPie, { type: 'pie', data: { labels: this.listaEntregas.map(p => p.origen || 'Sin destino'), datasets: [{ data: this.listaEntregas.map(p => p.entregas), backgroundColor: ['#481165', '#9d4edd', '#c8b6ff', '#5a189a'] }] }, options: { responsive: true, maintainAspectRatio: false } });
    }
  }
}