import { Component, OnInit, Inject, PLATFORM_ID, ChangeDetectorRef } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-combustible',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './combustible.html',
  styleUrls: ['./combustible.css']
})
export class CombustibleComponent implements OnInit {
  // Listas de datos
  public listaTanqueos: any[] = [];
  public tanqueosFiltrados: any[] = [];
  
  // Control de interfaz
  public mostrarForm: boolean = false;
  public editandoId: string | null = null;
  public searchTerm: string = '';
  public errorPlaca: string = '';

  // API URL conectada al backend de MongoDB
  private apiUrl = 'https://makand-sas.onrender.com/api/combustible';

  // Opciones del formulario
  public listaMetodosPago: string[] = ['Efectivo', 'Tarjeta', 'Chip Corporativo', 'Vale'];

  // Objeto del formulario
  public nuevoTanqueo = {
    placa: '',
    conductor: '',
    fecha: '',
    kmActual: 0,
    galones: 0,
    precioTotal: 0,
    estacionServicio: '',
    metodoPago: 'Efectivo'
  };

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit() {
    this.cargarDatos();
  }

  cargarDatos() {
    this.http.get<any[]>(this.apiUrl).subscribe({
      next: (data) => {
        this.listaTanqueos = data;
        this.tanqueosFiltrados = [...data];
        this.cdr.detectChanges();
      },
      error: (err) => console.error('❌ Error al conectar con MongoDB:', err)
    });
  }

  filtrarTanqueos() {
    const busqueda = this.searchTerm.toLowerCase().trim();
    if (!busqueda) {
      this.tanqueosFiltrados = [...this.listaTanqueos];
      return;
    }
    this.tanqueosFiltrados = this.listaTanqueos.filter(t => 
      (t.placa && t.placa.toLowerCase().includes(busqueda)) ||
      (t.conductor && t.conductor.toLowerCase().includes(busqueda)) ||
      (t.estacionServicio && t.estacionServicio.toLowerCase().includes(busqueda))
    );
  }

  guardarTanqueo() {
    // Validar formato de placa Colombiana
    const regexPlaca = /^[A-Z]{3}-?\d{3}$/i;
    if (!regexPlaca.test(this.nuevoTanqueo.placa)) {
      this.errorPlaca = 'Formato inválido (Ej: ABC123)';
      return;
    }
    this.errorPlaca = '';

    if (this.editandoId) {
      // Modificar registro existente en Mongo
      this.http.put(`${this.apiUrl}/${this.editandoId}`, this.nuevoTanqueo).subscribe({
        next: () => {
          this.cargarDatos();
          this.finalizarFormulario();
        },
        error: (err) => console.error('Error al actualizar en Mongo:', err)
      });
    } else {
      // Crear nuevo registro en Mongo
      this.http.post(this.apiUrl, this.nuevoTanqueo).subscribe({
        next: () => {
          this.cargarDatos();
          this.finalizarFormulario();
        },
        error: (err) => console.error('Error al guardar en Mongo:', err)
      });
    }
  }

  startEdit(item: any) {
    this.editandoId = item._id || item.id;
    // Formatear la fecha para el input type="date" (YYYY-MM-DD)
    const fechaFormateada = item.fecha ? new Date(item.fecha).toISOString().split('T')[0] : '';
    
    this.nuevoTanqueo = {
      placa: item.placa,
      conductor: item.conductor,
      fecha: fechaFormateada,
      kmActual: item.kmActual,
      galones: item.galones,
      precioTotal: item.precioTotal,
      estacionServicio: item.estacionServicio || '',
      metodoPago: item.metodoPago || 'Efectivo'
    };
    this.mostrarForm = true;
    if (isPlatformBrowser(this.platformId)) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }

  eliminarTanqueo(id: string) {
    if (confirm('¿Estás seguro de eliminar este registro de tanqueo?')) {
      this.http.delete(`${this.apiUrl}/${id}`).subscribe({
        next: () => this.cargarDatos(),
        error: (err) => console.error('Error al eliminar de Mongo:', err)
      });
    }
  }

  // REPORTE GENERAL: Descarga todo el histórico sin filtros
  exportarExcelGeneral() {
    const exportUrl = 'https://makand-sas.onrender.com/api/combustibles/exportar-excel';
    if (isPlatformBrowser(this.platformId)) {
      window.open(exportUrl, '_blank');
    }
  }

  // REPORTE ESPECÍFICO: Envía la búsqueda actual como query parameter al backend
  exportarExcelEspecifico() {
    const busqueda = this.searchTerm ? this.searchTerm.trim() : '';
    const queryParam = busqueda ? `?buscar=${encodeURIComponent(busqueda)}` : '';
    const exportUrl = `https://makand-sas.onrender.com/api/combustibles/exportar-excel${queryParam}`;
    
    if (isPlatformBrowser(this.platformId)) {
      window.open(exportUrl, '_blank');
    }
  }

  finalizarFormulario() {
    this.editandoId = null;
    this.nuevoTanqueo = {
      placa: '',
      conductor: '',
      fecha: '',
      kmActual: 0,
      galones: 0,
      precioTotal: 0,
      estacionServicio: '',
      metodoPago: 'Efectivo'
    };
    this.mostrarForm = false;
    this.cdr.detectChanges();
  }
}