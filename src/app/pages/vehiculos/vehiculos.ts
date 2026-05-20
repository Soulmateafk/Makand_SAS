import { Component, OnInit, Inject, PLATFORM_ID, ChangeDetectorRef } from '@angular/core'; // Agregamos ChangeDetectorRef
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-lista-vehiculos',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './vehiculos.html',
  styleUrls: ['./vehiculos.css']
})
export class ListaVehiculosComponent implements OnInit {
  
  public mostrarFormulario = false;
  public isEditing = false;
  public editId: string | null = null; 
  public listaVehiculos: any[] = [];
  public vehiculosFiltrados: any[] = []; 
  public filtro: string = ''; 

  private apiUrl = 'https://makand-sas.onrender.com/api/vehiculos';

  // Usamos exactamente los mismos nombres del modelo de Node.js
  vehiculo: any = {
    placa: '',
    marcaModelo: '',
    anio: null,
    numFlota: '',
    soat: '',
    tecno: '',
    seguro: '',
    kmActual: null,
    kmProximo: null,
    estado: 'Bueno',
    conductor: '',
    area: ''
  };

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cdr: ChangeDetectorRef // Inyectamos el detector de cambios
  ) {}

  ngOnInit() {
    // Forzamos la carga inicial
    if (isPlatformBrowser(this.platformId)) {
      this.cargarDatos();
    }
  }

  cargarDatos() {
    if (isPlatformBrowser(this.platformId)) {
      console.log('Cargando vehículos desde la base de datos...');
      this.http.get<any[]>(this.apiUrl).subscribe({
        next: (data) => {
          this.listaVehiculos = data || [];
          this.filtrar(); // Sincroniza la tabla inmediatamente
          
          // --- ESTA ES LA CLAVE ---
          // Obliga a Angular a refrescar la vista aunque los datos lleguen tarde
          this.cdr.detectChanges(); 
        },
        error: (err) => {
          console.error('Error al conectar:', err);
          this.listaVehiculos = [];
          this.vehiculosFiltrados = [];
          this.cdr.detectChanges();
        }
      });
    }
  }

  filtrar() {
    const busqueda = this.filtro.toLowerCase().trim();
    if (!busqueda) {
      this.vehiculosFiltrados = [...this.listaVehiculos];
    } else {
      this.vehiculosFiltrados = this.listaVehiculos.filter(v => 
        (v.placa && v.placa.toLowerCase().includes(busqueda)) || 
        (v.conductor && v.conductor.toLowerCase().includes(busqueda)) || 
        (v.marcaModelo && v.marcaModelo.toLowerCase().includes(busqueda))
      );
    }
    this.cdr.detectChanges(); // Refresca la vista después de filtrar
  }

  guardarVehiculo() {
    if (!this.vehiculo.placa) {
      alert("La placa es obligatoria");
      return;
    }

    // CREACIÓN DEL OBJETO LIMPIO: 
    const payload = {
      placa: this.vehiculo.placa.toUpperCase().replace(/\s/g, ''),
      marcaModelo: this.vehiculo.marcaModelo || '',
      anio: this.vehiculo.anio ? Number(this.vehiculo.anio) : null,
      numFlota: this.vehiculo.numFlota || '',
      soat: this.vehiculo.soat || null,
      tecno: this.vehiculo.tecno || null,
      seguro: this.vehiculo.seguro || null,
      kmActual: this.vehiculo.kmActual ? Number(this.vehiculo.kmActual) : 0,
      kmProximo: this.vehiculo.kmProximo ? Number(this.vehiculo.kmProximo) : 0,
      estado: this.vehiculo.estado || 'Bueno',
      conductor: this.vehiculo.conductor || '',
      area: this.vehiculo.area || ''
    };

    if (this.isEditing && this.editId) {
      this.http.put(`${this.apiUrl}/${this.editId}`, payload).subscribe({
        next: () => {
          alert("✅ Actualizado");
          this.cargarDatos(); // Recarga y refresca vista
          this.cancelar();
        },
        error: (err) => alert("Error 400: Revisa que la placa no esté repetida.")
      });
    } else {
      this.http.post(this.apiUrl, payload).subscribe({
        next: () => {
          alert("✅ Registrado");
          this.cargarDatos(); // Recarga y refresca vista
          this.cancelar();
        },
        error: (err) => alert("Error al registrar: " + (err.error?.message || "Dato inválido"))
      });
    }
  }

  editar(vehiculoSeleccionado: any) {
    this.isEditing = true;
    this.editId = vehiculoSeleccionado._id; 
    
    this.vehiculo = { ...vehiculoSeleccionado };
    
    if (this.vehiculo.soat) this.vehiculo.soat = this.formatDate(this.vehiculo.soat);
    if (this.vehiculo.tecno) this.vehiculo.tecno = this.formatDate(this.vehiculo.tecno);
    if (this.vehiculo.seguro) this.vehiculo.seguro = this.formatDate(this.vehiculo.seguro);

    this.mostrarFormulario = true;
    if (isPlatformBrowser(this.platformId)) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }

  formatDate(date: any) {
    if (!date) return '';
    return new Date(date).toISOString().split('T')[0];
  }

  eliminar(id: string) {
    if (confirm('¿Eliminar vehículo?')) {
      this.http.delete(`${this.apiUrl}/${id}`).subscribe({
        next: () => { 
          this.cargarDatos(); // Recarga y refresca vista
        },
        error: (err) => console.error(err)
      });
    }
  }

  cancelar() {
    this.mostrarFormulario = false;
    this.isEditing = false;
    this.editId = null;
    this.vehiculo = {
      placa: '', marcaModelo: '', anio: null, numFlota: '',
      soat: '', tecno: '', seguro: '', kmActual: null,
      kmProximo: null, estado: 'Bueno', conductor: '', area: ''
    };
    this.cdr.detectChanges();
  }
}