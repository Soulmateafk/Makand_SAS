import { Component, OnInit, Inject, PLATFORM_ID, ChangeDetectorRef } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-conductores',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './conductores.html',
  styleUrls: ['./conductores.css']
})
export class ConductoresComponent implements OnInit {
  public conductores: any[] = [];
  public isEditing: boolean = false;
  public editIndex: number | null = null;
  public intentoGuardar: boolean = false;
  
  private apiUrl = 'http://localhost:3000/api/conductores';

  public registro: any = {
    nombre: '',
    fechaNacimiento: '',
    cedula: '',
    telefono: '',
    fechaUnion: '',
    email: ''
  };

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cdr: ChangeDetectorRef // Crucial para refrescar la tabla inferior
  ) {}

  ngOnInit() {
    if (isPlatformBrowser(this.platformId)) {
      this.cargarDesdeDB();
    }
  }

  // --- CARGAR DATOS DESDE MONGODB ---
  cargarDesdeDB() {
    if (isPlatformBrowser(this.platformId)) {
      this.http.get<any[]>(this.apiUrl).subscribe({
        next: (res) => {
          this.conductores = res || [];
          this.cdr.detectChanges(); // Fuerza a Angular a pintar los datos recibidos
        },
        error: (err) => {
          console.error('❌ Error al cargar datos', err);
          this.cdr.detectChanges();
        }
      });
    }
  }

  // --- VALIDACIONES ---
  esValido(campo: string): boolean {
    if (!this.intentoGuardar) return true;
    const v = this.registro[campo];
    switch (campo) {
      case 'nombre': return v && /^[a-zA-ZÀ-ÿ\s]+$/.test(v) && v.trim().length > 3;
      case 'cedula': return v && /^[0-9]{7,10}$/.test(v);
      case 'telefono': return v && /^[0-9]{10}$/.test(v);
      case 'email': return v && /^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,4}$/.test(v);
      case 'fechaNacimiento':
      case 'fechaUnion': return v && v !== '';
      default: return true;
    }
  }

  // --- GUARDAR O ACTUALIZAR ---
  guardarConductor() {
    this.intentoGuardar = true;

    const valido = this.esValido('nombre') && this.esValido('cedula') && 
                   this.esValido('telefono') && this.esValido('email') &&
                   this.esValido('fechaNacimiento') && this.esValido('fechaUnion');

    if (!valido) {
      alert('⚠️ Verifique que todos los campos sean correctos.');
      return;
    }

    // Objeto mapeado EXACTAMENTE igual a tu conductores.js (Node.js)
    const datosParaEnviar = {
      nombre: this.registro.nombre,
      cedula: this.registro.cedula,
      fechaNacimiento: this.registro.fechaNacimiento,
      telefono: this.registro.telefono,
      fechaUnion: this.registro.fechaUnion,
      email: this.registro.email
    };

    if (this.isEditing && this.registro._id) {
      // ACTUALIZAR
      this.http.put(`${this.apiUrl}/${this.registro._id}`, datosParaEnviar).subscribe({
        next: () => {
          alert('✅ Cambios actualizados correctamente');
          this.cargarDesdeDB(); // Recarga y refresca vista
          this.resetForm();
        },
        error: (err) => console.error('Error al actualizar:', err)
      });
    } else {
      // GUARDAR NUEVO
      this.http.post(this.apiUrl, datosParaEnviar).subscribe({
        next: (res) => {
          alert('✅ Guardado en MongoDB');
          this.cargarDesdeDB(); // Recarga y refresca vista
          this.resetForm();
        },
        error: (err) => {
          console.error(err);
          alert('❌ Error al conectar con el servidor');
        }
      });
    }
  }

  cargarParaEditar(conductor: any) {
    this.isEditing = true;
    this.registro = { ...conductor }; 
    
    if (isPlatformBrowser(this.platformId)) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
    this.cdr.detectChanges();
  }

  eliminarConductor(id: string) {
    if (!id) return;
    if (confirm('¿Eliminar permanentemente a este conductor?')) {
      this.http.delete(`${this.apiUrl}/${id}`).subscribe({
        next: () => {
          this.cargarDesdeDB(); // Recarga y refresca vista
        },
        error: (err) => console.error('Error al eliminar:', err)
      });
    }
  }

  resetForm() {
    this.isEditing = false;
    this.editIndex = null;
    this.intentoGuardar = false;
    this.registro = { 
      nombre: '', 
      fechaNacimiento: '', 
      cedula: '', 
      telefono: '', 
      fechaUnion: '', 
      email: '' 
    };
    this.cdr.detectChanges();
  }
}