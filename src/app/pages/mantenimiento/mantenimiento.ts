import { Component, OnInit, Inject, PLATFORM_ID, ChangeDetectorRef, NgZone } from '@angular/core'; 
import { isPlatformBrowser } from '@angular/common';
import * as XLSX from 'xlsx';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { Chart, registerables } from 'chart.js';

Chart.register(...registerables);

@Component({
  selector: 'app-mantenimiento',
  templateUrl: './mantenimiento.html',
  standalone: true,
  imports: [CommonModule, FormsModule],
  styleUrls: ['./mantenimiento.css']
})
export class MantenimientoComponent implements OnInit {

  public isEditing: boolean = false;
  public editId: string | null = null; 
  public historico: any[] = [];
  public intentoGuardar: boolean = false;
  public mostrarForm: boolean = true; 
  public itemSeleccionado: any = null;

  // Ajustado a singular para coincidir con la ruta más estable del backend
  private apiUrl = 'https://makand-sas.onrender.com/api/mantenimiento'; 
  private chart: any; 

  public registro: any = {
    placa: '',
    marcaModelo: '',
    numMotor: '',
    conductor: '',
    kilometraje: 0,
    fechaEntrada: '',
    fechaProxima: '',
    tipoIntervencion: 'Preventivo',
    taller: '',
    estadoVehiculo: 'En Taller',
    descripcion: '',
    repuestos: '',
    vencimientoSoat: '',
    vencimientoTecno: '',
    costoManoObra: 0,
    costoTotal: 0,
    fotos: [],
    checklist: {
      frenos: false,
      aceite: false,
      neumaticos: false,
      bateria: false,
      luces: false,
      refrigerante: false
    }
  };

  constructor(
    @Inject(PLATFORM_ID) private platformId: Object,
    private http: HttpClient,
    private cd: ChangeDetectorRef,
    private ngZone: NgZone 
  ) {}

  ngOnInit() {
    if (isPlatformBrowser(this.platformId)) {
      this.cargarDatos();
    }
  }

  cargarDatos() {
    this.http.get<any[]>(this.apiUrl).subscribe({
      next: (data) => {
        this.ngZone.run(() => {
          this.historico = data || [];
          this.actualizarGrafica();
          this.cd.detectChanges();
        });
      },
      error: (err) => {
        console.error('Error al conectar con la API:', err);
        // Fallback al almacenamiento local si el servidor falla
        if (isPlatformBrowser(this.platformId)) {
          const db = localStorage.getItem('db_mantenimientos');
          if (db) {
            this.ngZone.run(() => {
              this.historico = JSON.parse(db);
              this.actualizarGrafica();
              this.cd.detectChanges();
            });
          }
        }
      }
    });
  }

  guardarMantenimiento() {
    this.intentoGuardar = true;

    if (!this.esValido('placa')) {
      alert('⚠️ Formato de placa inválido (Debe ser ABC123).');
      return;
    }
    if (!this.esValido('conductor')) {
      alert('⚠️ El nombre del conductor no es válido.');
      return;
    }
    if (!this.registro.fechaEntrada || !this.registro.fechaProxima) {
      alert('⚠️ Las fechas de entrada y próxima cita son obligatorias.');
      return;
    }

    // Normalización de placa antes de enviar
    const registroParaEnviar = { ...this.registro };
    registroParaEnviar.placa = registroParaEnviar.placa.toUpperCase().replace(/\s/g, '');

    if (this.isEditing && this.editId) {
      this.http.put(`${this.apiUrl}/${this.editId}`, registroParaEnviar).subscribe({
        next: () => {
          alert('✅ Registro actualizado correctamente');
          this.finalizarOperacion();
        },
        error: (err) => alert('Error al actualizar: ' + err.message)
      });
    } else {
      this.http.post(this.apiUrl, registroParaEnviar).subscribe({
        next: () => {
          alert('✅ Mantenimiento guardado con éxito');
          this.finalizarOperacion();
        },
        error: (err) => alert('Error al guardar: ' + err.message)
      });
    }
  }

  private finalizarOperacion() {
    this.resetForm();
    this.cargarDatos(); 
  }

  verDetalle(item: any) {
    this.itemSeleccionado = item;
  }

  cerrarModal() {
    this.itemSeleccionado = null;
  }

  actualizarGrafica() {
    if (!isPlatformBrowser(this.platformId)) return;

    const enTaller = this.historico.filter(x => x.estadoVehiculo === 'En Taller').length;
    const operativos = this.historico.filter(x => x.estadoVehiculo === 'Finalizado' || x.estadoVehiculo === 'Operativo').length;

    // Uso de setTimeout para asegurar que el DOM esté listo
    setTimeout(() => {
      const canvas = document.getElementById('estadoChart') as HTMLCanvasElement;
      if (!canvas) return;
      
      if (this.chart) {
        this.chart.destroy();
      }

      this.chart = new Chart(canvas, {
        type: 'pie',
        data: {
          labels: ['En Taller', 'Operativos'],
          datasets: [{
            data: [enTaller, operativos],
            backgroundColor: ['#481165', '#0e6836'], // Ajustado al morado corporativo #481165
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom' }
          }
        }
      });
      this.cd.detectChanges();
    }, 50);
  }

  get totalEnTaller() {
    return this.historico.filter(x => x.estadoVehiculo === 'En Taller').length;
  }

  get proximosMantenimientos() {
    const hoy = new Date();
    const proximaSemana = new Date();
    proximaSemana.setDate(hoy.getDate() + 7);

    return this.historico.filter(x => {
      if (!x.fechaProxima) return false;
      const fechaMante = new Date(x.fechaProxima);
      return fechaMante >= hoy && fechaMante <= proximaSemana;
    }).length;
  }

  esValido(campo: string): boolean {
    const valor = this.registro[campo];
    if (!this.intentoGuardar && (valor === '' || valor === 0 || valor === null)) return true;

    switch (campo) {
      case 'conductor':
      case 'marcaModelo':
      case 'taller':
        return /^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s.]{3,60}$/.test(valor || '');
      case 'placa':
        const placaLimpia = (this.registro.placa || '').toUpperCase().replace(/\s/g, '');
        return /^[A-Z]{3}[0-9]{3}$/.test(placaLimpia);
      case 'kilometraje':
      case 'costoManoObra':
      case 'costoTotal':
        return valor !== null && valor >= 0;
      case 'fechaEntrada':
      case 'fechaProxima':
        return !!valor;
      default:
        return true;
    }
  }

  onFilesSelected(event: any) {
    const files = event.target.files;
    if (!files) return;

    for (let file of files) {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        this.ngZone.run(() => {
          this.registro.fotos.push(e.target.result);
          this.cd.detectChanges();
        });
      };
      reader.readAsDataURL(file);
    }
  }

  cargarParaEditar(item: any) {
    this.isEditing = true;
    this.editId = item._id; 
    this.registro = JSON.parse(JSON.stringify(item));
    this.intentoGuardar = false; 
    this.mostrarForm = true;
    if (isPlatformBrowser(this.platformId)) {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
    this.cd.detectChanges();
  }

  eliminar(id: string) {
    if (!id) return;
    if (confirm('¿Desea eliminar este registro permanentemente?')) {
      this.http.delete(`${this.apiUrl}/${id}`).subscribe({
        next: () => {
          alert('🗑️ Registro eliminado');
          this.cargarDatos(); 
        },
        error: (err) => console.error('Error al eliminar:', err)
      });
    }
  }

  resetForm() {
    this.isEditing = false;
    this.editId = null;
    this.intentoGuardar = false;
    this.registro = {
      placa: '', marcaModelo: '', numMotor: '', conductor: '',
      kilometraje: 0, fechaEntrada: '', fechaProxima: '',
      tipoIntervencion: 'Preventivo', taller: '', estadoVehiculo: 'En Taller',
      descripcion: '', repuestos: '', vencimientoSoat: '', vencimientoTecno: '',
      costoManoObra: 0, costoTotal: 0, fotos: [],
      checklist: { frenos: false, aceite: false, neumaticos: false, bateria: false, luces: false, refrigerante: false }
    };
    this.cd.detectChanges();
  }

  exportarExcel() {
    if (this.historico.length === 0) {
      alert('No hay datos para exportar');
      return;
    }
    const dataExcel = this.historico.map(h => ({
      Fecha: h.fechaEntrada,
      Placa: h.placa,
      Conductor: h.conductor,
      Tipo: h.tipoIntervencion,
      Costo_Mano_Obra: h.costoManoObra,
      Costo_Total: h.costoTotal,
      Check_Frenos: h.checklist?.frenos ? 'OK' : 'X',
      Check_Aceite: h.checklist?.aceite ? 'OK' : 'X',
      Proxima_Cita: h.fechaProxima,
      Estado: h.estadoVehiculo
    }));

    const ws = XLSX.utils.json_to_sheet(dataExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mantenimientos");
    XLSX.writeFile(wb, "Historial_Completo_Mantenimiento.xlsx");
  }
}