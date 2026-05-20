import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-alertas',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './alertas.html',
  styleUrls: ['./alertas.css']
})
export class AlertasComponent implements OnInit {
  public listaMantenimientos: any[] = [];
  public alertasCriticas: any[] = []; // Rojo: Vencidos
  public alertasProximas: any[] = [];  // Amarillo: Vencen pronto
  public flotaAlDia: number = 0;

  private apiUrl = 'https://makand-sas.onrender.com/api/mantenimiento';

  constructor(private http: HttpClient) {}

  ngOnInit() {
    this.obtenerAlertas();
  }

  obtenerAlertas() {
    this.http.get<any[]>(this.apiUrl).subscribe({
      next: (data) => {
        this.procesarAlertas(data);
      },
      error: (err) => console.error('Error al obtener alertas:', err)
    });
  }

  private procesarAlertas(data: any[]) {
    const hoy = new Date().getTime();
    
    // Reiniciamos contadores
    this.alertasCriticas = [];
    this.alertasProximas = [];
    let contadorAlDia = 0;

    data.forEach(m => {
      // Tomamos la fecha de vencimiento más próxima (SOAT, Tecno o Próximo Mantenimiento)
      const fechas = [
        m.vencimientoSoat ? new Date(m.vencimientoSoat + 'T00:00:00').getTime() : null,
        m.vencimientoTecno ? new Date(m.vencimientoTecno + 'T00:00:00').getTime() : null,
        m.fechaProxima ? new Date(m.fechaProxima + 'T00:00:00').getTime() : null
      ].filter(f => f !== null) as number[];

      if (fechas.length > 0) {
        const masProxima = Math.min(...fechas);
        const diasRestantes = (masProxima - hoy) / (1000 * 60 * 60 * 24);

        const registroAlerta = {
          placa: m.placa,
          conductor: m.conductor,
          tipo: diasRestantes < 0 ? 'Vencido' : 'Próximo',
          dias: Math.floor(diasRestantes),
          documento: this.identificarDocumentoCritico(m, masProxima)
        };

        if (diasRestantes < 0) {
          this.alertasCriticas.push(registroAlerta);
        } else if (diasRestantes <= 15) {
          this.alertasProximas.push(registroAlerta);
        } else {
          contadorAlDia++;
        }
      }
    });

    this.flotaAlDia = contadorAlDia;
  }

  private identificarDocumentoCritico(m: any, fechaMilis: number): string {
    if (m.vencimientoSoat && new Date(m.vencimientoSoat + 'T00:00:00').getTime() === fechaMilis) return 'SOAT';
    if (m.vencimientoTecno && new Date(m.vencimientoTecno + 'T00:00:00').getTime() === fechaMilis) return 'Tecnomecánica';
    return 'Mantenimiento';
  }
}