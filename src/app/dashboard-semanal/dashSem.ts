import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

export const dashSem = 'dashboardsemanal';

@Component({
  selector: 'app-dash-semanal',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="dashboard-container">
      <header class="dashboard-header">
        <h1>Panel de Control Semanal</h1>
        <p>Monitoreo de efectividad y cumplimiento logístico en tiempo real</p>
      </header>

      <div class="dashboard-main-grid">
        
        <div class="metrics-column">
          <div class="card total">
            <h3>Total Viajes Programados</h3>
            <p class="metric-value">{{ totalViajes }}</p>
          </div>
          
          <div class="card alert">
            <h3>Novedades / Retrasos</h3>
            <div class="input-container">
              <input 
                type="number" 
                [(ngModel)]="retrasos" 
                (input)="calcularEfectividad()" 
                min="0" 
                [max]="totalViajes" 
              />
            </div>
          </div>
        </div>

        <div class="card chart-card">
          <h3>Efectividad de la Operación</h3>
          <div class="radial-chart-wrapper">
            <svg viewBox="0 0 120 120" class="radial-svg">
              <circle class="radial-bg" cx="60" cy="60" r="50"></circle>
              <circle 
                class="radial-progress" 
                cx="60" 
                cy="60" 
                r="50"
                [style.strokeDashoffset]="strokeDashoffset"
              ></circle>
            </svg>
            <div class="chart-overlay">
              <span class="chart-percentage">{{ porcentajeEfectividad }}%</span>
              <span class="chart-label">On-Time</span>
            </div>
          </div>
        </div>

      </div>

      <section class="table-section">
        <div class="table-header">
          <h2>Estado de las Rutas Recientes</h2>
        </div>
        
        <div class="table-responsive">
          <table>
            <thead>
              <tr>
                <th>Identificador de Ruta</th>
                <th>Vehículo Asignado</th>
                <th>Hora de Ingreso</th>
                <th>Hora de Salida</th>
                <th>Estado de Cumplimiento</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let operacion of registroOperaciones">
                <td>{{ operacion.ruta }}</td>
                <td><span class="badge-plate">{{ operacion.placa }}</span></td>
                <td>{{ operacion.horaIngreso }}</td>
                <td>{{ operacion.horaSalida }}</td>
                <td>
                  <span [ngClass]="operacion.cumple ? 'status-ok' : 'status-fail'">
                    {{ operacion.cumple ? 'CUMPLE' : 'NO CUMPLE' }}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>
    </div>
  `,
  styles: [`
    /* Contenedor Principal Dark Mode */
    .dashboard-container {
      padding: 32px;
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background-color: #0f111a;
      color: #f8fafc;
      min-height: 100vh;
    }
    .dashboard-header h1 {
      margin: 0;
      color: #ffffff;
      font-size: 28px;
      font-weight: 700;
      letter-spacing: -0.5px;
    }
    .dashboard-header p {
      margin: 6px 0 0 0;
      color: #64748b;
      font-size: 14px;
    }

    /* Distribución en Grid */
    .dashboard-main-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 24px;
      margin: 32px 0;
    }
    @media (max-width: 768px) {
      .dashboard-main-grid { grid-template-columns: 1fr; }
    }

    .metrics-column {
      display: flex;
      flex-direction: column;
      gap: 24px;
    }

    /* Tarjetas Estilo Cyber/Premium */
    .card {
      background: #161925;
      padding: 24px;
      border-radius: 16px;
      box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
      border: 1px solid #23273a;
      position: relative;
      overflow: hidden;
    }
    .card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; width: 4px; height: 100%;
      background: #3b82f6;
    }
    .card.alert::before { background: #f43f5e; }
    .card.chart-card::before { background: #10b981; }

    .card h3 {
      margin: 0 0 16px 0;
      font-size: 12px;
      text-transform: uppercase;
      color: #94a3b8;
      letter-spacing: 1px;
    }
    .metric-value {
      margin: 0;
      font-size: 42px;
      font-weight: 800;
      color: #ffffff;
    }

    /* Inputs Estilizados */
    .input-container input {
      width: 100%;
      max-width: 140px;
      padding: 8px 16px;
      font-size: 24px;
      font-weight: 700;
      background: #0f111a;
      border: 2px solid #2e344e;
      border-radius: 8px;
      color: #f43f5e;
      text-align: center;
    }
    .input-container input:focus {
      border-color: #f43f5e;
      outline: none;
      box-shadow: 0 0 12px rgba(244, 63, 94, 0.2);
    }

    /* Gráfico Radial Nativo (Pure SVG) */
    .chart-card {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
    }
    .radial-chart-wrapper {
      position: relative;
      width: 180px;
      height: 180px;
    }
    .radial-svg {
      transform: rotate(-90deg);
      width: 100%;
      height: 100%;
    }
    .radial-bg {
      fill: none;
      stroke: #1f2335;
      stroke-width: 10;
    }
    .radial-progress {
      fill: none;
      stroke: #10b981; /* Verde esmeralda neón */
      stroke-width: 10;
      stroke-linecap: round;
      stroke-dasharray: 314.16; /* Circunferencia exacta (2 * pi * 50) */
      transition: stroke-dashoffset 0.4s ease-out;
    }
    .chart-overlay {
      position: absolute;
      top: 0; left: 0; width: 100%; height: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
    }
    .chart-percentage {
      font-size: 32px;
      font-weight: 800;
      color: #ffffff;
    }
    .chart-label {
      font-size: 11px;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: 1px;
    }

    /* Tabla de Datos Modernizada */
    .table-section {
      background: #161925;
      border-radius: 16px;
      border: 1px solid #23273a;
      padding: 24px;
      box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
    }
    .table-header h2 {
      margin: 0 0 20px 0;
      color: #ffffff;
      font-size: 18px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
    }
    th {
      background-color: #0f111a;
      color: #94a3b8;
      padding: 16px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border-bottom: 2px solid #23273a;
    }
    td {
      padding: 16px;
      font-size: 14px;
      color: #cbd5e1;
      border-bottom: 1px solid #1f2335;
    }
    tr:hover td { background-color: #1e2235; }
    
    .badge-plate {
      background: #2e344e;
      color: #f8fafc;
      padding: 4px 10px;
      border-radius: 6px;
      font-family: monospace;
      font-weight: 600;
    }
    .status-ok {
      color: #10b981;
      background: rgba(16, 185, 129, 0.1);
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
    }
    .status-fail {
      color: #f43f5e;
      background: rgba(244, 63, 94, 0.1);
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
    }
  `]
})
export class DashSemanalComponent implements OnInit {
  
  totalViajes: number = 100;
  retrasos: number = 10;
  porcentajeEfectividad: number = 90;
  strokeDashoffset: number = 0; // Controla visualmente el arco del SVG

  registroOperaciones = [
    { ruta: 'Ruta Zona Norte', placa: 'WFH623', horaIngreso: '05:51:12', horaSalida: '11:45:02', cumple: true },
    { ruta: 'Ruta Distribución Centro', placa: 'PRZ066', horaIngreso: '04:44:07', horaSalida: '10:23:13', cumple: true },
    { ruta: 'Ruta Sabana Siberia', placa: 'KNM116', horaIngreso: '06:17:07', horaSalida: '11:03:57', cumple: false }
  ];

  ngOnInit(): void {
    this.calcularEfectividad();
  }

  calcularEfectividad(): void {
    if (this.retrasos < 0) this.retrasos = 0;
    if (this.retrasos > this.totalViajes) this.retrasos = this.totalViajes;

    const entregasExitosas = this.totalViajes - this.retrasos;
    this.porcentajeEfectividad = Math.round((entregasExitosas / this.totalViajes) * 100);

    // Actualiza la animación del anillo SVG basado en la circunferencia base (314.16)
    const circunferencia = 314.16;
    this.strokeDashoffset = circunferencia - (circunferencia * this.porcentajeEfectividad) / 100;
  }
}