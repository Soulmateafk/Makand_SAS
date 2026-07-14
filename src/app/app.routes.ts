import { Routes } from '@angular/router';
import { RouterOutlet, RouterLink } from '@angular/router';

// Asegúrate de que las rutas de los archivos sean correctas según tu estructura de carpetas
import { dashboard } from './dashboard/dashboard';
import { ConductoresComponent} from './pages/conductores/conductores';
import { Entregas } from './pages/entregas/entregas';
import { ListaVehiculosComponent } from './pages/vehiculos/vehiculos';
import { MantenimientoComponent } from './pages/mantenimiento/mantenimiento';
import { LoginComponent } from './login/login';
import { AlertasComponent } from './pages/alertas/alertas';
import { CombustibleComponent } from './components/combustible/combustible';
import { MapaEntregasComponent } from './pages1/mapa-entregas/mapa-entregas';
import { DashexcComponent } from './dasboardexcel/dashexc';
import { ReporteSemanalComponent } from './reporte-semanal/reporte-semanal.component';

// IMPORTAMOS EL GUARDIA QUE CREAMOS
import { authGuard } from './login/auth.guard';

export const routes: Routes = [
  // Redirección inicial al login
  { path: '', redirectTo: 'login', pathMatch: 'full' },
  
  { path: 'login', component: LoginComponent },
  
  // RUTAS PROTEGIDAS (Solo acceden administrativos)
  { path: 'dashboard', component: dashboard, canActivate: [authGuard] },
  { path: 'conductores', component: ConductoresComponent, canActivate: [authGuard] },  
  { path: 'vehiculos', component: ListaVehiculosComponent, canActivate: [authGuard] },
  { path: 'entregas', component: Entregas, canActivate: [authGuard] },
  { path: 'mantenimiento', component: MantenimientoComponent, canActivate: [authGuard] },
  { path: 'alertas', component: AlertasComponent, canActivate: [authGuard] },
  { path: 'combustible', component: CombustibleComponent, canActivate: [authGuard] },
  { path: 'mapa', component: MapaEntregasComponent, canActivate: [authGuard] },
  { path: 'dash-entregas', component: DashexcComponent, canActivate: [authGuard] },
  { path: 'reporte-semanal', component: ReporteSemanalComponent, canActivate: [authGuard] }
];