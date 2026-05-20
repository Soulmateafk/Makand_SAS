import { Injectable, Inject, PLATFORM_ID } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { tap } from 'rxjs/operators';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  private apiUrl = 'https://makand-sas.onrender.com/api';

  constructor(
    private http: HttpClient,
    @Inject(PLATFORM_ID) private platformId: Object
  ) { }

  // 🔑 LOGIN: Guarda el token y el rol al recibir la respuesta
  login(credentials: any): Observable<any> {
    return this.http.post(`${this.apiUrl}/login`, credentials).pipe(
      tap((res: any) => {
        if (res.token && isPlatformBrowser(this.platformId)) {
          localStorage.setItem('token_logistica', res.token);
          localStorage.setItem('user_role', res.user.rol);
          localStorage.setItem('user_name', res.user.nombre);
        }
      })
    );
  }

  // MÉTODO PARA SABER EL ROL EN CUALQUIER PARTE DE LA APP
  getRol(): string | null {
    if (isPlatformBrowser(this.platformId)) {
      return localStorage.getItem('user_role');
    }
    return null;
  }

  // MÉTODO PARA VERIFICAR SI ES ADMINISTRATIVO
  esAdministrativo(): boolean {
    return this.getRol() === 'administrativo';
  }

  // CERRAR SESIÓN
  logout() {
    if (isPlatformBrowser(this.platformId)) {
      localStorage.clear();
    }
  }

  obtenerEntregas(): Observable<any[]> {
    return this.http.get<any[]>(`${this.apiUrl}/entregas`);
  }
}