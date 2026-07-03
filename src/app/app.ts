import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet, Router, NavigationEnd, RouterLink } from '@angular/router';
import { filter } from 'rxjs/operators';
import { AuthService } from './login/auth.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet, RouterLink],
  templateUrl: './app.html',
  styleUrls: ['./app.css']
})
export class AppComponent {
  showNavbar: boolean = false;
  title = 'logistica-app';

  constructor(private router: Router, public authService: AuthService) {
    this.router.events.pipe(
      filter(event => event instanceof NavigationEnd)
    ).subscribe((event: any) => {
      const currentUrl = event.urlAfterRedirects || event.url;
      this.showNavbar = !currentUrl.includes('/login');
    });
  }

  // --- FUNCIÓN PARA CERRAR SESIÓN ---
  cerrarSesion() {
    this.authService.logout(); // Ejecuta el localStorage.clear()
    this.router.navigate(['/login']); // Redirige al login
  }
}