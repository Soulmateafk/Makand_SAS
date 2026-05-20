import { inject } from '@angular/core';
import { Router, CanActivateFn } from '@angular/router';
import { AuthService } from './auth.service';

export const authGuard: CanActivateFn = (route, state) => {
  const authService = inject(AuthService);
  const router = inject(Router);

  // Verificamos si el usuario es administrativo
  if (authService.esAdministrativo()) {
    return true; // ¡Pasa!
  } else {
    // Si no es administrativo, lo echamos al login
    router.navigate(['/login']);
    return false; // ¡Bloqueado!
  }
};