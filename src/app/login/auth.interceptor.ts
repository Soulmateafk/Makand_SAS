import { HttpInterceptorFn } from '@angular/common/http';
import { inject } from '@angular/core';
import { PLATFORM_ID } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';

export const authInterceptor: HttpInterceptorFn = (req, next) => {
  const platformId = inject(PLATFORM_ID);
  let token = null;

  // Verificamos de forma segura si estamos en el navegador para leer el localStorage
  if (isPlatformBrowser(platformId)) {
    token = localStorage.getItem('token_logistica'); // Usa la clave exacta con la que guardas tu token en el login
  }

  // Si encontramos el token, clonamos la petición y le inyectamos la cabecera Authorization
  if (token) {
    const clonedReq = req.clone({
      setHeaders: {
        Authorization: `Bearer ${token}`
      }
    });
    return next(clonedReq);
  }

  // Si no hay token (como en el login o registro), la petición sigue su curso normal
  return next(req);
};