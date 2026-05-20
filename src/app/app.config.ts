import { ApplicationConfig } from '@angular/core';
import { provideRouter } from '@angular/router';
import { routes } from './app.routes';
// 1. Añadimos 'withFetch' a los imports
import { provideHttpClient, withFetch, withInterceptors } from '@angular/common/http';
// 🔑 Tu interceptor se mantiene igual
import { authInterceptor } from './login/auth.interceptor';

export const appConfig: ApplicationConfig = {
  providers: [
    // 2. Combinamos ambos: withFetch() y withInterceptors() dentro de provideHttpClient
    provideHttpClient(
      withFetch(), 
      withInterceptors([authInterceptor])
    ),
    provideRouter(routes)
  ]
};