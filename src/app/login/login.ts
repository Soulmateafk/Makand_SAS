import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { AuthService } from './auth.service';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './login.html',
  styleUrls: ['./login.css']
})
export class LoginComponent {
  loginForm: FormGroup;
  isLoading = false;
  errorMessage = '';

  constructor(
    private fb: FormBuilder,
    private router: Router,
    private authService: AuthService // <--- INYECTAMOS EL SERVICIO
  ) {
    this.loginForm = this.fb.group({
      email: ['', [Validators.required, Validators.email]],
      password: ['', [Validators.required, Validators.minLength(6)]],
      rememberMe: [false]
    });
  }

  onLogin() {
    if (this.loginForm.valid) {
      this.isLoading = true;
      this.errorMessage = '';

      // LLAMADA REAL A TU SERVIDOR NODE.JS / MONGODB
      this.authService.login(this.loginForm.value).subscribe({
        next: (res) => {
          // 1. Guardamos los datos básicos del usuario que ya traías
          localStorage.setItem('user', JSON.stringify(res.user));

          // 🔑 2. GUARDAMOS EL TOKEN JWT SEGURO QUE VIENE DEL BACKEND
          // Si el servidor envía 'res.token', lo almacenamos para usarlo en las peticiones HTTP
          if (res.token) {
          localStorage.setItem('token_logistica', res.token);          }

          // 3. Redireccionamos al dashboard
          this.router.navigate(['/dashboard']);
        },
        error: (err) => {
          // Si el servidor responde 401 (Credenciales inválidas)
          this.errorMessage = 'Credenciales corporativas inválidas.';
          this.isLoading = false;
        }
      });

    } else {
      this.loginForm.markAllAsTouched();
    }
  }
}