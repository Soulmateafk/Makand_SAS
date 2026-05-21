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
      this.isLoading = true; // Empieza a cargar
      this.errorMessage = '';

      this.authService.login(this.loginForm.value).subscribe({
        next: (res) => {
          // --- AQUÍ ESTABA EL PROBLEMA ---
          this.isLoading = false; // <--- ¡DEBES APAGAR EL CARGADOR AQUÍ!
          // -------------------------------

          // 1. Guardamos los datos
          localStorage.setItem('user', JSON.stringify(res.user));

          // 2. Guardamos el token
          if (res.token) {
             localStorage.setItem('token_logistica', res.token);
          }

          // 3. Redireccionamos
          this.router.navigate(['/dashboard']);
        },
        error: (err) => {
          this.errorMessage = 'Credenciales corporativas inválidas.';
          this.isLoading = false; // Aquí sí lo tenías bien
        }
      });
    } else {
      this.loginForm.markAllAsTouched();
    }
  }
}