import { Component, Inject, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { MsalService, MsalBroadcastService, MSAL_GUARD_CONFIG, MsalGuardConfiguration } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { Subject, filter, takeUntil } from 'rxjs';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App implements OnInit, OnDestroy {
  title = 'Pedidos360';
  loginDisplay = false;
  apiResponse = '';
  private readonly _destroying$ = new Subject<void>();

  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private http: HttpClient
  ) {}

  ngOnInit(): void {
    // 1. Inicializamos MSAL antes de hacer cualquier otra cosa
    this.authService.initialize().subscribe(() => {
      this.verificarCuentaActiva();
    });

    // 2. Escuchamos los cambios de estado (cuando termina de cargar un popup)
    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe(() => {
        this.verificarCuentaActiva();
      });
  }

  verificarCuentaActiva() {
    let cuentaActiva = this.authService.instance.getActiveAccount();

    if (!cuentaActiva && this.authService.instance.getAllAccounts().length > 0) {
      let cuentasGuardadas = this.authService.instance.getAllAccounts();
      this.authService.instance.setActiveAccount(cuentasGuardadas[0]);
      cuentaActiva = cuentasGuardadas[0];
    }

    this.loginDisplay = !!cuentaActiva;
  }

  login() {
    this.authService.loginPopup().subscribe({
      next: (result) => {
        if (result && result.account) {
          this.authService.instance.setActiveAccount(result.account);
          this.verificarCuentaActiva();
        }
      },
      error: (error) => console.error('Error en el login Popup:', error)
    });
  }

  llamarBackend() {
    const cuentaActiva = this.authService.instance.getActiveAccount();
    if (!cuentaActiva) {
      this.apiResponse = 'Error: No hay cuenta activa.';
      return;
    }

    // Le pedimos a Microsoft el token actual para nuestra sesión
    this.authService.acquireTokenSilent({
      account: cuentaActiva,
      scopes: ["70a98cd7-3e6e-4e11-a209-d0ae414fb4db/.default"]
    }).subscribe({
      next: (resultado) => {
        // Adjuntamos el token de identidad como un pase VIP en la cabecera
        const headers = new HttpHeaders().set('Authorization', 'Bearer ' + resultado.idToken);

        this.http.get<{mensaje: string}>('http://localhost:8080/api/test', { headers })
          .subscribe({
            next: (res) => this.apiResponse = res.mensaje,
            error: (err) => {
              console.error('Error en la llamada al backend:', err);
              this.apiResponse = 'Error: No se pudo conectar o el token fue rechazado por el backend.';
            }
          });
      },
      error: (err) => {
        console.error('Error obteniendo el token:', err);
        this.apiResponse = 'Error al obtener el pase de seguridad de Microsoft.';
      }
    });
  }

  logout() {
    this.authService.logoutPopup().subscribe(() => {
      this.authService.instance.setActiveAccount(null);
      this.verificarCuentaActiva();
    });
  }

  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
  }
}