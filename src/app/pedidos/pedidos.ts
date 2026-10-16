import { Component, OnInit, inject } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { API_SCOPE } from '../config';
import { Pedido, PedidosService } from '../services/pedidos.service';

@Component({
  selector: 'app-pedidos',
  standalone: true,
  templateUrl: './pedidos.html'
})
export class Pedidos implements OnInit {
  private pedidosService = inject(PedidosService);
  private auth = inject(MsalService);

  pedidos: Pedido[] = [];
  error = '';
  roles: string[] = [];
  scopes: string[] = [];

  get esAdmin(): boolean {
    return this.roles.includes('Admin');
  }

  ngOnInit(): void {
    const instance = this.auth.instance;
    const cuenta = instance.getActiveAccount() ?? instance.getAllAccounts()[0];

    // roles: vienen en los claims del token del usuario
    const claims = cuenta?.idTokenClaims as { roles?: string[] } | undefined;
    this.roles = claims?.roles ?? [];

    // scopes: vienen en el access token que se usa para llamar a la API
    if (cuenta) {
      this.auth.acquireTokenSilent({ scopes: [API_SCOPE], account: cuenta }).subscribe({
        next: (r) => (this.scopes = r.scopes.map((s) => s.split('/').pop() ?? s)),
        error: () => (this.scopes = [])
      });
    }

    this.cargar();
  }

  cargar() {
    this.pedidosService.listar().subscribe({
      next: (data) => { this.pedidos = data; this.error = ''; },
      error: (err) => (this.error = `Error ${err.status} al cargar los pedidos`)
    });
  }

  crear(nombre: string, email: string, descripcion: string, total: string) {
    this.pedidosService
      .crear({ cliente: { nombre, email }, descripcion, total: Number(total) })
      .subscribe({
        next: () => this.cargar(),
        error: (err) => (this.error = `Error ${err.status} al crear el pedido`)
      });
  }

  cambiarEstado(p: Pedido, estado: string) {
    if (p.id == null) return;
    this.pedidosService.cambiarEstado(p.id, estado).subscribe({
      next: () => this.cargar(),
      error: (err) => (this.error = `Error ${err.status} al cambiar el estado`)
    });
  }

  eliminar(p: Pedido) {
    if (p.id == null) return;
    this.pedidosService.eliminar(p.id).subscribe({
      next: () => this.cargar(),
      error: (err) => (this.error = err.status === 403 ? 'No tienes permiso (solo Admin)' : `Error ${err.status}`)
    });
  }
}
