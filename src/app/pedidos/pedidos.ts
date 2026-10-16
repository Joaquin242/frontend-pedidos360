import { Component, OnInit, inject } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
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

  get esAdmin(): boolean {
    return this.roles.includes('Admin');
  }

  ngOnInit(): void {
    // roles que vienen en los claims del token
    const claims = this.auth.instance.getActiveAccount()?.idTokenClaims as { roles?: string[] } | undefined;
    this.roles = claims?.roles ?? [];
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
