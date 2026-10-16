import { Routes } from '@angular/router';
import { MsalGuard } from '@azure/msal-angular';
import { Inicio } from './inicio/inicio';
import { Pedidos } from './pedidos/pedidos';

export const routes: Routes = [
  { path: '', component: Inicio },
  // ruta protegida: si no hay sesion, el guard abre el login
  { path: 'pedidos', component: Pedidos, canActivate: [MsalGuard] },
  { path: '**', redirectTo: '' }
];
