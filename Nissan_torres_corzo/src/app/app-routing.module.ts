import { NgModule } from '@angular/core';
import { PreloadAllModules, RouterModule, Routes } from '@angular/router';
import { guardGuard } from './guard/guard-guard';

const routes: Routes = [
  {
    path: 'home',
    loadChildren: () => import('./home/home.module').then( m => m.HomePageModule)
  },
  {
    path: '',
    redirectTo: 'login',
    pathMatch: 'full'
  },
  {
    path: 'login',
    loadChildren: () => import('./pages/login/login.module').then( m => m.LoginPageModule)
  },
  {
    path: 'apv/:sucursal',
    loadChildren: () => import('./pages/apv/apv.module').then( m => m.APVPageModule), canActivate:[guardGuard]
  },
  {
    path: 'dashboard',
    loadChildren: () => import('./pages/dashboard/dashboard.module').then( m => m.DashboardPageModule), canActivate:[guardGuard]
  },
  {
    path: 'global-sucursal/:sucursal',
    loadChildren: () => import('./pages/global-sucursal/global-sucursal.module').then( m => m.GlobalSucursalPageModule), canActivate:[guardGuard]
  },
  {
    path: 'global-gerente/:sucursal',
    loadChildren: () => import('./pages/global-gerente/global-gerente.module').then( m => m.GlobalGerentePageModule), canActivate:[guardGuard]
  },

];

@NgModule({
  imports: [
    RouterModule.forRoot(routes, { preloadingStrategy: PreloadAllModules })
  ],
  exports: [RouterModule]
})
export class AppRoutingModule { }
