import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

import { GlobalSucursalPage } from './global-sucursal.page';

const routes: Routes = [
  {
    path: '',
    component: GlobalSucursalPage
  }
];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule],
})
export class GlobalSucursalPageRoutingModule {}
