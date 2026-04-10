import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

import { GlobalGerentePage } from './global-gerente.page';

const routes: Routes = [
  {
    path: '',
    component: GlobalGerentePage
  }
];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule],
})
export class GlobalGerentePageRoutingModule {}
