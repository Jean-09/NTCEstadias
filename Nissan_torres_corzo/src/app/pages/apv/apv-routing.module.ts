import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

import { APVPage } from './apv.page';

const routes: Routes = [
  {
    path: '',
    component: APVPage
  }
];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule],
})
export class APVPageRoutingModule {}
