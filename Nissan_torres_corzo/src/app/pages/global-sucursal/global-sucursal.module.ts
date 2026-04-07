import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

import { IonicModule } from '@ionic/angular';

import { GlobalSucursalPageRoutingModule } from './global-sucursal-routing.module';

import { GlobalSucursalPage } from './global-sucursal.page';

@NgModule({
  imports: [
    CommonModule,
    FormsModule,
    IonicModule,
    GlobalSucursalPageRoutingModule
  ],
  declarations: [GlobalSucursalPage]
})
export class GlobalSucursalPageModule {}
