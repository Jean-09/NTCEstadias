import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

import { IonicModule } from '@ionic/angular';

import { GlobalGerentePageRoutingModule } from './global-gerente-routing.module';

import { GlobalGerentePage } from './global-gerente.page';

@NgModule({
  imports: [
    CommonModule,
    FormsModule,
    IonicModule,
    GlobalGerentePageRoutingModule
  ],
  declarations: [GlobalGerentePage]
})
export class GlobalGerentePageModule {}
