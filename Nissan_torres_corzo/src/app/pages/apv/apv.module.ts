import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

import { IonicModule } from '@ionic/angular';

import { APVPageRoutingModule } from './apv-routing.module';

import { APVPage } from './apv.page';

@NgModule({
  imports: [
    CommonModule,
    FormsModule,
    IonicModule,
    APVPageRoutingModule
  ],
  declarations: [APVPage]
})
export class APVPageModule {}
