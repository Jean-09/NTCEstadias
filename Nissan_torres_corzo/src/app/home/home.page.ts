import { Component, OnInit } from '@angular/core';
import { ExcelService } from '../service/exel-service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-home',
  templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
  standalone: false,
})
export class HomePage implements OnInit {
constructor(private excelService: ExcelService) {}

  ngOnInit() {
  
  }
}
