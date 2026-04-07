import { Component } from '@angular/core';
import { ExcelService } from './service/exel-service';

@Component({
  selector: 'app-root',
  templateUrl: 'app.component.html',
  styleUrls: ['app.component.scss'],
  standalone: false,
})

export class HomePage {
  constructor(private excelService: ExcelService) { }

  exportar() {
    this.excelService.generateExcel();
  }
}

