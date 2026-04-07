import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.page.html',
  styleUrls: ['./dashboard.page.scss'],
  standalone: false
})
export class DashboardPage implements OnInit {

  constructor(private router: Router) { }

  ngOnInit() {

  }

  globalSucursales(sucursal:any) {
    this.router.navigate(['/global-sucursal/', sucursal]);
  }

  apv() {
    this.router.navigate(['/apv']);
  }

}
