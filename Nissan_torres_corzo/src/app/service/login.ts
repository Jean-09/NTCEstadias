import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import axios, { AxiosHeaders } from 'axios';
import { Router } from '@angular/router';
import { environment } from 'src/environments/environment.prod';

@Injectable({
  providedIn: 'root',
})
export class Login {

   private url = environment.urlapi

  

  constructor(private route: Router) {
  }

   async login(data: any) {
    const res = await axios.post(this.url + '/auth/local', data);
    return res.data;
  }

  logout(){
    localStorage.removeItem('token');
    this.route.navigate(['/login']);
  }
}
