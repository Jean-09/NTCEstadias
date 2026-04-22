import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import axios, { AxiosHeaders } from 'axios';
import { environment } from 'src/environments/environment.prod';


@Injectable({
  providedIn: 'root',
})
export class Sucursales {

  private url = environment.urlapi
  constructor() { }

  private headers() {
    const token = localStorage.getItem('token')

    return {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    }
  }

  async putSucursal(id: any, data: any) {
    let res = await axios.put(`${this.url}/sucursals/${id}`, { data }, this.headers());
    return res.data;
  }

  async createSucursal(data: any) {
    let res = await axios.post(`${this.url}/sucursals`, { data }, this.headers());
    return res.data;
  }

  async delSucursal(id: any) {
    let res = await axios.delete(`${this.url}/sucursals/${id}`, this.headers());
    return res.data;
  }

  async createApv(data: any) {
    let res = await axios.post(`${this.url}/numero-apvs`, { data }, this.headers());
    return res.data;
  }

  async updateApv(id: any, data: any) {
    console.log(data)
    let res = await axios.put(`${this.url}/numero-apvs/${id}`, { data }, this.headers());
    console.log(res) 
    return res.data;
  }

  async delApv(id: any) {
    let res = await axios.delete(`${this.url}/numero-apvs/${id}`, this.headers());
    return res.data;
  }

}
