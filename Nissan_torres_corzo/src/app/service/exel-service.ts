import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import axios from 'axios';
import { environment } from 'src/environments/environment.prod';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {

  private url = environment.urlapi;
  private flaskUrl = 'http://localhost:5000';

  constructor() { }

  private headers() {
    const token = localStorage.getItem('token');
    return {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    };
  }

  // Peticiones al servidor Flask
  async actualizarDatos(mes: number, anio: number, sucursal: string) {
    return axios.post(`${this.flaskUrl}/actualizar-datos`, { mes, anio, sucursal }, this.headers());
  }

  async diaLimite(dia: number, sucursal: any) {
    return axios.post(`${this.flaskUrl}/ejecutar-reporte`, { dia, sucursal }, this.headers());
  }

  async ExtraerDatosGerente(dia: number, sucursal: any) {
    return axios.post(`${this.flaskUrl}/ejecutar-gerente`, { dia, sucursal }, this.headers());
  }

  async ExtraerDatosApv(dia: string, sucursal: any) {
    return axios.post(`${this.flaskUrl}/ejecutar-apv`, { dia, sucursal }, this.headers());
  }

  async postApvDia(dia: string) {
    return axios.post(`${this.flaskUrl}/ejecutar-reporte`, { dia }, this.headers());
  }

  // Peticiones a Strapi
  async getSucursales() {
    let res = await axios.get(`${this.url}/sucursals?populate[numero_apv]=true`, this.headers());
    return res.data;
  }

  async getBySucursales(id: any) {
    let res = await axios.get(`${this.url}/sucursals?filters[documentId][$eq]=${id}`, this.headers());
    return res.data;
  }

  async getDataGlobalGerente(id: any) {
    return await axios.get(
      `${this.url}/globals?filters[sucursal][documentId][$eq]=${id}&filters[Gerente][$null]=false&pagination[pageSize]=1000`,
      this.headers()
    );
  }

async getDataGlobal(id: any, mes: any) {

  console.log('Mes recibido:', mes); // "04-2026"

  // 🔥 separar mes y año
  const [anioStr, mesStr] = mes.split('-');

  const mesNum = Number(mesStr);
  const anio = Number(anioStr);

  // validar por si acaso
  if (!mesNum || !anio) {
    console.error('Mes inválido:', mes);
    return;
  }

  // ✅ formato correcto
  const fechaInicio = `${anio}-${mesStr}-01`;

  const ultimoDia = new Date(anio, mesNum, 0).getDate();

  const fechaFin = `${anio}-${mesStr}-${String(ultimoDia).padStart(2, '0')}`;

  console.log('Fechas calculadas:', fechaInicio, fechaFin);

const res = await axios.get(
  `${this.url}/globals?filters[sucursal][documentId][$eq]=${id}&filters[Gerente][$null]=true&filters[fecha][$gte]=${fechaInicio}&filters[fecha][$lte]=${fechaFin}&sort=fecha:asc`,
  this.headers()
);

  return res.data;
}

  async getApv(Sucursal: any) {
    let res = await axios.get(
      `${this.url}/apvs?filters[sucursal][documentId][$eq]=${Sucursal}&pagination[pageSize]=1000`,
      this.headers()
    );
    return res.data;
  }

  async getApvSucursal(id: any) {
    return axios.get(
      `${this.url}/numero-apvs?filters[sucursal][documentId][$eq]=${id}&filters[gerente][$null]=true`,
      this.headers()
    );
  }

  async getApvGerente(id: any) {
    return axios.get(
      `${this.url}/numero-apvs?filters[sucursal][documentId][$eq]=${id}&filters[gerente][$notNull]=true`,
      this.headers()
    );
  }

  // Generación de Excel
  async generateExcel() {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Mi Hoja');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Nombre', key: 'name', width: 30 },
      { header: 'Fecha', key: 'date', width: 15 }
    ];

    worksheet.addRow({ id: 1, name: 'Ejemplo 1', date: new Date() });
    worksheet.addRow({ id: 2, name: 'Ejemplo 2', date: new Date() });

    worksheet.getRow(1).font = { bold: true };

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }
}