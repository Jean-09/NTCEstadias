import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import axios, { AxiosHeaders } from 'axios';
import { environment } from 'src/environments/environment.prod';

@Injectable({
  providedIn: 'root',
})
export class ExcelService {

  private url = environment.urlapi
  constructor() { }

  private headers() {
    const token = localStorage.getItem('token')
    console.log(token)
    return {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    }
  }

  async getSucursales() {
    let res = await axios.get(`${this.url}/sucursals?populate[numero_apv]=true`, this.headers());
    return res.data;
  }

  async getBySucursales(id: any) {
    let res = await axios.get(`${this.url}/sucursals?filters[documentId][$eq]=${id}`, this.headers());
    return res.data;
  }

  async getDataGlobalGerente(id: any) {
    // Se cambia $null a false para traer los que NO están vacíos
    let res = await axios.get(
        `${this.url}/globals?filters[sucursal][documentId][$eq]=${id}&filters[Gerente][$null]=false&pagination[pageSize]=1000`, 
        this.headers()
    );
    return res;
}

  async generateExcel() {
    // 1. Crear libro y hoja
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Mi Hoja');

    // 2. Definir columnas
    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Nombre', key: 'name', width: 30 },
      { header: 'Fecha', key: 'date', width: 15 }
    ];

    // 3. Agregar filas (datos)
    worksheet.addRow({ id: 1, name: 'Ejemplo 1', date: new Date() });
    worksheet.addRow({ id: 2, name: 'Ejemplo 2', date: new Date() });

    // 4. Estilos (Ejemplo: negrita en el encabezado)
    worksheet.getRow(1).font = { bold: true };

    // 5. Generar archivo y guardar (Web)
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  async getDataGlobal(id: any) {
    // Se añade el filtro $null para el campo Gerente
    let res = await axios.get(
      `${this.url}/globals?filters[sucursal][documentId][$eq]=${id}&filters[Gerente][$null]=true&pagination[pageSize]=1000`,
      this.headers()
    );
    return res;
  }

  async getApv(Sucursal: any) {
    let res = await axios.get(`${this.url}/apvs?filters[sucursal][documentId][$eq]=${Sucursal}&pagination[pageSize]=1000`, this.headers());
    return res.data;
  }

  async diaLimite(dia: string, sucursal: any) {
    const res = await axios.post('http://localhost:5000/ejecutar-reporte', { dia: dia, sucursal: sucursal }, this.headers());
    console.log('Respuesta del servidor:', res.data);
    return res;
  }

    async ExtraerDatosGerente(dia: string, sucursal: any) {
    const res = await axios.post('http://localhost:5000/ejecutar-gerente', { dia: dia, sucursal: sucursal }, this.headers());
    console.log('Respuesta del servidor:', res.data);
    return res;
  }

  async ExtraerDatosApv(dia: string, sucursal: any) {
    const res = await axios.post('http://localhost:5000/ejecutar-apv', { dia: dia, sucursal: sucursal }, this.headers());
    console.log('Respuesta del servidor:', res.data);
    return res;
  }


  async postApvDia(dia: string) {
    return await axios.post('http://localhost:5000/ejecutar-reporte', { dia: dia }, this.headers());
  }

  async getApvSucursal(id: any) {
    return await axios.get(
      `${this.url}/numero-apvs?filters[sucursal][documentId][$eq]=${id}`,
      this.headers()
    );
  }
}
