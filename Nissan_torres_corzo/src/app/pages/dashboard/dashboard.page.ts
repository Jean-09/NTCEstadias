import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { ExcelService } from 'src/app/service/exel-service';
import { Sucursales } from 'src/app/service/sucursales';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.page.html',
  styleUrls: ['./dashboard.page.scss'],
  standalone: false
})
export class DashboardPage implements OnInit {

  sucursales: any[] = [];

  mes!: number;
  anio!: number;
  sucursalSeleccionada: string = '';
  fechaSeleccionada: string = new Date().toISOString();

  constructor(private router: Router, private api: ExcelService, private apiSuc: Sucursales) { }

  async ngOnInit() {
    await this.actualizarFecha();
    await this.getGlobalSucursales();
  }

  formSucursal: any = {
    id: null,
    Sucursal: '',
    Plantilla: '',
    apvs: []
  };

  sucursal: any[] = [];
  listaApv: any[] = [];

  guardarSucursal() {

    if (this.formSucursal.id) {

      // 1. ACTUALIZAR APV
      this.apiSuc.updateApv(this.formSucursal.apv_id, {
        Num_apv: this.formSucursal.numero_apv
      })
        .then(() => {

          // 2. ACTUALIZAR SUCURSAL
          return this.apiSuc.putSucursal(this.formSucursal.id, {
            Sucursal: this.formSucursal.Sucursal,
            Plantilla: this.formSucursal.Plantilla
          });

        })
        .then(() => {
          this.getGlobalSucursales();
          this.resetForm();
        })
        .catch((err: any) => console.error(err));

    } else {

      // CREATE (el de arriba)
      const dataApv = {
        Num_apv: this.formSucursal.numero_apv
      };

      this.apiSuc.createApv(dataApv)
        .then((resApv: any) => {

          const apvId = resApv.data.id;

          return this.apiSuc.createSucursal({
            Sucursal: this.formSucursal.Sucursal,
            Plantilla: this.formSucursal.Plantilla,
            numero_apv: {
              connect: [{ id: apvId }]
            }
          });
        })
        .then(() => {
          this.getGlobalSucursales();
          this.resetForm();
        })
        .catch(err => console.error(err));
    }
  }

  actualizarFecha() {
    const fecha = new Date(this.fechaSeleccionada);
    this.mes = fecha.getMonth() + 1;
    this.anio = fecha.getFullYear();
  }

  async actualizar() {
    const token = localStorage.getItem('token') || '';

    if (!this.sucursalSeleccionada) {
      alert("Por favor selecciona una sucursal");
      return;
    }

    try {
      await this.api.actualizarDatos(
        this.mes,
        this.anio,
        this.sucursalSeleccionada
      );
      alert("Proceso iniciado para " + this.sucursalSeleccionada);
    } catch (err) {
    }
  }

  eliminarSucursal(id: number) {
    this.apiSuc.delSucursal(id).then(() => {
      this.getGlobalSucursales();
    });
  }

  editarSucursal(s: any) {
    this.formSucursal = {
      id: s.documentId,
      Sucursal: s.Sucursal,
      Plantilla: s.Plantilla,
      numero_apv: s.numero_apv?.Num_apv,
      apv_id: s.numero_apv?.documentId
    };
  }

  resetForm() {
    this.formSucursal = {
      id: null,
      Sucursal: '',
      Plantilla: '',
      apvs: []
    };
  }

  async getGlobalSucursales() {
    try {
      const data = await this.api.getSucursales();
      this.sucursales = data.data;
    } catch (error) {
      console.error('Error al obtener sucursales:', error);
    }
  }

  globalSucursales(sucursal: any) {
    this.router.navigate(['/global-sucursal/', sucursal]);
  }
  globalSucursalesGerente(sucursal: any) {
    this.router.navigate(['/global-gerente/', sucursal]);
  }

  apvSucursales(sucursal: any) {
    this.router.navigate(['/apv/', sucursal]);
  }

}
