import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { AlertController, ToastController } from '@ionic/angular';
import { ExcelService } from 'src/app/service/exel-service';
import { Login } from 'src/app/service/login';
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

  formSucursal: any = {
    id: null,
    Sucursal: '',
    Plantilla: '',
    apvs: []
  };

  constructor(
    private router: Router,
    private api: ExcelService,
    private apiSuc: Sucursales,
    private alertCtrl: AlertController,
    private toastcontroller: ToastController,
    private login: Login
  ) { }

  async ngOnInit() {
    await this.actualizarFecha();
    await this.getGlobalSucursales();
  }

  guardarSucursal() {
    if (this.formSucursal.id) {
      // ACTUALIZAR
      this.apiSuc.updateApv(this.formSucursal.apv_id, { Num_apv: this.formSucursal.numero_apv })
        .then(() => {
          return this.apiSuc.putSucursal(this.formSucursal.id, {
            Sucursal: this.formSucursal.Sucursal,
            Plantilla: this.formSucursal.Plantilla
          });
        })
        .then(() => {
          this.presentToast('Sucursal actualizada con éxito', 'success');
          this.getGlobalSucursales();
          this.resetForm();
        })
        .catch(() => this.presentAlert('Error', 'No se pudo actualizar la sucursal'));

    } else {
      // CREAR
      const dataApv = { Num_apv: this.formSucursal.numero_apv };
      this.apiSuc.createApv(dataApv)
        .then((resApv: any) => {
          const apvId = resApv.data.id;
          return this.apiSuc.createSucursal({
            Sucursal: this.formSucursal.Sucursal,
            Plantilla: this.formSucursal.Plantilla,
            numero_apv: { connect: [{ id: apvId }] }
          });
        })
        .then(() => {
          this.presentToast('Sucursal creada con éxito', 'success');
          this.getGlobalSucursales();
          this.resetForm();
        })
        .catch(() => this.presentAlert('Error', 'No se pudo crear la sucursal'));
    }
  }

  async actualizar() {
    if (!this.sucursalSeleccionada) {
      this.presentAlert('Atención', 'Por favor selecciona una sucursal primero');
      return;
    }

    try {
      this.presentToast('Iniciando actualización de datos...', 'primary');
      await this.api.actualizarDatos(this.mes, this.anio, this.sucursalSeleccionada);
      this.presentAlert('Éxito', 'El proceso de actualización para ' + this.sucursalSeleccionada + ' ha finalizado.');
    } catch (err) {
      this.presentAlert('Error', 'Falló la conexión con el script de Excel');
    }
  }

  async eliminarSucursal(id: number) {
    const alert = await this.alertCtrl.create({
      header: 'Confirmar eliminación',
      message: '¿Estás seguro de que deseas eliminar esta sucursal?',
      mode: 'ios',
      buttons: [
        { text: 'Cancelar', role: 'cancel' },
        {
          text: 'Eliminar',
          handler: () => {
            this.apiSuc.delSucursal(id).then(() => {
              this.presentToast('Sucursal eliminada', 'danger');
              this.getGlobalSucursales();
            });
          }
        }
      ]
    });
    await alert.present();
  }

  // --- Helpers de UI ---
  async presentAlert(header: string, msg: string) {
    const alert = await this.alertCtrl.create({
      header: header,
      message: msg,
      buttons: ['OK'],
      mode: 'ios'
    });
    await alert.present();
  }

  async presentToast(message: string, color: string = 'dark') {
    const toast = await this.toastcontroller.create({
      message,
      duration: 2500,
      position: 'bottom',
      color,
      mode: 'ios'
    });
    await toast.present();
  }

  modalActualizar: boolean = false;

  async ejecutarActualizacion() {
    if (!this.sucursalSeleccionada) {
      this.presentAlert('Atención', 'Debes seleccionar una sucursal primero');
      return;
    }
    this.modalActualizar = false;
    await this.actualizar();
  }

  actualizarFecha() {
    const fecha = new Date(this.fechaSeleccionada);
    this.mes = fecha.getMonth() + 1;
    this.anio = fecha.getFullYear();
  }

  editarSucursal(s: any) {
    this.formSucursal = {
      id: s.documentId,
      Sucursal: s.Sucursal,
      Plantilla: s.Plantilla,
      numero_apv: s.numero_apv?.[0]?.Num_apv,
      apv_id: s.numero_apv?.[0]?.documentId
    };
    this.presentToast('Editando sucursal: ' + s.Sucursal);
  }

  resetForm() {
    this.formSucursal = { id: null, Sucursal: '', Plantilla: '', apvs: [] };
  }

  async getGlobalSucursales() {
    try {
      const data = await this.api.getSucursales();
      this.sucursales = data.data;
    } catch (error) {
      this.presentToast('Error al cargar sucursales', 'danger');
    }
  }

  logout() {
    try {
      this.login.logout();
      this.presentToast('Sesión cerrada', 'success');
      this.router.navigate(['/login']);
    } catch (error) {
      this.presentToast('Error al cerrar sesión', 'danger');
    }
  }

  globalSucursales(sucursal: any) { this.router.navigate(['/global-sucursal/', sucursal]); }
  globalSucursalesGerente(sucursal: any) { this.router.navigate(['/global-gerente/', sucursal]); }
  apvSucursales(sucursal: any) { this.router.navigate(['/apv/', sucursal]); }
}
