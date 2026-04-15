import { Component, OnInit } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { AlertController, LoadingController, ToastController } from '@ionic/angular';
import { ExcelService } from 'src/app/service/exel-service';
import { Login } from 'src/app/service/login';

@Component({
  selector: 'app-login',
  templateUrl: './login.page.html',
  styleUrls: ['./login.page.scss'],
  standalone: false
})
export class LoginPage implements OnInit {
  accesToken: string | null = null;

  constructor(private api: Login, private route: Router, private alertController: AlertController, private toastcontroller: ToastController, private act: ActivatedRoute,) {
    this.accesToken = act.snapshot.queryParams['access_token'] as string;
  }

  identifier: string = "Admin";
  password: string = "12345678";

  async ngOnInit() {
  }

  ionViewWillEnter() {

  }
  async login() {
    if (!this.identifier || !this.password) {
      this.presentAlert('Campos incompletos', 'Por favor, ingresa tu nombre de usuario y contraseña.');
      return;
    }
    const data = {
      identifier: this.identifier,
      password: this.password
    }

    this.api.login(data).then((res: any) => {
      this.saveToken(res)
      console.log(res)
      
    }).catch((error: any) => {

      if (error.code === 'ERR_BAD_REQUEST') {
        this.presentAlert('Verifica los datos', 'Verifica tus datos');
        return
      }
      if (error.code === 'ERR_NETWORK') {
        this.presentAlert('Error del servidor', 'No se puede conectar al servidor intentalo más tarde');
        return
      }
    })
  }


  async presentAlert(header: string, msg: string) {
    const alert = await this.alertController.create({
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
      position: 'middle',
      color,
    });
    toast.present();
  }

  async saveToken(data: any) {
    try {
      localStorage.setItem('token', data.jwt)
      this.presentAlert('Bienvenido', "NISSAN TORRES CORZO");
      setTimeout(() =>{
        this.route.navigateByUrl('dashboard')
      },2000)
    } catch (error) {
      this.presentAlert('Error', 'No se pudo guardar los datos de sesión')
    }
  }
}
