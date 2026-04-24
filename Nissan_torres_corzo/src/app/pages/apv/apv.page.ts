import { Component, OnInit } from '@angular/core';
import { ExcelService } from 'src/app/service/exel-service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { ActivatedRoute, Router } from '@angular/router';
import { AlertController, ToastController } from '@ionic/angular';
import { Login } from 'src/app/service/login';

@Component({
  selector: 'app-apv',
  templateUrl: './apv.page.html',
  styleUrls: ['./apv.page.scss'],
  standalone: false
})
export class APVPage implements OnInit {

  paginas: any = {};
  gerenteSeleccionado: string = '';
  diaLimite = '';
  id: string = '';
  Object = Object;
  sucursal: any[] = [];

  apvData: any[] = [];
  conceptos = [
    { nombre: 'CITAS ASISTIDAS', key: 'citas_asistidas' },
    { nombre: 'SOLICITUDES C/DATOS COMP.', key: 'solicitudes_cdatos' },
    { nombre: 'DOC. COMPLETA (MES)', key: 'doc_compl_mes' },
    { nombre: 'AUTORIZADAS', key: 'autorizadas' },
    { nombre: 'PEDIDOS C/ANTICIPO', key: 'pedidos_canticipo' },
    { nombre: 'FACTURADAS', key: 'facturadas' },
    { nombre: 'DESEMBOLSADAS', key: 'desenbolsadas' },
    { nombre: 'ENTREGAS', key: 'entregas' }
  ];
  constructor(private api: ExcelService, private act: ActivatedRoute, private alertCtrl: AlertController, private toastCtrl: ToastController, private login: Login, private router: Router) {
    this.id = this.act.snapshot.paramMap.get('sucursal') as string;
  }

  async ngOnInit() {
    await this.getSucursal();
    await this.getApv();
  }

  async getSucursal() {
    try {
      const res = await this.api.getBySucursales(this.id);
      if (res.data && res.data.length > 0) {
        this.sucursal = res.data[0].Sucursal;
      } else {
        await this.presentAlert('Error', 'No se encontró la información de la sucursal.');
      }
    } catch (error) {
      await this.presentAlert('Error de Conexión', 'No se pudo obtener la sucursal desde el servidor.');
    }
  }

  getApv() {
    this.api.getApv(this.id).then((res: any) => {
      this.apvData = res.data;
      if (this.apvData.length === 0) {
        this.presentToast('No hay registros de APV para esta sucursal', 'primary');
      }
      this.filtrarPorGerente(this.apvData);
    }).catch((error: any) => {
      this.presentAlert('Error', 'Fallo al cargar los datos de los vendedores.');
    });
  }

  async dispararAutomatizacion() {
    // Validación de campo vacío
    if (!this.diaLimite) {
      await this.presentAlert('Campo Requerido', 'Por favor, selecciona una fecha límite antes de continuar.');
      return;
    }

    try {
      await this.presentToast('Iniciando extracción de datos...', 'primary');
      await this.api.ExtraerDatosApv(this.diaLimite, this.sucursal);
      await this.getApv();
      await this.presentToast('Datos actualizados correctamente', 'success');
    } catch (error) {
      console.error('Error en el servicio Nissan:', error);
      await this.presentAlert('Error de Sincronización', 'No se pudieron extraer los datos de Nissan. Intenta de nuevo.');
    }
  }

  // Agrupa datos por gerente
  filtrarPorGerente(data: any[]) {
    const global = data.filter(x => x.tipo_registro === 'GLOBAL');

    const gerentes = [
      ...new Set(
        data
          .filter(x => x.tipo_registro === 'GERENTE')
          .map(x => x.Gerente)
      )
    ];

    const resultado: any = {};

    gerentes.forEach(nombreGerente => {
      const datosGerente = data.filter(x => x.Gerente === nombreGerente);
      resultado[nombreGerente] = datosGerente;
    });

    this.filtrarPorFechas(resultado, global);
  }

  getConversionD(conceptoIndex: number, item: any): string {
    if (!item || !item.Maduracion) return '0%';

    const m = item.Maduracion;
    let resultado = 0;

    try {
      switch (conceptoIndex) {
        case 2: // DOC COMPL MES
          resultado = (m.doc_compl_mes || 0) / (m.solicitudes_cdatos || 1);
          break;
        case 3: // AUTORIZADAS
          resultado = (m.autorizadas || 0) / (m.doc_compl_mes || 1);
          break;
        case 4: // PEDIDOS C/ANTICIPO
          resultado = (m.pedidos_canticipo || 0) / (m.autorizadas || 1);
          break;
        case 6: // DESEMBOLSADAS
          resultado = (m.desenbolsadas || 0) / (m.pedidos_canticipo || 1);
          break;
        default:
          return '';
      }

      if (!isFinite(resultado) || isNaN(resultado)) return '0%';
      return Math.round(resultado * 100) + '%';
    } catch (e) {
      return '0%';
    }
  }

  semaforos: any = {};

  procesarSemaforos(gerente: string) {
    this.semaforos = {};
    const columnas = this.getColumnas(this.paginas[gerente]);

    if (!columnas || columnas.length === 0) return;

    columnas.forEach((col: any, iCol: number) => {
      if (iCol === 0) return;

      const colAnterior = columnas[iCol - 1];

      this.compararBloque(col.global, colAnterior.global, `global-${iCol}`, col, colAnterior);

      this.compararBloque(col.gerente, colAnterior.gerente, `gerente-${iCol}`, col, colAnterior);

      if (col.vendedores && colAnterior.vendedores) {
        col.vendedores.forEach((v: any, iV: number) => {
          const vAnterior = colAnterior.vendedores[iV];
          if (vAnterior) {
            this.compararBloque(v, vAnterior, `vendedor-${iV}-${iCol}`, col, colAnterior);
          }
        });
      }
    });
  }
  compararBloque(actual: any, anterior: any, id: string, colAct: any, colAnt: any) {
    const numApv = id.includes('global') ? 22 :
      id.includes('gerente') ? colAct.vendedores.length : 1;

    this.conceptos.forEach((c: any, index: number) => {
      const valAct = Math.trunc(this.getPorcentaje(actual?.Mes?.[c.key], index, colAct, numApv));
      const valAnt = Math.trunc(parseFloat(this.getPorcentaje(anterior?.Mes?.[c.key], index, colAnt, numApv).toString()) || 0);

      const valMadAct = Math.trunc(parseFloat(this.getPorcentaje(actual?.Maduracion?.[c.key], index, colAct, numApv).toString()) || 0);
      const valMadAnt = Math.trunc(parseFloat(this.getPorcentaje(anterior?.Maduracion?.[c.key], index, colAnt, numApv).toString()) || 0);

      this.semaforos[`${id}-${c.key}-mes`] = this.calcularColor(valAct, valAnt);
      this.semaforos[`${id}-${c.key}-mad`] = this.calcularColor(valMadAct, valMadAnt);
    });
  }
  calcularColor(act: number, ant: number): any {
    if (act > ant) {
      return { 'background-color': '#d4edda', 'color': '#155724', 'font-weight': 'bold' };
    } else if (act < ant) {
      return { 'background-color': '#f8d7da', 'color': '#721c24', 'font-weight': 'bold' };
    } else {
      return { 'background-color': '#fff3cd', 'color': '#856404', 'font-weight': 'bold' };
    }
  }
  // Agrupa por rango de fechas y separa global/gerente/vendedor
  filtrarPorFechas(dataPorGerente: any, global: any[]) {
    const resultadoFinal: any = {};

    Object.keys(dataPorGerente).forEach(nombreGerente => {
      const data = dataPorGerente[nombreGerente];
      const grupos: any = {};

      data.forEach((item: any) => {
        const key = `${item.Fecha_inicio}|${item.Fecha_fin}`;

        if (!grupos[key]) {
          grupos[key] = {
            global: null,
            gerente: null,
            vendedores: []
          };
        }

        if (item.tipo_registro === 'GERENTE') {
          grupos[key].gerente = item;
        }

        if (item.tipo_registro === 'VENDEDOR') {
          grupos[key].vendedores.push(item);
        }
      });

      // Agregar global a cada grupo
      Object.keys(grupos).forEach(key => {
        const g = global.find(x =>
          x.Fecha_inicio === key.split('|')[0] &&
          x.Fecha_fin === key.split('|')[1]
        );
        grupos[key].global = g || null;
      });

      resultadoFinal[nombreGerente] = grupos;
    });

    this.paginas = resultadoFinal;


    // LÓGICA DE PAGINACIÓN: Seleccionar el primer gerente encontrado para que no inicie vacío
    const nombresGerentes = Object.keys(this.paginas);
    if (nombresGerentes.length > 0) {
      this.gerenteSeleccionado = nombresGerentes[0];
      this.procesarSemaforos(this.gerenteSeleccionado);
    }
  }

  // Calcula días laborables descontando domingos y feriados oficiales de México 2026
  calcularDiasHabiles(inicio: string, fin: string): number {
    let count = 0;
    let fecha = new Date(inicio.replace(/-/g, '\/'));
    const finDate = new Date(fin.replace(/-/g, '\/'));

    const festivos2026 = [
      '01-01', // Año Nuevo
      '02-02', // Constitución (5 Feb)
      '03-16', // Natalicio Juárez (21 Mar)
      '04-02', // Jueves Santo 
      '04-03', // Viernes Santo
      '04-04', // Sabado Santo
      '05-01', // Día del Trabajo
      '09-16', // Independencia
      '11-16', // Revolución (20 Nov)
      '12-25'  // Navidad
    ];

    while (fecha <= finDate) {
      const mesDia = `${(fecha.getMonth() + 1).toString().padStart(2, '0')}-${fecha.getDate().toString().padStart(2, '0')}`;

      if (fecha.getDay() !== 0 && !festivos2026.includes(mesDia)) {
        count++;
      }
      fecha.setDate(fecha.getDate() + 1);
    }
    return count;
  }

  // Cambia tu getObjMes actual por este:
  getObjMes(i: number, numApv: number): number {
    const base = [13, 24, 12, 9, 7, 6, 6, 6];
    return base[i] * numApv;
  }

  // Cambia tu getObjDia actual por este:
  getObjDia(i: number, col: any, numApv: number): number {
    const objMes = this.getObjMes(i, numApv);
    const dias = this.calcularDiasHabiles(col.fechaInicio, col.fechaFin);
    return Math.round((objMes / 26) * dias);
  }

  // También actualiza getPorcentaje para que los semáforos cuadren con los nuevos objetivos:
  getPorcentaje(valor: number, i: number, col: any, numApv: number): number {
    const objDia = this.getObjDia(i, col, numApv);
    if (!objDia || !valor) return 0;
    return Math.round((valor / objDia) * 100);
  }
  getFilas(columnas: any[]) {
    const filas: any[] = [];
    let temp: any[] = [];

    columnas.forEach((c, i) => {
      temp.push(c);
      if ((i + 1) % 3 === 0) {
        filas.push(temp);
        temp = [];
      }
    });

    if (temp.length) filas.push(temp);
    return filas;
  }

  getColumnas(pagina: any): any[] {
    if (!pagina) return [];

    return Object.keys(pagina)
      .sort((a, b) => {
        const f1 = new Date(a.split('|')[1]).getTime();
        const f2 = new Date(b.split('|')[1]).getTime();
        return f1 - f2;
      })
      .map(key => ({
        ...pagina[key],
        fechaInicio: key.split('|')[0],
        fechaFin: key.split('|')[1]
      }));
  }

  normalizarData(data: any[]) {
    return data.map(item => ({
      ...item,
      mes: item.Mes || {},
      maduracion: item.Maduracion || {}
    }));
  }

  anteriorGerente() {
    const nombres = Object.keys(this.paginas);
    const index = nombres.indexOf(this.gerenteSeleccionado);
    if (index > 0) {
      this.gerenteSeleccionado = nombres[index - 1];
    } else {
      // Opcional: Volver al último si está en el primero (loop)
      this.gerenteSeleccionado = nombres[nombres.length - 1];
    }
    this.procesarSemaforos(this.gerenteSeleccionado);
  }

  // Método para ir al siguiente gerente
  siguienteGerente() {
    const nombres = Object.keys(this.paginas);
    const index = nombres.indexOf(this.gerenteSeleccionado);
    if (index < nombres.length - 1) {
      this.gerenteSeleccionado = nombres[index + 1];
    } else {
      // Opcional: Volver al primero si está en el último (loop)
      this.gerenteSeleccionado = nombres[0];
    }
    this.procesarSemaforos(this.gerenteSeleccionado);
  }

  async exportarExcel() {
    const workbook = new ExcelJS.Workbook();
    await this.presentToast('Generando archivo Excel, por favor espera...', 'primary');

    for (const nombreGerente of Object.keys(this.paginas)) {
      // Recalcular semáforos para el gerente actual de la iteración
      this.procesarSemaforos(nombreGerente);

      const sheet = workbook.addWorksheet(nombreGerente.substring(0, 31));
      const columnas = this.getColumnas(this.paginas[nombreGerente]);
      const borderStyle: Partial<ExcelJS.Borders> = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };

      sheet.getColumn(1).width = 35;

      columnas.forEach((col, iCol) => {
        const tieneD = iCol > 0;
        const anchoBloque = tieneD ? 7 : 6;
        const colInicioNum = 2 + (iCol * 8);

        let fActual = 1;
        const celdaDia = sheet.getCell(fActual, colInicioNum);
        celdaDia.value = `DIA HÁBIL: ${this.calcularDiasHabiles(col.fechaInicio, col.fechaFin)}`;
        celdaDia.font = { bold: true };
        celdaDia.alignment = { horizontal: 'center' };
        sheet.mergeCells(fActual, colInicioNum, fActual, colInicioNum + (anchoBloque - 1));

        fActual++;
        const celdaRango = sheet.getCell(fActual, colInicioNum);
        celdaRango.value = `INI: ${col.fechaInicio} | FIN: ${col.fechaFin}`;
        celdaRango.font = { size: 9 };
        celdaRango.alignment = { horizontal: 'center' };
        sheet.mergeCells(fActual, colInicioNum, fActual, colInicioNum + (anchoBloque - 1));

        let filaDeTablas = 5;

        const dibujarTabla = (dataItem: any, titulo: string, idPrefix: string, iV?: number) => {
          // Título de la sección
          sheet.getCell(filaDeTablas, 1).value = titulo;
          sheet.getCell(filaDeTablas, 1).font = { bold: true };

          const numCalculo = idPrefix === 'global' ? 22 :
            idPrefix === 'gerente' ? col.vendedores.length : 1;

          const headers = ['OBJ MES', 'OBJ DIA', 'MES TOT', 'MES %', 'MAD TOT', 'MAD %'];
          if (tieneD) headers.push('D');

          headers.forEach((h, idx) => {
            const cell = sheet.getCell(filaDeTablas, colInicioNum + idx);
            cell.value = h;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } };
            cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 9 };
            cell.alignment = { horizontal: 'center' };
            cell.border = borderStyle;
          });

          filaDeTablas++;

          this.conceptos.forEach((c, iConc) => {
            const r = filaDeTablas + iConc;

            if (iCol === 0) {
              sheet.getCell(r, 1).value = c.nombre;
              sheet.getCell(r, 1).border = borderStyle;
            }

            // Valores con el nuevo parámetro numCalculo
            sheet.getCell(r, colInicioNum).value = this.getObjMes(iConc, numCalculo);
            sheet.getCell(r, colInicioNum + 1).value = this.getObjDia(iConc, col, numCalculo);
            sheet.getCell(r, colInicioNum + 2).value = dataItem?.Mes?.[c.key] || 0;

            // Porcentaje MES con semáforo
            const pctMes = this.getPorcentaje(dataItem?.Mes?.[c.key], iConc, col, numCalculo);
            const cPctMes = sheet.getCell(r, colInicioNum + 3);
            cPctMes.value = pctMes / 100;
            cPctMes.numFmt = '0%';
            this.aplicarColorExcel(cPctMes, idPrefix, iCol, c.key, 'mes', iV);

            sheet.getCell(r, colInicioNum + 4).value = dataItem?.Maduracion?.[c.key] || 0;

            // Porcentaje MAD con semáforo
            const pctMad = this.getPorcentaje(dataItem?.Maduracion?.[c.key], iConc, col, numCalculo);
            const cPctMad = sheet.getCell(r, colInicioNum + 5);
            cPctMad.value = pctMad / 100;
            cPctMad.numFmt = '0%';
            this.aplicarColorExcel(cPctMad, idPrefix, iCol, c.key, 'mad', iV);

            if (tieneD) {
              sheet.getCell(r, colInicioNum + 6).value = this.getConversionD(iConc, dataItem);
            }

            // Bordes y alineación
            for (let k = 0; k < anchoBloque; k++) {
              sheet.getCell(r, colInicioNum + k).border = borderStyle;
              sheet.getCell(r, colInicioNum + k).alignment = { horizontal: 'center' };
            }
          });
          filaDeTablas += this.conceptos.length + 2;
        };

        // Ejecutar el dibujo de tablas para cada nivel
        dibujarTabla(col.global, 'GLOBAL', 'global');
        dibujarTabla(col.gerente, nombreGerente, 'gerente');
        col.vendedores.forEach((v: any, iV: number) => {
          dibujarTabla(v, v.Apv_nombre, 'vendedor', iV);
        });
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    await this.presentToast('¡Excel exportado con éxito!');
    saveAs(new Blob([buffer]), `Reporte_APV_${this.sucursal}.xlsx`);

    this.procesarSemaforos(this.gerenteSeleccionado);
  }

  // Función auxiliar para aplicar colores de fondo y fuente basados en el estado
  aplicarColorExcel(celda: any, tipo: string, iCol: number, key: string, campo: string, iV?: number) {
    const id = iV !== undefined ? `vendedor-${iV}-${iCol}-${key}-${campo}` : `${tipo}-${iCol}-${key}-${campo}`;
    const estilo = this.semaforos[id];

    if (estilo && estilo['background-color']) {
      celda.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: estilo['background-color'].replace('#', 'FF').toUpperCase() }
      };
      if (estilo['color']) {
        celda.font = { color: { argb: estilo['color'].replace('#', 'FF').toUpperCase() }, bold: true };
      }
    }
  }

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
    const toast = await this.toastCtrl.create({
      message,
      duration: 2500,
      position: 'middle',
      color,
    });
    toast.present();
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

}