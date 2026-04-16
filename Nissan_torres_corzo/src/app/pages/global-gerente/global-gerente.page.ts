import { Component, OnInit } from '@angular/core';
import { ExcelService } from 'src/app/service/exel-service';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { ActivatedRoute } from '@angular/router';
import { AlertController } from '@ionic/angular';
import { Sucursales } from 'src/app/service/sucursales';


@Component({
  selector: 'app-global-gerente',
  templateUrl: './global-gerente.page.html',
  styleUrls: ['./global-gerente.page.scss'],
  standalone: false
})

export class GlobalGerentePage implements OnInit {
  diasHabiles = 26; // Celda A3
  numAPV = 22;      // Celda B3
  diaActual = 2;    // Usado para cálculos de columna H
  reportesPorMes: any = {};
  mesActual: string = '';
  id: string = '';



  constructor(private api: ExcelService, private apiSuc: Sucursales, private act: ActivatedRoute, private alertCtrl: AlertController) {
    this.id = this.act.snapshot.paramMap.get('sucursal') as string;
  }

  async ngOnInit() {
    await this.getSucursal();
    await this.getSucursales();
    await this.getapv();
    await this.getGlobal();
    this.actualizarSemanal();
    this.agruparPorMes();
  }
  // =========================
  // CRUD TABLA NUM_APV
  // =========================

  apvData: any[] = [];
  sucursales: any[] = [];

  modalApv = false;
  modoEdicion = false;

  formApv = {
    id: null,
    gerente: '',
    sucursal: '',
    num_apv: 0
  };

  // NORMALIZAR Ñ → N
  normalizarTexto(texto: string): string {
    if (!texto) return '';
    return texto
      .replace(/ñ/g, 'n')
      .replace(/Ñ/g, 'N');
  }

  // ABRIR MODAL
  abrirModalApv() {
    this.modalApv = true;
    this.getApv();
    this.obtenerSucursales();
  }

  // CERRAR
  cerrarModalApv() {
    this.modalApv = false;
    this.resetForm();
  }

  // RESET
  resetForm() {
    this.formApv = {
      id: null,
      gerente: '',
      sucursal: '',
      num_apv: 0
    };
    this.modoEdicion = false;
  }


  // CREAR
  async crearApv() {
    try {

      const payload = {
        gerente: this.normalizarTexto(this.formApv.gerente),
        sucursal: this.formApv.sucursal,
        num_apv: Number(this.formApv.num_apv)
      };

      await this.apiSuc.createApv(payload);

      await this.getGlobal();
      await this.getapv();
      this.resetForm();

    } catch (err) {
      console.error(err);
    }
  }

  // EDITAR
  editarApv(item: any) {
    this.formApv = {
      id: item.id,
      gerente: item.gerente,
      sucursal: item.sucursal,
      num_apv: item.num_apv
    };
    this.modoEdicion = true;
  }

  // ACTUALIZAR
  async actualizarApv() {
    try {

      const payload = {
        gerente: this.normalizarTexto(this.formApv.gerente),
        sucursal: this.formApv.sucursal,
        num_apv: Number(this.formApv.num_apv)
      };

      await this.apiSuc.updateApv(this.formApv.id, payload);

      await this.obtenerApv();
      this.resetForm();

    } catch (err) {
      console.error(err);
    }
  }

  // BORRAR
  async eliminarApv(id: number) {
    try {
      await this.apiSuc.deleteApv(id);
      await this.obtenerApv();
    } catch (err) {
      console.error(err);
    }
  }
  async getapv() {

    await this.api.getApvGerente(this.id).then(res => {
      res.data.data.forEach((dato: any) => {
        const gerente = this.normalizarNombre(dato.Gerente);
        this.apvPorGerente[gerente] = Number(dato.Num_apv || 0);
      });

    })
      .catch(err => {
        console.error(err)
      });
  }

  normalizarNombre(nombre: string): string {
    return (nombre || '')
      .toUpperCase()
      .trim()
      .replace(/\s+/g, ' ');
  }

  getApvPorGerente(gerente: string): number {

    const cleanGerente = (gerente || '')
      .replace(/\s+/g, ' ')   // elimina dobles espacios
      .trim()
      .toUpperCase();

    const key = Object.keys(this.apvPorGerente).find(k =>
      k.replace(/\s+/g, ' ').trim().toUpperCase() === cleanGerente
    );

    return this.apvPorGerente[key || ''] || 0;
  }

  actualizarSemanal() {
    this.filasPrincipales.forEach(f => {
      f.sem = this.calcularSemanal(f.concepto, f.mes);
    });

    this.filasFinales.forEach(f => {
      f.sem = this.calcularSemanal(f.concepto, f.mes);
    });
  }

  diaLimite: number = 0;

  sucursal: any = {};

  async getSucursal() {
    try {
      const res = await this.api.getBySucursales(this.id);
      this.sucursal = res.data[0];

    } catch (error) {

    }
  }

  sucursales: any[] = [];

  async getSucursales() {
    try {
      const res = await this.api.getSucursales();
      this.sucursales = res.data;
      console.log(this.sucursales)
    } catch (error) {

    }
  }

  async dispararAutomatizacion() {

    if (!this.diaLimite || this.diaLimite < 1 || this.diaLimite > 31) {
      const alert = await this.alertCtrl.create({
        header: 'Error de Configuración',
        message: 'Por favor, ingresa un día válido entre 1 y 31 para procesar el reporte.',
        buttons: ['OK']
      });
      await alert.present();
      return;
    }
    try {

      await this.api.ExtraerDatosGerente(this.diaLimite, this.sucursal.Sucursal);

      await this.getGlobal();

    } catch (error) {
      throw error;
    }
  }



  agruparPorMes() {
    this.reportesPorMes = {};

    for (let r of this.Global) {
      // CORRECCIÓN: Dividir la cadena manualmente para evitar desfase de UTC
      const partes = r.fecha.split('-'); // ["2026", "04", "01"]
      const anio = partes[0];
      const mes = partes[1].replace(/^0/, ''); // Quita el cero inicial (ej: "04" -> "4")

      const claveMes = `${anio}-${mes}`;

      if (!this.reportesPorMes[claveMes]) {
        this.reportesPorMes[claveMes] = [];
      }


      this.reportesPorMes[claveMes].push(r);
    }

    // ... resto del código (meses Actual, etc)
  }

  get reportesMesActual() {
    return this.reportesPorMes[this.mesActual] || [];
  }


  async exportarExcelmes() {
    const workbook = new ExcelJS.Workbook();

    // Se inicia en 1 para omitir el primer elemento de this.Global
    for (let i = 1; i < this.Global.length; i++) {
      const reporte = this.Global[i];
      this.paginaActual = i;

      const fechaStr = reporte.fecha;
      const diaHojaOriginal = fechaStr.substring(8, 10).replace(/^0/, '');
      const anio = fechaStr.substring(0, 4);
      const mes = fechaStr.substring(5, 7);
      const diaActual = parseInt(fechaStr.substring(8, 10), 10);
      const diaAnterior = diaActual - 1;

      // VALIDACIÓN: Evita el error duplicando el nombre si ya existe
      let nombreFinal = diaHojaOriginal;
      if (workbook.getWorksheet(nombreFinal)) {
        nombreFinal = `${diaHojaOriginal}_${i}`; // Agrega sufijo único si se repite el día
      }
      const worksheet = workbook.addWorksheet(nombreFinal);

      // TÍTULO (Fila 1)
      worksheet.mergeCells('A1:M1');
      const titulo = worksheet.getCell('A1');
      titulo.value = `DESEMPEÑO GLOBAL C/ APV Y F & I (${this.id}) ${diaActual}/${mes}/${anio}`;
      titulo.font = { bold: true, size: 14 };
      titulo.alignment = { horizontal: 'center', vertical: 'middle' };
      titulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB3B' } };

      // DÍA HÁBIL ACTUAL (N1)
      const diasHabilesCell = worksheet.getCell('N1');
      diasHabilesCell.value = reporte.DIA_HABIL;
      diasHabilesCell.font = { bold: true };
      diasHabilesCell.alignment = { horizontal: 'center', vertical: 'middle' };
      diasHabilesCell.border = this.getBorder();

      // ENCABEZADOS (Fila 2)
      const headers = [
        'Días habiles al mes', '#APV', `${this.apvDisponibles} APV`, 'DESEMPEÑO SEMANAL POR APV',
        'DESEMPEÑO MENSUAL POR APV', 'OBJ DIARIO', 'OBJETIVO MENSUAL',
        `OBJ. ACUM MENSUAL AL DÍA (${diaActual})`, `REAL ACUMULADO AL DÍA (${diaAnterior})`,
        'OBJ. DIARIO', `REAL DEL DÍA (${diaActual})`, `ACUMUL. REAL DEL DIA (${diaActual})`, '%'
      ];

      const headerRow = worksheet.addRow(headers);
      headerRow.height = 40;
      headerRow.eachCell(cell => {
        cell.font = { bold: true, size: 9 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = this.getBorder();
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } };
      });

      // Función para procesar y agregar filas de datos
      const agregarFilas = (filas: any[]) => {
        for (let f of filas) {
          const campo6 = this.getCampo6(f.concepto, reporte);
          const campo7 = this.getRealDelDia(f.concepto, reporte);
          const campo9 = this.getCampo9(f.concepto, reporte);
          const porcentaje = this.getPorcentaje(f.concepto, reporte, f.mes, reporte.DIA_HABIL);

          const row = worksheet.addRow([
            this.diasHabiles, this.apvDisponibles, f.concepto, f.sem, f.mes,
            Math.round(this.calcObjDiario(f.mes, reporte.Gerente)), Math.round(this.calcObjMensual(f.mes, reporte.Gerente)),
            Math.round(this.calcObjAcumAlDia(f.mes, reporte.DIA_HABIL, reporte.Gerente)),
            Math.round(campo6), Math.round(this.calcObjDiario(f.mes, reporte.Gerente)),
            Math.round(campo7), Math.round(campo9), porcentaje
          ]);

          row.getCell(13).numFmt = '0%';
          row.eachCell(cell => {
            cell.border = this.getBorder();
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          });

          const clase = this.getSemaforo(f.concepto, reporte, f.mes, reporte.DIA_HABIL);
          this.aplicarColorSemaforoExcel(row.getCell(13), clase);
        }
      };

      agregarFilas(this.filasPrincipales.slice(1));
      agregarFilas(this.filasFinales);

      // Ajuste de anchos de columna
      worksheet.columns.forEach((column, index) => {
        column.width = (index === 2) ? 25 : 10;
      });
    }

    // Generar y guardar archivo
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Reporte_Global_${this.gerenteSeleccionado}.xlsx`);
  }

  getBorder(): Partial<ExcelJS.Borders> {
    return {
      top: { style: 'thin' as ExcelJS.BorderStyle },
      left: { style: 'thin' as ExcelJS.BorderStyle },
      bottom: { style: 'thin' as ExcelJS.BorderStyle },
      right: { style: 'thin' as ExcelJS.BorderStyle }
    };
  }

  aplicarColorSemaforoExcel(cell: any, clase: string) {

    if (clase === 'verde') {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC8E6C9' }
      };
      cell.font = { bold: true, color: { argb: 'FF1B5E20' } };
    }

    if (clase === 'rojo') {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCDD2' }
      };
      cell.font = { bold: true, color: { argb: 'FFB71C1C' } };
    }

    if (clase === 'amarillo') {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFF9C4' }
      };
      cell.font = { bold: true, color: { argb: 'FFF57F17' } };
    }
  }

  Global: any[] = [];

  // Variables para el control de la vista
  reportesPorGerente: any = {};
  gerentesLista: string[] = [];
  gerenteSeleccionado: string = '';
  async getGlobal() {
    try {
      const res = await this.api.getDataGlobalGerente(this.id);
      const datos = res.data.data;


      this.reportesPorGerente = {};

      // 🔥 👉 AGREGAR ESTO (CLAVE)
      const datosProcesados = datos.map((registro: any) => {
        if (!registro.fecha) return { ...registro, DIA_HABIL: 0 };

        const partes = registro.fecha.split('-');
        const fechaActual = new Date(
          parseInt(partes[0]),
          parseInt(partes[1]) - 1,
          parseInt(partes[2])
        );

        let contadorHabilesInMes = 0;
        let fechaAux = new Date(fechaActual.getFullYear(), fechaActual.getMonth(), 1);

        while (fechaAux <= fechaActual) {
          if (!this.esDiaInhabilMexico(fechaAux)) contadorHabilesInMes++;
          fechaAux.setDate(fechaAux.getDate() + 1);
        }

        return { ...registro, DIA_HABIL: contadorHabilesInMes };
      });

      // 🔥 USA LOS PROCESADOS, NO LOS ORIGINALES
      datosProcesados.sort((a: any, b: any) =>
        new Date(a.fecha).getTime() - new Date(b.fecha).getTime()
      );

      datosProcesados.forEach((r: any) => {
        const nombreG = (r.Gerente || r.gerente || 'SIN ASIGNAR')
          .trim()
          .toUpperCase();
        if (!this.reportesPorGerente[nombreG]) this.reportesPorGerente[nombreG] = {};

        const fechaPartes = r.fecha.split('-');
        const mesClave = `${fechaPartes[0]}-${fechaPartes[1].replace(/^0/, '')}`;

        if (!this.reportesPorGerente[nombreG][mesClave]) {
          this.reportesPorGerente[nombreG][mesClave] = [];
        }

        this.reportesPorGerente[nombreG][mesClave].push(r);
      });

      this.gerentesLista = Object.keys(this.reportesPorGerente);


      if (this.gerentesLista.length > 0) {
        this.gerenteSeleccionado = this.gerentesLista[0];
        this.actualizarMeses();
      }



    } catch (e) {
      console.error(e);
    }
  }
  actualizarMeses() {
    const meses = Object.keys(this.reportesPorGerente[this.gerenteSeleccionado]);
    this.mesActual = meses.length > 0 ? meses[0] : '';
    this.actualizarVistaPrevia();
  }

  // NUEVA FUNCIÓN: Para que la tabla se actualice al filtrar
  actualizarVistaPrevia() {
    if (this.gerenteSeleccionado && this.mesActual) {
      this.Global = this.reportesPorGerente[this.gerenteSeleccionado][this.mesActual];
      this.paginaActual = 0;
      this.historico.clear();
    }
  }

  // NUEVA FUNCIÓN: Exportación masiva por Gerente
  async exportarTodosLosGerentes() {
    const backupGlobal = [...this.Global];

    for (const gerente of this.gerentesLista) {
      const reportesDelGerente: any[] = [];

      Object.keys(this.reportesPorGerente[gerente]).forEach(mes => {
        reportesDelGerente.push(...this.reportesPorGerente[gerente][mes]);
      });

      this.Global = reportesDelGerente;
      await this.exportarExcelmes();
    }

    this.Global = backupGlobal;
  }

  // NUEVA FUNCIÓN: Exportar solo el mes que se está viendo
  async exportarMesSeleccionado() {
    const backupGlobal = [...this.Global];
    this.Global = this.reportesPorGerente[this.gerenteSeleccionado][this.mesActual];
    await this.exportarExcelmes();
    this.Global = backupGlobal;
  }

  getRealDelDia(concepto: string, g: any): number {
    switch (concepto) {
      case 'PRECONTACTOS': return g?.preContactos || 0;
      case 'CONTACTOS': return g?.Contactos || 0;
      case 'PROSPECTOS': return g?.prospectos || 0;
      case 'SOL C/ DATOS COMPLETOS': return g?.solCDatosCompletos || 0;
      case 'VIABLES(PRE-AUTORIZADAS)': return g?.viablesPreAutorizadas || 0;
      case 'CITAS AGENDADAS': return g?.citasAgendadas || 0;
      case 'CITAS REALES': return g?.citasReales || 0;
      case 'DOC COMPLETA': return g?.docCompleta || 0;
      case 'AUTORIZADAS': return g?.autorizadas || 0;
      case 'PEDIDO CON ANTICIPO': return g?.pedidosConAnticipo || 0;
      case 'DEMOS': return g?.demos || 0;
      case 'ENTREGAS': return g?.entregas || 0;
      case 'DESEMBOLSOS': return g?.desembolsos || 0;
      default: return 0;
    }
  }

  esDiaInhabilMexico(fecha: Date): boolean {
    const anio = fecha.getFullYear();
    const mes = fecha.getMonth();
    const dia = fecha.getDate();
    const diaSemana = fecha.getDay();

    if (diaSemana === 0) return true;

    const festivosFijos = [
      { m: 0, d: 1 },
      { m: 3, d: 1 },
      { m: 3, d: 2 },
      { m: 3, d: 3 },
      { m: 4, d: 1 },
      { m: 8, d: 16 },
      { m: 11, d: 25 }
    ];
    if (festivosFijos.some(f => f.m === mes && f.d === dia)) return true;

    if (mes === 1) {
      const primerLunes = this.getN_esimoDiaDelaSemana(anio, 1, 1, 1);
      if (dia === primerLunes) return true;
    }

    if (mes === 2) {
      const tercerLunes = this.getN_esimoDiaDelaSemana(anio, 2, 1, 3);
      if (dia === tercerLunes) return true;
    }

    if (mes === 10) {
      const tercerLunes = this.getN_esimoDiaDelaSemana(anio, 10, 1, 3);
      if (dia === tercerLunes) return true;
    }

    return false;
  }

  acumuladoActual: number = 0;

  cambiarPagina(index: number) {

    this.paginaActual = index;

    if (index === 0) {
      this.acumuladoActual = 0;
      return;
    }
  }

  private historico = new Map<string, number>();

  // Compara el porcentaje actual contra el del día anterior redondeando a enteros para el semáforo
  getSemaforo(concepto: string, reporteActual: any, mes: any, diaHabil: number): string {

    const reportes = this.reportesMesActual;
    const i = reportes.findIndex((r: any) => r.fecha === reporteActual.fecha);

    if (i === -1) return 'verde';

    const fechaHoy = reporteActual.fecha;
    const fechaAnt = i > 0 ? reportes[i - 1].fecha : null;

    const llaveHoy = `${concepto}-${fechaHoy}`;
    const llaveAnt = fechaAnt ? `${concepto}-${fechaAnt}` : null;

    const redondear = (v: number) => Math.round(v * 100);

    if (!this.historico.has(llaveHoy)) {
      const v = this.getPorcentaje(concepto, reporteActual, mes, reporteActual.DIA_HABIL || diaHabil);
      this.historico.set(llaveHoy, redondear(v));
    }

    if (llaveAnt && !this.historico.has(llaveAnt)) {
      const reporteAnt = reportes[i - 1];
      const vAnt = this.getPorcentaje(concepto, reporteAnt, mes, reporteAnt.DIA_HABIL);
      this.historico.set(llaveAnt, redondear(vAnt));
    }

    const actual = this.historico.get(llaveHoy);
    const anterior = llaveAnt ? this.historico.get(llaveAnt) : undefined;

    if (i <= 0 || actual === undefined || anterior === undefined) return 'verde';

    if (actual > anterior) return 'verde';
    if (actual < anterior) return 'rojo';

    return 'amarillo';
  }

  exportarExcel() {

    const tabla = document.getElementById('tablaReporte');

    if (!tabla) return;

    const html = tabla.outerHTML;

    const blob = new Blob([html], {
      type: 'application/vnd.ms-excel'
    });

    const url = window.URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `Reporte_Dia_${this.paginaActual + 1}.xls`;
    a.click();

    window.URL.revokeObjectURL(url);
  }

  getPorcentaje(concepto: string, reporteActual: any, mes: number, diaHabil: number): number {
    const realAcumulado = this.getCampo9(concepto, reporteActual);

    const objAcumulado = this.calcObjAcumAlDia(mes, diaHabil, reporteActual.Gerente);

    if (!objAcumulado || objAcumulado === 0) return 0;

    return realAcumulado / objAcumulado;
  }

  getN_esimoDiaDelaSemana(anio: number, mes: number, diaBuscado: number, n: number): number {
    let count = 0;
    let d = new Date(anio, mes, 1);
    while (d.getMonth() === mes) {
      if (d.getDay() === diaBuscado) {
        count++;
        if (count === n) return d.getDate();
      }
      d.setDate(d.getDate() + 1);
    }
    return -1;
  }

  // Filas antes de las notas
  filasPrincipales = [
    { concepto: 'PRECONTACTOS', desc: ' base de datos, 10 por APV(REGISTRO EN EL SISTEMA, NOMBRE NO. TEL CORREO)', sem: 0, mes: 260, objD: 190, objM: 4940, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'CONTACTOS', desc: ' 10 por APV(CARPETA DE NEG. FICHA TECNICA, OFERTA COMERCIAL)', sem: 0, mes: 260, objD: 190, objM: 4940, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'PROSPECTOS', desc: ' (2 diarios por APV) MANDAR SOLICITUD CON DATOS COMPLETOS Y FIRMADA', mes: 24, objD: 18, objM: 456, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'SOL C/ DATOS COMPLETOS', desc: ' 80% TRAFICO A PISO', sem: 0, mes: 24, objD: 18, objM: 456, rAnterior: 0, rHoy: 0, clase: 'yellow' },
    { concepto: 'VIABLES(PRE-AUTORIZADAS)', desc: '', sem: 0, mes: 19, objD: 14, objM: 361, rAnterior: 0, rHoy: 0, clase: 'orange' },
    { concepto: 'CITAS AGENDADAS', desc: '', sem: 0, mes: 24, objD: 18, objM: 456, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'CITAS REALES', desc: '', sem: 0, mes: 13, objD: 10, objM: 247, rAnterior: 0, rHoy: 0, clase: 'yellow' },
    { concepto: 'DOC COMPLETA', desc: ' 80% VIABLES', mes: 12, sem: 0, objD: 9, objM: 228, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'AUTORIZADAS', desc: ' 40% SOLICITUDES CON DATOS COMPLETOS', sem: 0, mes: 9, objD: 7, objM: 171, rAnterior: 0, rHoy: 0, clase: 'yellow' },
    { concepto: 'PEDIDO CON ANTICIPO', desc: ' 30% SOL. CON DATOS COMPLETOS', sem: 0, mes: 7, objD: 5, objM: 137, rAnterior: 0, rHoy: 0, clase: 'yellow' }
  ];

  // Filas después de las notas
  filasFinales = [
    { concepto: 'DEMOS', desc: ' 80% TRAFICO A PISO', sem: 0, mes: 15, objD: 11, objM: 274, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'ENTREGAS', desc: ' 100% DESEMBOLSOS', sem: 0, mes: 6, objD: 4, objM: 110, rAnterior: 0, rHoy: 0, clase: '' },
    { concepto: 'DESEMBOLSOS', desc: ' 70% AUTORIZADAS', sem: 0, mes: 6, objD: 4, objM: 110, rAnterior: 0, rHoy: 0, clase: '' }
  ];

  // home.page.ts
  obtenerDia(fecha: any): number {
    if (!fecha) return 0;
    return new Date(fecha).getDate();
  }

  apvDisponibles: any[] = [];
  apvPorGerente: { [key: string]: number } = {};

  calcObjDiario(mes: number, gerente: string): number {
    const apv = this.getApvPorGerente(gerente);
    const valor = (mes * apv) / this.diasHabiles;
    return Number(valor.toFixed(3));
  }

  calcObjMensual(mes: number, gerente: string): number {
    return this.calcObjDiario(mes, gerente) * this.diasHabiles;
  }

  calcObjAcumAlDia(mes: number, diaHabil: number, gerente: string): number {

    const objDiarioReal = this.calcObjDiario(mes, gerente);

    const objParaCalculo = objDiarioReal === 0 ? 1 : objDiarioReal;
    const diaParaCalculo = diaHabil === 0 ? 1 : diaHabil;

    return objParaCalculo * diaParaCalculo;
  }
  // filaAnteriorAcum: valor acumulado del día anterior (ACUMUL. REAL DEL DÍA)
  // realHoy: valor del REAL DEL DÍA actual
  calcAcumRealDia(filaAnteriorAcum: number, realHoy: number): number {
    return (filaAnteriorAcum || 0) + (realHoy || 0);
  }





  getCampo6(concepto: string, reporteActual: any): number {

    const reportes = this.reportesMesActual;
    const index = reportes.findIndex((r: any) => r.fecha === reporteActual.fecha);

    if (index <= 0) return 0;

    let acumulado = 0;

    for (let i = 0; i < index; i++) {
      const reporte = reportes[i];
      acumulado += this.getRealDelDia(concepto, reporte);
    }

    return acumulado;
  }

  // Método para obtener el valor que estaba en la columna "ACUMUL. REAL DEL DIA"
  getAcumuladoRealDia(concepto: string, registro: any): number {
    // Aquí obtienes el valor que se mostró en la columna "ACUMUL. REAL DEL DIA"
    // para ese registro. Puede ser un campo directo de Strapi o un cálculo

    // Si en Strapi ya tienes un campo que guarda ese valor:
    switch (concepto) {
      case 'PRECONTACTOS': return registro?.acumuladoPreContactos || 0;
      case 'CONTACTOS': return registro?.acumuladoContactos || 0;
      case 'PROSPECTOS': return registro?.acumuladoProspectos || 0;
      // ... etc
      default: return 0;
    }
  }

  getAcumuladoDelDia(c6: any, c7: any): number {
    return Number(c6 || 0) + Number(c7 || 0);
  }

  calcRealAcumulado(rAnt: number, rHoy: number): number {
    return (rAnt || 0) + (rHoy || 0);
  }

  calcRealAcumDiaAnterior(filas: any[], indiceActual: number, g: any): number {
    const diaHabil = g.DIA_HABIL;

    // Día 1: no hay acumulado anterior
    if (diaHabil === 1) return 0;

    // Día 2: usar valor del día 1 desde Strapi
    if (diaHabil === 2) {
      return this.getRealDelDia(filas[indiceActual].concepto, g);
    }

    // Día >= 3: usar ACUMULADO REAL DEL DÍA del registro anterior
    if (indiceActual > 0) {
      const filaAnterior = filas[indiceActual - 1];
      return this.calcAcumRealDelDia(filaAnterior.rAnterior, filaAnterior.rHoy);
    }

    return 0;
  }

  calcAcumRealDelDia(rAnterior: number, rHoy: number): number {
    // Aseguramos que nunca sea undefined
    const rAnt = rAnterior ?? 0;
    const rH = rHoy ?? 0;

    return rAnt + rH;
  }

  calcularSemanal(concepto: string, mes: number): number {

    const fijos: { [key: string]: number } = {
      'VIABLES(PRE-AUTORIZADAS)': 5,
      'CITAS REALES': 3,
      'AUTORIZADAS': 2,
      'PEDIDO CON ANTICIPO': 2,
      'DEMOS': 4,
      'ENTREGAS': 2,
      'DESEMBOLSOS': 2
    };

    // Si el concepto es fijo, regresa el valor fijo
    if (fijos[concepto] !== undefined) {
      return fijos[concepto];
    }

    // Si no es fijo, dividir mensual entre 4
    return Math.round(mes / 4);
  }

  calcularPorcentaje(realAcumulado: number, objAcumulado: number): number {
    if (!objAcumulado || objAcumulado === 0) return 0;
    return realAcumulado / objAcumulado;
  }

  calcH(objD: number) { return objD * this.diaActual; }
  calcN(rAnt: number, rHoy: number) {

    return (rAnt || 0) + (rHoy || 0);
  }

  calcO(rAnt: number, rHoy: number) {

    return Number(rAnt || 0) + Number(rHoy || 0);
  }

  getCampo9(concepto: string, reporteActual: any): number {

    const campo6 = this.getCampo6(concepto, reporteActual);
    const campo7 = this.getRealDelDia(concepto, reporteActual);

    return Number(campo6 || 0) + Number(campo7 || 0);
  }

  paginaActual: number = 6;
  reportesPorPagina: number = 1;


  get reporteActual() {
    return this.Global[this.paginaActual];
  }

  get totalPaginas(): number {
    // Restamos 1 porque el índice 0 es el que siempre omites
    return this.Global.length > 0 ? this.Global.length - 1 : 0;
  }

  paginaAnterior() {
    if (this.paginaActual > 0) {
      this.paginaActual--;
    }
  }

  // Asegúrate de que al navegar nunca superes length - 1
  paginaSiguiente() {
    if (this.paginaActual < this.Global.length - 1) {
      this.paginaActual++;
    }
  }

  irAUltima() {
    this.paginaActual = this.Global.length - 1;
  }

  irAPrimera() {
    this.paginaActual = 0;
  }

}