
let cachedData = {}; // Definir un objeto para almacenar los datos en memoria temporalmente
// Abrir hojas de calculo

const SSmaestroCot = SpreadsheetApp.openById(maestroCotId);
const sheetDatos = SSmaestroCot.getSheetByName('Datos');
const lastRowDat = sheetDatos.getLastRow();
const lastColumnDat = sheetDatos.getLastColumn();
const servicio = sheetDatos.getRange(lastRowDat,15).getValue()
const sheetTarifas = SSmaestroCot.getSheetByName('Tarifas');
const sheetCiiu = SSmaestroCot.getSheetByName('CIIU');
const tarifas = sheetTarifas.getDataRange().getValues();

//Datos estandar para cotizacion

const carpetaRaiz = DriveApp.getFolderById(carpetaRaizId);
const linkMaestro = SSmaestroCot.getUrl();
let nit = sheetDatos.getRange(lastRowDat,5).getValue();
let razonSocial = sheetDatos.getRange(lastRowDat,6).getValue();
let cliCargo = sheetDatos.getRange(lastRowDat,9).getValue();
let cliContacto = sheetDatos.getRange(lastRowDat,7).getValue();
let area = sheetDatos.getRange(lastRowDat,8).getValue();
let numEmp = sheetDatos.getRange(lastRowDat,11).getValue();
let numTra = sheetDatos.getRange(lastRowDat,28).getValue(); //numero exacto de trabajadores
let numContra = 0 //numero exacto de contratistas
let numCon = sheetDatos.getRange(lastRowDat,12).getValue();
let datCent = sheetDatos.getRange(lastRowDat,13).getValue();
let ciudades = sheetDatos.getRange(lastRowDat,14).getValue();
let claseRiesgo = sheetDatos.getRange(lastRowDat,17).getValue();
let numCiudades = ciudades.split(",").length;
let clientEmail1 = searchValues(maestroCotId,nit,"Datos","Nit","Dirección de correo electrónico");
let clientEmail2 = searchValues(maestroCotId,nit,"Datos","Nit","Segundo correo electronico (opcional)");

  
//fecha de hoy
let today = new Date();
let dd = today.getDate();
let mm = today.getMonth()+1; //January is 0!
let yyyy = today.getFullYear();

if(dd<10) {
    dd = '0'+dd
} 

if(mm<10) {
    mm = '0'+mm
} 

today = mm + '/' + dd + '/' + yyyy;

