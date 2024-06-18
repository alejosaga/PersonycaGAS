
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
let clientCod = sheetDatos.getRange(lastRowDat,lastColumnDat-3).getValue();
let nit = sheetDatos.getRange(lastRowDat,5).getValue();
let razonSocial = sheetDatos.getRange(lastRowDat,6).getValue();
let cliCargo = sheetDatos.getRange(lastRowDat,9).getValue();
let cliContacto = sheetDatos.getRange(lastRowDat,7).getValue();
let area = sheetDatos.getRange(lastRowDat,8).getValue();
let numEmp = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","¿Cuántos trabajadores tiene actualmente directos?*");
let numTra = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Por favor indique la cantidad de trabajadores que deben aplicar para la bateria de riesgo Psicosocial."); //numero exacto de trabajadores
let numContra = 0 //numero exacto de contratistas
let numCon = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","¿Cuántos contratistas tiene actualmente?");
let datCent = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","¿Cuántos centros de trabajo tienes? (en numeros)");
let ciudades = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Indicanos las ciudades principales donde tiene trabajadores*");
let claseRiesgo = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Clase de riesgo");
let numCiudades = ciudades.split(",").length;
let consultoria = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","De acuerdo a sus necesidades seleccione el sistema de gestión sobre el cual requiere consultoría");
let clientEmail1 = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Dirección de correo electrónico");
let clientEmail2 = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Segundo correo electronico (opcional)");

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

