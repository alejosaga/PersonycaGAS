
var cachedData = {}; // Definir un objeto para almacenar los datos en memoria temporalmente
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
var nit = sheetDatos.getRange(lastRowDat,5).getValue();
var razonSocial = sheetDatos.getRange(lastRowDat,6).getValue();
var cliCargo = sheetDatos.getRange(lastRowDat,9).getValue();
var cliContacto = sheetDatos.getRange(lastRowDat,7).getValue();
var area = sheetDatos.getRange(lastRowDat,8).getValue();
var numEmp = sheetDatos.getRange(lastRowDat,11).getValue();
var numTra = sheetDatos.getRange(lastRowDat,28).getValue(); //numero exacto de trabajadores
var numContra = 0 //numero exacto de contratistas
var numCon = sheetDatos.getRange(lastRowDat,12).getValue();
var datCent = sheetDatos.getRange(lastRowDat,13).getValue();
var ciudades = sheetDatos.getRange(lastRowDat,14).getValue();
var claseRiesgo = sheetDatos.getRange(lastRowDat,17).getValue();
var numCiudades = ciudades.split(",").length;
var clientEmail1 = searchValues(maestroCotId,nit,"Datos","Nit","Dirección de correo electrónico");
var clientEmail2 = searchValues(maestroCotId,nit,"Datos","Nit","Segundo correo electronico (opcional)");
//Plantillas
var slideBatPsiId = "1LneOkKixIm1zOFnLpkjROqzcZ3O9ItMOs_a2EBGzShk";
var slideSevenStanId = "1Hx5R0CPuA85A8ThQJijn8-j9oc2qyB7OPFlsL9FT52s"
var slideSgsstId = "1QsXMBN4IznmhdsEO_iL1cajdmj3J5dWFF3Lm5KUXrpQ"
var contrato7estandares = "1Vh08aAOlnOC8CAi5IIFk0mLa-W62ISHNw9F8_1WAJag"
var contratoSgsst = "1eK31KnMMn1pT3BmBIeHsmAWwVfamDnqSetH_FY-4TH4"
  
//fecha de hoy
var today = new Date();
var dd = today.getDate();
var mm = today.getMonth()+1; //January is 0!
var yyyy = today.getFullYear();

if(dd<10) {
    dd = '0'+dd
} 

if(mm<10) {
    mm = '0'+mm
} 

today = mm + '/' + dd + '/' + yyyy;

