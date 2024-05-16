function sieteEstandares() {
  const BD_servicio = sevenStandarServiceId;
  const SServicio = SpreadsheetApp.openById(BD_servicio);
  const sheetCotizaciones = SServicio.getSheetByName(servicio);
  const lastRowCot = sheetCotizaciones.getLastRow();
  const lastColumnCot = sheetCotizaciones.getLastColumn();
  var slide = slideSevenStanId;
  
   
  
  var tarifaBasica = 1300000;
  var total = tarifaBasica;

  //Setear valores
  
  sheetCotizaciones.getRange(lastRowCot+1,1).setValue(nit);
  sheetCotizaciones.getRange(lastRowCot+1,2).setValue(razonSocial);
  sheetCotizaciones.getRange(lastRowCot+1,4).setValue(today);
  sheetCotizaciones.getRange(lastRowCot+1,8).setValue(tarifaBasica);
  sheetCotizaciones.getRange(lastRowCot+1,9).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,10).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,11).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,12).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,13).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,14).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,15).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,16).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,17).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,18).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,19).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,20).setValue(0);
  sheetCotizaciones.getRange(lastRowCot+1,5).setValue(total);
  
  addRowNumber(SServicio,servicio,3);
  
  
  var carpeta = crearCarpetaCot("SG-SST-7-ES")
  var folderCotId = carpeta.id;
  var folderCotUrl = carpeta.url;
  sheetCotizaciones.getRange(lastRowCot+1,6).setValue(folderCotUrl);   
  
  //Obtener Actividad Economica
  var ciiu = sheetDatos.getRange(lastRowDat,16).getValue();
  var result = obtenerActividad(ciiu);

  
  
  //Identificaciones
  var slideID = slide;
  var servi = "SG-SST-7-ES";
  var ids = getFolderIds(folderCotId);  
  var pdfID = ids.pdf;
  var temporal = ids.doc;

  
  //Conexiones
  var slide = SlidesApp.openById(slide);
  var archivoslide = DriveApp.getFileById(slideID);
  var carpetaPDF = DriveApp.getFolderById(pdfID);
  var carpetaTemporal = DriveApp.getFolderById(temporal);
      
  //Datos Cotizacion.
  var resConsecutivo = sheetCotizaciones.getRange(lastRowCot+1, 3).getValue();
  var valor = sheetCotizaciones.getRange(lastRowCot+1,5).getValue();
  var pesos = formatoColombiano(valor);
      
  var valorletras = numeroALetras(valor, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });

  var copiaArchivoslide = archivoslide.makeCopy(carpetaTemporal);
  var copiaID = copiaArchivoslide.getId();
  var nombreDoc = "CO-"+servi+"-"+resConsecutivo+"-"+yyyy+" "+razonSocial+" NIT "+nit;
  var archivo = DriveApp.getFileById(copiaID)
  archivo.setName(nombreDoc);
  var slide = SlidesApp.openById(copiaID);
  var slideLink = slide.getUrl();

  var presentacion= SlidesApp.openById(copiaID)
  presentacion.replaceAllText("{{fecha}}", today)
  presentacion.replaceAllText("{{anio}}", yyyy)
  presentacion.replaceAllText("{{servicio}}", servicio)
  presentacion.replaceAllText("{{numCot}}", "COT-"+servi+"-"+resConsecutivo)
  presentacion.replaceAllText("{{cargo}}", cliCargo)
  presentacion.replaceAllText("{{nombre}}", cliContacto)
  presentacion.replaceAllText("{{razonSocial}}", razonSocial)
  presentacion.replaceAllText("{{numEmp}}", numEmp)
  presentacion.replaceAllText("{{claseRiesgo}}", claseRiesgo)
  presentacion.replaceAllText("{{valor}}", pesos)
  presentacion.replaceAllText("{{valorLetras}}", valorletras.toLowerCase())
  presentacion.replaceAllText("{{area}}", area)
   
  presentacion.saveAndClose();

  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot).setValue(nombreDoc);  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-1).setValue(copiaID)
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-2).setValue(pdfID)

  //Datos del cliente y respuestas del form

  
  var prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot,23);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);

  var dataClient = htmlData(SSmaestroCot,"Datos",2,23);
  var dataValue = htmlData(SServicio,servicio,3,6);
  var dataToSend = result+dataClient+dataValue;

  sendEmail(nombreDoc,slideLink,dataToSend) 

}
