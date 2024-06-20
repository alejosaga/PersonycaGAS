function sieteEstandares() {
  const BD_servicio = sevenStandarServiceId;
  const SServicio = SpreadsheetApp.openById(BD_servicio);
  const sheetCotizaciones = SServicio.getSheetByName(servicio);
  const lastRowCot = sheetCotizaciones.getLastRow();
  const lastColumnCot = sheetCotizaciones.getLastColumn();
  //let slide = slideSevenStanId;
  
   
  
  let tarifaBasica = 1300000;
  let total = tarifaBasica;

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
  
  
  let carpeta = crearCarpetaCot("SG-SST-7-ES")
  let folderCotId = carpeta.id;
  let folderCotUrl = carpeta.url;
  sheetCotizaciones.getRange(lastRowCot+1,6).setValue(folderCotUrl);   
  
  //Obtener Actividad Economica
  let ciiu = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Código CIIU de la empresa");

  let result = obtenerActividad(ciiu);

  
  
  //Identificaciones
  //let slideID = slide;
  let servi = "SG-SST-7-ES";
  let ids = getFolderIds(folderCotId);  
  let pdfID = ids.pdf;
  let temporal = ids.doc;

  
  //Conexiones
  //let TemplateSlide = SlidesApp.openById(slideSevenStanId);
  let archivoslide = DriveApp.getFileById(slideSevenStanId);
  let carpetaPDF = DriveApp.getFolderById(pdfID);
  let carpetaTemporal = DriveApp.getFolderById(temporal);
      
  //Datos Cotizacion.
  let resConsecutivo = sheetCotizaciones.getRange(lastRowCot+1, 3).getValue();
  let valor = sheetCotizaciones.getRange(lastRowCot+1,5).getValue();
  let pesos = formatoColombiano(valor);
      
  let valorletras = numeroALetras(valor, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });

  let copiaArchivoslide = archivoslide.makeCopy(carpetaTemporal);
  let copiaID = copiaArchivoslide.getId();
  let nombreDoc = "CO-"+servi+"-"+resConsecutivo+"-"+yyyy+" "+razonSocial+" NIT "+nit;
  let archivo = DriveApp.getFileById(copiaID)
  archivo.setName(nombreDoc);
  let slide = SlidesApp.openById(copiaID);
  let slideLink = slide.getUrl();

  let presentacion= SlidesApp.openById(copiaID)
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

  sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(nombreDoc);

  //Datos del cliente y respuestas del form

  
  let prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot,lastColumnCot,servicio);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);

  let dataClient = htmlData(SSmaestroCot,"Datos",1,13);
  let dataClient1 = htmlData(SSmaestroCot,"Datos",23,10);
  let dataValue = htmlData(SServicio,servicio,3,6);
  let dataToSend = result+dataClient+dataClient1+dataValue;

  sendEmail(nombreDoc,slideLink,dataToSend,servicio) 

}
