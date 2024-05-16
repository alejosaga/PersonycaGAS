function sgSst() {
  const BD_servicio = sgsstServiceId;
  const SServicio = SpreadsheetApp.openById(BD_servicio);
  const sheetCotizaciones = SServicio.getSheetByName(servicio);
  const lastRowCot = sheetCotizaciones.getLastRow();
  const lastColumnCot = sheetCotizaciones.getLastColumn()
  var plantilla = slideSgsstId
      
  var tarifaBasica = tarifas[1][6]
  var costosOperativos = tarifas[2][6];
  var marketingSst = tarifas[3][6];
  var vlrMant = 1000000
     
  var caracteristicas = [];

  var vlrEstandares = servicio+"estandares"+ sheetDatos.getRange(lastRowDat,19).getValue();
  var vlrrepAuto = servicio+"reporteAutoevaluacion"+ sheetDatos.getRange(lastRowDat,20).getValue();
  var vlrNumTra = servicio+"numeroTrabajadores"+ numEmp;
  var vlrNumCon = servicio+"numeroContratistas"+ numCon
  var vlrVehi = servicio+"vehiculos";
  var datVehi = sheetDatos.getRange(lastRowDat,21).getValue();
  var vlrCentros = servicio+"centros";
  var vlrAltRies = servicio+"altoRiesgo";
  var datAltRies = sheetDatos.getRange(lastRowDat,22).getValue();
  var datAltRiesSplited = datAltRies.split(",").length;
  var vlrEnf = servicio+"enfermedades";
  var datEnf = sheetDatos.getRange(lastRowDat,24).getValue();
  var vlrAc = servicio+"accidentes";
  var datAc = sheetDatos.getRange(lastRowDat,25).getValue();
  var vlrNivRies = servicio+"nivelRiesgo"+ claseRiesgo;

  caracteristicas.push(vlrEstandares);
  caracteristicas.push(vlrrepAuto);
  caracteristicas.push(vlrNumTra);
  caracteristicas.push(vlrNumCon);
  caracteristicas.push(vlrVehi);
  caracteristicas.push(vlrCentros);
  caracteristicas.push(vlrAltRies);
  caracteristicas.push(vlrEnf);
  caracteristicas.push(vlrAc);
  caracteristicas.push(vlrNivRies);



  var valoresEncontrados = buscarTarifas(caracteristicas); 

  
  
  var estandares = valoresEncontrados[0]* tarifaBasica;
  var reporteAutoevaluacion = valoresEncontrados[1]* tarifaBasica;
  var numeroTrabajadores = valoresEncontrados[2]* tarifaBasica;
  var numeroContratistas = valoresEncontrados[3]* tarifaBasica;
  var vehiculos = valoresEncontrados[4]*datVehi* tarifaBasica;
  var centros = valoresEncontrados[5]*datCent* tarifaBasica;
  
  if(datAltRies=="Ninguno de los anteriores"){
    var altoRiesgo = 0
  }
  else{
    var altoRiesgo = valoresEncontrados[6]*datAltRiesSplited* tarifaBasica

  }
  ;
  var enfermedades = valoresEncontrados[7]*datEnf*datEnf* tarifaBasica;
  var accidentes = valoresEncontrados[8]*datAc* tarifaBasica;
  var nivelRiesgo = valoresEncontrados[9]* tarifaBasica;

  
  
  var total = tarifaBasica+estandares+reporteAutoevaluacion+numeroTrabajadores+numeroContratistas+vehiculos+centros+altoRiesgo+enfermedades+accidentes+nivelRiesgo+costosOperativos+marketingSst;


  //Setear valores
  
  sheetCotizaciones.getRange(lastRowCot+1,1).setValue(nit);
  sheetCotizaciones.getRange(lastRowCot+1,2).setValue(razonSocial);
  sheetCotizaciones.getRange(lastRowCot+1,4).setValue(today);
  sheetCotizaciones.getRange(lastRowCot+1,8).setValue(tarifaBasica);
  sheetCotizaciones.getRange(lastRowCot+1,9).setValue(estandares);
  sheetCotizaciones.getRange(lastRowCot+1,10).setValue(reporteAutoevaluacion);
  sheetCotizaciones.getRange(lastRowCot+1,11).setValue(numeroTrabajadores);
  sheetCotizaciones.getRange(lastRowCot+1,12).setValue(numeroContratistas);
  sheetCotizaciones.getRange(lastRowCot+1,13).setValue(vehiculos);
  sheetCotizaciones.getRange(lastRowCot+1,14).setValue(centros);
  sheetCotizaciones.getRange(lastRowCot+1,15).setValue(altoRiesgo);
  sheetCotizaciones.getRange(lastRowCot+1,16).setValue(enfermedades);
  sheetCotizaciones.getRange(lastRowCot+1,17).setValue(accidentes);
  sheetCotizaciones.getRange(lastRowCot+1,18).setValue(nivelRiesgo);
  sheetCotizaciones.getRange(lastRowCot+1,19).setValue(costosOperativos);
  sheetCotizaciones.getRange(lastRowCot+1,20).setValue(marketingSst);
  sheetCotizaciones.getRange(lastRowCot+1,5).setValue(total);
  

  addRowNumber(SServicio,servicio,3);
  /*
  var prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);
  */

  var carpeta = crearCarpetaCot("SG-SST")
  var folderCotId = carpeta.id;
  var folderCotUrl = carpeta.url;
  sheetCotizaciones.getRange(lastRowCot+1,6).setValue(folderCotUrl);   
  
  //Obtener Actividad Economica
  var ciiu = sheetDatos.getRange(lastRowDat,16).getValue();
  var result = obtenerActividad(ciiu);

  

  //Identificaciones
  var plantillaID = plantilla;
  var servi = "SG-SST";
  var ids = getFolderIds(folderCotId);  
  var pdfID = ids.pdf;
  var temporal = ids.doc;

  
  //Conexiones
  var slide = SlidesApp.openById(plantilla);
  var archivoPlantilla = DriveApp.getFileById(plantillaID);
  var carpetaPDF = DriveApp.getFolderById(pdfID);
  var carpetaTemporal = DriveApp.getFolderById(temporal);
      
  //Datos Cotizacion.
  var resConsecutivo = sheetCotizaciones.getRange(lastRowCot+1, 3).getValue();
  var valor = sheetCotizaciones.getRange(lastRowCot+1,5).getValue();
  var pesos = formatoColombiano(valor);
  var pesosMant = formatoColombiano(vlrMant);
    
  var valorletras = numeroALetras(valor, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });

  var mantLetras = numeroALetras(vlrMant, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });
 
  var copiaArchivoPlantilla = archivoPlantilla.makeCopy(carpetaTemporal);
  var copiaID = copiaArchivoPlantilla.getId();
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
  presentacion.replaceAllText("{{vlrMant}}", pesosMant)
  presentacion.replaceAllText("{{mantLetras}}", mantLetras.toLowerCase())
  
  presentacion.saveAndClose();

  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot).setValue(nombreDoc);  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-1).setValue(copiaID)
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-2).setValue(pdfID)
 
   //Datos del cliente y respuestas del form

  var prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot,23);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);

  var dataClient = htmlData(SSmaestroCot,"Datos",2,23);
  var dataValue = htmlData(SServicio,servicio,3,18);
  var dataToSend = result+dataClient+dataValue;


  sendEmail(nombreDoc,slideLink,dataToSend) 

}

  

