function sgSst() {
  const BD_servicio = sgsstServiceId;
  const SServicio = SpreadsheetApp.openById(BD_servicio);
  const sheetCotizaciones = SServicio.getSheetByName(servicio);
  const lastRowCot = sheetCotizaciones.getLastRow();
  const lastColumnCot = sheetCotizaciones.getLastColumn()
  //let plantilla = slideSgsstId

  let datEstandares = searchValues(maestroCotId,nit,"Nit","Dando alcance a la resolución 0312 del 2019 indicanos si tienen estándares mínimos actualmente y en que %");
  let datRepAuto = searchValues(maestroCotId,nit,"Nit","Tienes reporte de la autoevaluación ante el ministerio de trabajo del año 2019, 2020, 2021,  2022 y 2023");    
  
  let tarifaBasica = tarifas[1][6]
  let costosOperativos = tarifas[2][6];
  let marketingSst = tarifas[3][6];
  let vlrMant = 1000000
     
  let caracteristicas = [];

  let vlrEstandares = servicio+"estandares"+ datEstandares;
  let vlrrepAuto = servicio+"reporteAutoevaluacion"+ datRepAuto;
  let vlrNumTra = servicio+"numeroTrabajadores"+ numEmp;
  let vlrNumCon = servicio+"numeroContratistas"+ numCon
  let vlrVehi = servicio+"vehiculos";
  let datVehi = searchValues(maestroCotId,nit,"Nit","¿Tienen vehiculos? indicanos cuántos (en numeros)");
  let vlrCentros = servicio+"centros";
  let vlrAltRies = servicio+"altoRiesgo";
  let datAltRies = searchValues(maestroCotId,nit,"Nit","La empresa realiza trabajos de alto riesgo tales como:*");
  let datAltRiesSplited = datAltRies.split(",").length;
  let vlrEnf = servicio+"enfermedades";
  let datEnf = searchValues(maestroCotId,nit,"Nit","Indique cuantos casos de trabajadores con alguna enfermedad laboral en trámite tiene la compañia actualmente (0 si ninguno)");
  let vlrAc = servicio+"accidentes";
  let datAc = searchValues(maestroCotId,nit,"Nit","Cuantos casos de trabajadores con accidentes laborales en proceso. (0 si ninguno)")
  let vlrNivRies = servicio+"nivelRiesgo"+ claseRiesgo;

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



  let valoresEncontrados = buscarTarifas(caracteristicas); 

  
  
  let estandares = valoresEncontrados[0]* tarifaBasica;
  let reporteAutoevaluacion = valoresEncontrados[1]* tarifaBasica;
  let numeroTrabajadores = valoresEncontrados[2]* tarifaBasica;
  let numeroContratistas = valoresEncontrados[3]* tarifaBasica;
  let vehiculos = valoresEncontrados[4]*datVehi* tarifaBasica;
  let centros = valoresEncontrados[5]*datCent* tarifaBasica;
  let altoRiesgo = 0
  
  if(datAltRies=="Ninguno de los anteriores"){
     altoRiesgo = 0
  }
  else{
     altoRiesgo = valoresEncontrados[6]*datAltRiesSplited* tarifaBasica

  }
  ;
  let enfermedades = valoresEncontrados[7]*datEnf*datEnf* tarifaBasica;
  let accidentes = valoresEncontrados[8]*datAc* tarifaBasica;
  let nivelRiesgo = valoresEncontrados[9]* tarifaBasica;

  
  
  let total = tarifaBasica+estandares+reporteAutoevaluacion+numeroTrabajadores+numeroContratistas+vehiculos+centros+altoRiesgo+enfermedades+accidentes+nivelRiesgo+costosOperativos+marketingSst;


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
  let prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);
  */

  let carpeta = crearCarpetaCot("SG-SST")
  let folderCotId = carpeta.id;
  let folderCotUrl = carpeta.url;
  sheetCotizaciones.getRange(lastRowCot+1,6).setValue(folderCotUrl);   
  
  //Obtener Actividad Economica
  let ciiu = sheetDatos.getRange(lastRowDat,16).getValue();
  let result = obtenerActividad(ciiu);

  

  //Identificaciones
  //let plantillaID = plantilla;
  let servi = "SG-SST";
  let ids = getFolderIds(folderCotId);  
  let pdfID = ids.pdf;
  let temporal = ids.doc;

  
  //Conexiones
  //let slide = SlidesApp.openById(plantilla);
  let archivoPlantilla = DriveApp.getFileById(slideSgsstId);
  let carpetaPDF = DriveApp.getFolderById(pdfID);
  let carpetaTemporal = DriveApp.getFolderById(temporal);
      
  //Datos Cotizacion.
  let resConsecutivo = sheetCotizaciones.getRange(lastRowCot+1, 3).getValue();
  let valor = sheetCotizaciones.getRange(lastRowCot+1,5).getValue();
  let pesos = formatoColombiano(valor);
  let pesosMant = formatoColombiano(vlrMant);
    
  let valorletras = numeroALetras(valor, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });

  let mantLetras = numeroALetras(vlrMant, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
  });
 
  let copiaArchivoPlantilla = archivoPlantilla.makeCopy(carpetaTemporal);
  let copiaID = copiaArchivoPlantilla.getId();
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
  presentacion.replaceAllText("{{vlrMant}}", pesosMant)
  presentacion.replaceAllText("{{mantLetras}}", mantLetras.toLowerCase())
  
  presentacion.saveAndClose();

  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot).setValue(nombreDoc);  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-1).setValue(copiaID)
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-2).setValue(pdfID)
 
   //Datos del cliente y respuestas del form

  let prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot,23);
  sheetCotizaciones.getRange(lastRowCot,7).setValue(prefilledForm);

  sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(nombreDoc);

  let dataClient = htmlData(SSmaestroCot,"Datos",2,13);
  let dataClient1 = htmlData(SSmaestroCot,"Datos",23,10);
  let dataValue = htmlData(SServicio,servicio,3,18);
  let dataToSend = result+dataClient+dataClient1+dataValue;


  sendEmail(nombreDoc,slideLink,dataToSend) 

}

  

