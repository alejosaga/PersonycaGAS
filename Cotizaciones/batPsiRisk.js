function batPsiRisk() {
  const BD_servicio = batPsiServiceId;
  const SServicio = SpreadsheetApp.openById(BD_servicio);
  const sheetCotizaciones = SServicio.getSheetByName(servicio);
  const lastRowCot = sheetCotizaciones.getLastRow();
  const lastColumnCot = sheetCotizaciones.getLastColumn()
  var plantilla = plantillaBatPsiId;
   if (numTra <=100){
    var tarifaBasica = tarifas[34][6]
  }
  else if (numTra>100 && numTra<=300){
    var tarifaBasica = tarifas[35][6]
  }
  else if (numTra>300 && numTra<=500){
    var tarifaBasica = tarifas[36][6]
  }
  else {
    var tarifaBasica = tarifas[37][6]
  }

  var costosOperativos = tarifas[38][6]*numTra;
  var marketingSst = tarifas[39][6]*numTra;
  var traNoLee = sheetDatos.getRange(lastRowDat,37).getValue();
    
  var caracteristicas = [];


  var vlrAplicoAntes = servicio+"aplicoAntes"+ sheetDatos.getRange(lastRowDat,38).getValue();
  var vlEmpNoLee = servicio+"empNoLee";
  var vlrNumTra = servicio+"numeroTrabajadores";
  var vlrNumCiudDifBog = servicio+"numCiudDifBog";
  var vlrCentros = servicio+"centros";
  

  caracteristicas.push(vlrAplicoAntes);
  caracteristicas.push(vlEmpNoLee);
  caracteristicas.push(vlrNumTra);
  caracteristicas.push(vlrNumCiudDifBog);
  caracteristicas.push(vlrCentros);
   
  var valoresEncontrados = buscarTarifas(caracteristicas); 
  console.log(valoresEncontrados)

  valoresEncontrados.forEach((elemento, indice)=>{
    valoresEncontrados[indice] = elemento * tarifaBasica;
  });

 
  var aplicoAntes = valoresEncontrados[0];
  var empNoLee = valoresEncontrados[1]*traNoLee;
  var centros = valoresEncontrados[2]*datCent;
  var numeroTrabajadores = valoresEncontrados[3]*numTra;
  var numCiudDifBog = valoresEncontrados[4]*(numCiudades-1);
    
  var total = aplicoAntes+empNoLee+numCiudDifBog+numeroTrabajadores+centros+costosOperativos+marketingSst;
  var valPer = (total-empNoLee)/(numContra+numTra);
  var valAnticipo = total * 0.40;
 

  //Setear valores
  
  sheetCotizaciones.getRange(lastRowCot+1,1).setValue(nit);
  sheetCotizaciones.getRange(lastRowCot+1,2).setValue(razonSocial);
  sheetCotizaciones.getRange(lastRowCot+1,4).setValue(today);
  sheetCotizaciones.getRange(lastRowCot+1,8).setValue(tarifaBasica);
  sheetCotizaciones.getRange(lastRowCot+1,12).setValue(numeroTrabajadores);
  sheetCotizaciones.getRange(lastRowCot+1,10).setValue(aplicoAntes);
  sheetCotizaciones.getRange(lastRowCot+1,11).setValue(empNoLee);
  sheetCotizaciones.getRange(lastRowCot+1,16).setValue(numCiudDifBog);
  sheetCotizaciones.getRange(lastRowCot+1,14).setValue(valPer);
  sheetCotizaciones.getRange(lastRowCot+1,15).setValue(valAnticipo);
  sheetCotizaciones.getRange(lastRowCot+1,9).setValue(centros);
  sheetCotizaciones.getRange(lastRowCot+1,18).setValue(costosOperativos);
  sheetCotizaciones.getRange(lastRowCot+1,17).setValue(marketingSst);
  sheetCotizaciones.getRange(lastRowCot+1,5).setValue(total);
  

  addRowNumber(SServicio,servicio,3);
  /*
  var prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);
*/
  var carpeta = crearCarpetaCot("BAT-PSI")
  var folderCotId = carpeta.id;
  var folderCotUrl = carpeta.url;
  sheetCotizaciones.getRange(lastRowCot+1,6).setValue(folderCotUrl);   
 
  
  

  //Identificaciones
  var plantillaID = plantilla;
  var servi = "BAT-PSI";
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
  var pesosTotal = formatoColombiano(total);
  var empNoLeePesos = formatoColombiano(empNoLee);
  var valAntiPesos = formatoColombiano(valAnticipo);
  var valPerPesos = formatoColombiano(valPer);
  
  var valorletras = numeroALetras(total, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
});
 var valAntiLetras = numeroALetras(valAnticipo, {
  plural: "PESOS",
  singular: "PESO",
  centPlural: "CENTAVOS",
  centSingular: "CENTAVO"
});
 var valPerLetras = numeroALetras(valPer, {
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
  presentacion.replaceAllText("{{valor}}", pesosTotal)
  presentacion.replaceAllText("{{valorLetras}}", valorletras.toLowerCase())
  presentacion.replaceAllText("{{area}}", area)
  presentacion.replaceAllText("{{numTra}}", numTra)
  presentacion.replaceAllText("{{valPer}}", valPerPesos)
  presentacion.replaceAllText("{{valPerLetras}}", valPerLetras.toLowerCase())
  presentacion.replaceAllText("{{valorLecto}}", empNoLeePesos)
  presentacion.replaceAllText("{{valAnti}}", valAntiPesos)
  presentacion.replaceAllText("{{valAntiLetras}}", valAntiLetras.toLowerCase())
  
  presentacion.saveAndClose();

 
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot).setValue(nombreDoc);  
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-1).setValue(copiaID)
  sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-2).setValue(pdfID);

  var prefilledForm = preFilledForm(total,sheetCotizaciones,lastRowCot,21);
  sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);

  //Datos del cliente y respuestas del form

  var dataClient1 = htmlData(SSmaestroCot,"Datos",2,9);
  var dataClient2 = htmlData(SSmaestroCot,"Datos",26,5);
  var dataClient3 = htmlData(SSmaestroCot,"Datos",37,2);
  var dataValue = htmlData(SServicio,servicio,3,16);
  var dataToSend = dataClient1+dataClient2+dataClient3+dataValue;

  sendEmail(nombreDoc,slideLink,dataToSend)

}
  

