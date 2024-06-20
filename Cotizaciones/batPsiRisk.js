function batPsiRisk() {
  try {
    const SServicio = SpreadsheetApp.openById(batPsiServiceId);
    const sheetCotizaciones = SServicio.getSheetByName(servicio);
    const lastRowCot = sheetCotizaciones.getLastRow();
    const lastColumnCot = sheetCotizaciones.getLastColumn();

    let tarifaBasica;
    if (numTra <= 100) {
      tarifaBasica = tarifas[34][6];
    } else if (numTra > 100 && numTra <= 300) {
      tarifaBasica = tarifas[35][6];
    } else if (numTra > 300 && numTra <= 500) {
      tarifaBasica = tarifas[36][6];
    } else {
      tarifaBasica = tarifas[37][6];
    }

    console.log('Tarifa Básica:', tarifaBasica);

    let costosOperativos = tarifas[38][6] * numTra;
    let marketingSst = tarifas[39][6] * numTra;
    let traNoLee = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","La empresa cuenta con trabajadores que no saben leer ni escribir? cuantos?");
    let datAplicoAntes = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Se ha aplicado la bateria de riesgo psicosocial antes?");
    console.log('Costos Operativos:', costosOperativos);
    console.log('Marketing SST:', marketingSst);
    console.log('Tra No Lee:', traNoLee);

    let caracteristicas = [
      servicio + "aplicoAntes" + datAplicoAntes,
      servicio + "empNoLee",
      servicio + "numeroTrabajadores",
      servicio + "numCiudDifBog",
      servicio + "centros"
    ];

    let valoresEncontrados = buscarTarifas(caracteristicas);
    console.log('Valores Encontrados:', valoresEncontrados);

    valoresEncontrados = valoresEncontrados.map((elemento) => elemento * tarifaBasica);
    console.log('Valores Encontrados Ajustados:', valoresEncontrados);

    let aplicoAntes = valoresEncontrados[0];
    let empNoLee = valoresEncontrados[1] * traNoLee;
    let centros = valoresEncontrados[2] * datCent;
    let numeroTrabajadores = valoresEncontrados[3] * numTra;
    let numCiudDifBog = valoresEncontrados[4] * (numCiudades - 1);

    console.log('Aplico Antes:', aplicoAntes);
    console.log('Emp No Lee:', empNoLee);
    console.log('Centros:', centros);
    console.log('Numero Trabajadores:', numeroTrabajadores);
    console.log('Num Ciud Dif Bog:', numCiudDifBog);

    let total = aplicoAntes + empNoLee + numCiudDifBog + numeroTrabajadores + centros + costosOperativos + marketingSst;
    let valPer = (total - empNoLee) / (numContra + numTra);
    let valAnticipo = total * 0.30;

    console.log('Total:', total);
    console.log('Val Per:', valPer);
    console.log('Val Anticipo:', valAnticipo);

    // Set values in the spreadsheet
    sheetCotizaciones.getRange(lastRowCot + 1, 1).setValue(nit);
    sheetCotizaciones.getRange(lastRowCot + 1, 2).setValue(razonSocial);
    sheetCotizaciones.getRange(lastRowCot + 1, 4).setValue(today);
    sheetCotizaciones.getRange(lastRowCot + 1, 8).setValue(tarifaBasica);
    sheetCotizaciones.getRange(lastRowCot + 1, 12).setValue(numeroTrabajadores);
    sheetCotizaciones.getRange(lastRowCot + 1, 10).setValue(aplicoAntes);
    sheetCotizaciones.getRange(lastRowCot + 1, 11).setValue(empNoLee);
    sheetCotizaciones.getRange(lastRowCot + 1, 16).setValue(numCiudDifBog);
    sheetCotizaciones.getRange(lastRowCot + 1, 14).setValue(valPer);
    sheetCotizaciones.getRange(lastRowCot + 1, 15).setValue(valAnticipo);
    sheetCotizaciones.getRange(lastRowCot + 1, 9).setValue(centros);
    sheetCotizaciones.getRange(lastRowCot + 1, 18).setValue(costosOperativos);
    sheetCotizaciones.getRange(lastRowCot + 1, 17).setValue(marketingSst);
    sheetCotizaciones.getRange(lastRowCot + 1, 5).setValue(total);

    addRowNumber(SServicio, servicio, 3);

    let carpeta = crearCarpetaCot("BAT-PSI");
    let folderCotId = carpeta.id;
    let folderCotUrl = carpeta.url;
    sheetCotizaciones.getRange(lastRowCot + 1, 6).setValue(folderCotUrl);

    let servi = "BAT-PSI";
    let ids = getFolderIds(folderCotId);
    let pdfID = ids.pdf;
    let temporal = ids.doc;

    let archivoPlantilla = DriveApp.getFileById(slideBatPsiId);
    let carpetaPDF = DriveApp.getFolderById(pdfID);
    let carpetaTemporal = DriveApp.getFolderById(temporal);

    let resConsecutivo = sheetCotizaciones.getRange(lastRowCot + 1, 3).getValue();
    let pesosTotal = formatoColombiano(total);
    let empNoLeePesos = formatoColombiano(empNoLee);
    let valAntiPesos = formatoColombiano(valAnticipo);
    let valPerPesos = formatoColombiano(valPer);

    let valorletras = numeroALetras(total, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
    let valAntiLetras = numeroALetras(valAnticipo, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });
    let valPerLetras = numeroALetras(valPer, {
      plural: "PESOS",
      singular: "PESO",
      centPlural: "CENTAVOS",
      centSingular: "CENTAVO"
    });

    let copiaArchivoPlantilla = archivoPlantilla.makeCopy(carpetaTemporal);
    let copiaID = copiaArchivoPlantilla.getId();
    let nombreDoc = "CO-" + servi + "-" + resConsecutivo + "-" + yyyy + " " + razonSocial + " NIT " + nit;
    let archivo = DriveApp.getFileById(copiaID);
    archivo.setName(nombreDoc);
    let slide = SlidesApp.openById(copiaID);
    let slideLink = slide.getUrl();

    let presentacion = SlidesApp.openById(copiaID);
    presentacion.replaceAllText("{{fecha}}", today);
    presentacion.replaceAllText("{{anio}}", yyyy);
    presentacion.replaceAllText("{{servicio}}", servicio);
    presentacion.replaceAllText("{{numCot}}", "COT-" + servi + "-" + resConsecutivo);
    presentacion.replaceAllText("{{cargo}}", cliCargo);
    presentacion.replaceAllText("{{nombre}}", cliContacto);
    presentacion.replaceAllText("{{razonSocial}}", razonSocial);
    presentacion.replaceAllText("{{numEmp}}", numEmp);
    presentacion.replaceAllText("{{valor}}", pesosTotal);
    presentacion.replaceAllText("{{valorLetras}}", valorletras.toLowerCase());
    presentacion.replaceAllText("{{area}}", area);
    presentacion.replaceAllText("{{numTra}}", numTra);
    presentacion.replaceAllText("{{valPer}}", valPerPesos);
    presentacion.replaceAllText("{{valPerLetras}}", valPerLetras.toLowerCase());
    presentacion.replaceAllText("{{valorLecto}}", empNoLeePesos);
    presentacion.replaceAllText("{{valAnti}}", valAntiPesos);
    presentacion.replaceAllText("{{valAntiLetras}}", valAntiLetras.toLowerCase());

    presentacion.saveAndClose();

    sheetCotizaciones.getRange(lastRowCot + 1, lastColumnCot).setValue(nombreDoc);
    sheetCotizaciones.getRange(lastRowCot + 1, lastColumnCot - 1).setValue(copiaID);
    sheetCotizaciones.getRange(lastRowCot + 1, lastColumnCot - 2).setValue(pdfID);

    let prefilledForm = preFilledForm(total, sheetCotizaciones, lastRowCot, lastColumnCot,servicio);
    sheetCotizaciones.getRange(lastRowCot + 1, 7).setValue(prefilledForm);

    sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(nombreDoc);

    let dataClient1 = htmlData(SSmaestroCot, "Datos", 1, 10);
    let dataClient2 = htmlData(SSmaestroCot, "Datos", 34, 7);
    let dataClient3 = htmlData(SSmaestroCot, "Datos", 47, 3);
    let dataValue = htmlData(SServicio, servicio, 3, 16);
    let dataToSend = dataClient1 + dataClient2 + dataClient3 + dataValue;

    sendEmail(nombreDoc, slideLink, dataToSend,servicio);

  } catch (e) {
    console.error('Error: ', e);
    throw e;
  }
}
