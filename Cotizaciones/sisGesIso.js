function generarCotizacionISO() {
    
    const SServicio = SpreadsheetApp.openById(isoSisGesCotId);
    const sheetCotizaciones = SServicio.getSheetByName(servicio);
    const lastRowCot = sheetCotizaciones.getLastRow();
    const lastColumnCot = sheetCotizaciones.getLastColumn()
    
     
    // Tarifas según el tipo de servicio
    let tarifaBasica = tarifas[1][6];
    let datNumProcesos = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Indique el número de procesos (áreas o departamentos) que tiene su organización. Ejemplo: Planificación Estratégica, Compras, Comercial, Talento Humano, etc.");     
    let datMaestroDocu = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","La compañía tiene un sistema de gestión documental con un listado maestro de documentos?");
    let datdirTec = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Actualmente la compañía cuenta con un director técnico notificado ante el INVIMA?  ");
    let datdirTecExp = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Cuántos años de experiencia tiene el director técnico  ");
    let datPrevAudi = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","la compañía ha recibido alguna de las siguientes opciones de auditorías?");
   
    
    let caracteristicas = [
      servicio + "NumProcesos",
      servicio + "MaestroDocu" + datMaestroDocu,
      consultoria + "dirTec" + datdirTec,
      consultoria + "dirTecExp" + datdirTecExp,
      servicio + "PrevAudi" + datPrevAudi,
      servicio + "capacitaciones",
      servicio + "Consultor",
      servicio + "costoAdmin",
      servicio + "costoOperativo",
      servicio + "rentPersonyca"

    ];

    let valoresEncontrados = buscarTarifas(caracteristicas);
   

    valoresEncontrados = valoresEncontrados.map((elemento) => elemento * tarifaBasica);
    

    
    let numProcesos = valoresEncontrados[0] * datNumProcesos;   
    let maestroCotIdaestroDocu = valoresEncontrados[1]; 
    let dirTec = valoresEncontrados[2]
    let dirTecExp;
    if (datdirTecExp < 1){
      dirTecExp = valoresEncontrados[3]
    }
    else{
      dirTecExp = 0
    }
    let prevAudi = valoresEncontrados[4];
    let costoCapacitaciones = valoresEncontrados[5]*datNumProcesos;
    let costoConsultor = valoresEncontrados[6]*datNumProcesos;
    let costoAdmin = valoresEncontrados[7]*datNumProcesos;
    let costoOperativo = valoresEncontrados[8]*datNumProcesos;
    


    let totalBruto = numProcesos + maestroCotIdaestroDocu + dirTec + dirTecExp + prevAudi + costoCapacitaciones + costoConsultor + costoAdmin + costoOperativo;
    let rentPersonyca = totalBruto*0.3
    let totalNeto = rentPersonyca + totalBruto
    let vlrMes = totalNeto / 12
    let anticipo = vlrMes * 0.5
    

    
    // Insertar los datos de la cotización en la hoja de cotizaciones
    
    sheetCotizaciones.getRange(lastRowCot + 1, 1).setValue(nit);
    sheetCotizaciones.getRange(lastRowCot + 1, 2).setValue(razonSocial);
    sheetCotizaciones.getRange(lastRowCot + 1, 4).setValue(today);
    sheetCotizaciones.getRange(lastRowCot + 1, 5).setValue(totalNeto);
    sheetCotizaciones.getRange(lastRowCot + 1, 8).setValue(numProcesos);
    sheetCotizaciones.getRange(lastRowCot + 1, 9).setValue(maestroCotIdaestroDocu);
    sheetCotizaciones.getRange(lastRowCot + 1, 10).setValue(dirTec);
    sheetCotizaciones.getRange(lastRowCot + 1, 11).setValue(dirTecExp);
    sheetCotizaciones.getRange(lastRowCot + 1, 12).setValue(prevAudi);
    sheetCotizaciones.getRange(lastRowCot + 1, 13).setValue(costoCapacitaciones);
    sheetCotizaciones.getRange(lastRowCot + 1, 14).setValue(costoConsultor);
    sheetCotizaciones.getRange(lastRowCot + 1, 15).setValue(costoAdmin);
    sheetCotizaciones.getRange(lastRowCot + 1, 16).setValue(costoOperativo);
    sheetCotizaciones.getRange(lastRowCot + 1, 17).setValue(rentPersonyca);
    sheetCotizaciones.getRange(lastRowCot + 1, 18).setValue(vlrMes);
    sheetCotizaciones.getRange(lastRowCot + 1, 19).setValue(anticipo);
    
    addRowNumber(SServicio, servicio, 3);

    let serviPartes = consultoria.split("-");
    let servi = serviPartes[0]
    let carpeta = crearCarpetaCot(servi);
    let folderCotId = carpeta.id;
    let folderCotUrl = carpeta.url;
    sheetCotizaciones.getRange(lastRowCot + 1, 6).setValue(folderCotUrl);

    let slideTemplate = "";
    switch(servi){
      case "ISO 9001 ":
        slideTemplate =  slideISO9001;
        break;
      case "ISO 13485 ":
        slideTemplate = slideISO13485;
        break;
      case "ISO 45001 ":
        slideTemplate = slideISO45001;
        break;
      case "ISO 14001 ":
        slideTemplate = slideISO14001;
        break;
      case "ISO 9001, 45001 y 45001 ":
        slideTemplate = slideIteg1;
        break;
      case "ISO 9001 y 13485 ":
        slideTemplate = slideInteg2;
        break;
        
    }

    
    let ids = getFolderIds(folderCotId);  
    let pdfID = ids.pdf;
    let temporal = ids.doc;
  
    
    //Conexiones
    
    let archivoPlantilla = DriveApp.getFileById(slideTemplate);
    let carpetaPDF = DriveApp.getFolderById(pdfID);
    let carpetaTemporal = DriveApp.getFolderById(temporal);
        
    //Datos Cotizacion.
    let resConsecutivo = sheetCotizaciones.getRange(lastRowCot+1, 3).getValue();
    let vlrNetoPesos = formatoColombiano(totalNeto);
    let vlrMesPesos = formatoColombiano(vlrMes);
    let vlrAntiPesos = formatoColombiano(anticipo);
    
      
    let totalNetoletras = numeroALetras(totalNeto, {
    plural: "PESOS",
    singular: "PESO",
    centPlural: "CENTAVOS",
    centSingular: "CENTAVO"
    });
  
    let vlrMesLetras = numeroALetras(vlrMes, {
    plural: "PESOS",
    singular: "PESO",
    centPlural: "CENTAVOS",
    centSingular: "CENTAVO"
    });  

    let anticipoLetras = numeroALetras(anticipo, {
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
    presentacion.replaceAllText("{{vlrMes}}", vlrMesPesos);
    presentacion.replaceAllText("{{vlrMesLetras}}", vlrMesLetras.toLowerCase());
    presentacion.replaceAllText("{{area}}", area);
    presentacion.replaceAllText("{{numProcesos}}", datNumProcesos);
  


    presentacion.saveAndClose();

    sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot).setValue(nombreDoc);  
    sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-1).setValue(copiaID)
    sheetCotizaciones.getRange(lastRowCot+1,lastColumnCot-2).setValue(pdfID)
  //Datos del cliente y respuestas del form

    let prefilledForm = preFilledForm(vlrMes,sheetCotizaciones,lastRowCot,lastColumnCot,servi);
    sheetCotizaciones.getRange(lastRowCot+1,7).setValue(prefilledForm);

    sheetDatos.getRange(lastRowDat,lastColumnDat).setValue(nombreDoc);

    let dataClient = htmlData(SSmaestroCot,"Datos",1,23);
    let dataValue = htmlData(SServicio,servicio,3,17);
    let dataToSend = dataClient+dataValue;


    sendEmail(nombreDoc,slideLink,dataToSend,consultoria) 
  }
  
  
    
  

