function generarCotizacionISO() {
    
    const SServicio = SpreadsheetApp.openById(isoSisGesCotId);
    const sheetCotizaciones = SServicio.getSheetByName(servicio);
    const lastRowCot = sheetCotizaciones.getLastRow();
    const lastColumnCot = sheetCotizaciones.getLastColumn()
    
     
    // Tarifas según el tipo de servicio
    let tarifaBasica = tarifas[1][6];
    let datNumProcesos = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Indique el número de procesos (áreas o departamentos) que tiene su organización. Ejemplo: Planificación Estratégica, Compras, Comercial, Talento Humano, etc.");     
    let datMaestroDocu = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","La compañía tiene un sistema de gestión documental con un listado maestro de documentos?");
    let datdirTec = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Actualmente la compañía cuenta con un director técnico notificado ante el INVIMA?");
    let datdirTecExp = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","Cuántos años de experiencia tiene el director técnico");
    let datPrevAudi = searchValues(maestroCotId,clientCod,"Datos","Codigo Cliente","la compañía ha recibido alguna de las siguientes opciones de auditorías?");
   
    console.log(datNumProcesos);
    console.log(datMaestroDocu);
    console.log(datdirTec);
    console.log(datdirTecExp);
    console.log(datPrevAudi);
     
    
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
    console.log('Valores Encontrados:', valoresEncontrados);

    valoresEncontrados = valoresEncontrados.map((elemento) => elemento * tarifaBasica);
    console.log('Valores Encontrados Ajustados:', valoresEncontrados);

    
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
    console.log(totalNeto)

    /*
    // Insertar los datos de la cotización en la hoja de cotizaciones
    sheetCotizaciones.appendRow([
      nit,
      razonSocial,
      new Date(),
      tarifaBasica,
      costosOperativos,
      marketing,
      total
    ]);
  
    // Enviar el correo con la cotización
    const emailTemplate = HtmlService.createTemplateFromFile('plantilla_email');
    emailTemplate.razonSocial = razonSocial;
    emailTemplate.total = total;
  
    const emailBody = emailTemplate.evaluate().getContent();
  
    MailApp.sendEmail({
      to: 'client@example.com', // Reemplaza con el correo del cliente
      subject: `Cotización de ${servicio}`,
      htmlBody: emailBody
    }); */
  }
  
  
    
  

