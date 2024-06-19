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

    /*
    NumProcesos = valoresEncontrados[0] * datNumProcesos * tarifaBasica;   
    MaestroDocu =valoresEncontrados[1] * tarifaBasica; ; 
    dirTec = 0;
    dirTecExp = 0;
    PrevAudi = 0;
    capacitaciones = 0;
    Consultor = 0;
    costoAdmin = 0;
    costoOperativo = 0;
    rentPersonyca = 0;


    const total = tarifaBasica + costosOperativos + marketing;
  
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
    });*/
  }
  
  
    
  

