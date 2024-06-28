function prefillForm() {
    const cotSheetId = cotApproveId;
    const maestroSheetId = maestroCotId;
    const formId = reminderForm;
    
    const cotSheet = SpreadsheetApp.openById(cotSheetId).getSheetByName('Aprobaciones');
    const maestroSheet = SpreadsheetApp.openById(maestroSheetId).getSheetByName('Datos');
    
    if (!cotSheet || !maestroSheet) {
        Logger.log('No se encontró una de las hojas. Asegúrate de que los nombres de las hojas son correctos.');
        return;
      }
      
      const cotDataRange = cotSheet.getDataRange();
      const cotData = cotDataRange.getValues();
      
      const maestroDataRange = maestroSheet.getDataRange();
      const maestroData = maestroDataRange.getValues();
      
      if (cotData.length === 0 || maestroData.length === 0) {
        Logger.log('El rango de datos está vacío en una de las hojas.');
        return;
      }
      
      Logger.log(`Datos de cotizaciones obtenidos: ${JSON.stringify(cotData)}`);
      Logger.log(`Datos maestros obtenidos: ${JSON.stringify(maestroData)}`);
      
      const form = FormApp.openById(formId);
      
      // Elimina todas las preguntas existentes
      const items = form.getItems();
      items.forEach(item => form.deleteItem(item));
      
      // Crear una nueva pregunta de tipo checkbox
      const checkboxItem = form.addCheckboxItem();
      checkboxItem.setTitle('Estas cotizaciones fueron enviadas el último mes, a cuales le envías un recordatorio');
      
      // Obtener las cotizaciones del último mes
      const oneMonthAgo = new Date();
      oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
      
      let options = [];
      let cotizacionesInfo = [];
      
      cotData.forEach((row, index) => {
        if (index === 0) return; // Saltar la fila de encabezado
        const dateSent = new Date(row[0]);
        Logger.log(`Fecha de envío: ${dateSent}`);
        if (dateSent >= oneMonthAgo) {
          const option = `Fecha: ${row[0]}, Cliente: ${row[3]}, Cotización: ${row[4]}`;
          options.push(option);
          Logger.log(`Opción agregada: ${option}`);
          
          // Buscar datos de contacto del cliente en el archivo maestro usando NIT
          const clienteNit = row[2].toString().trim();
          const clienteInfo = maestroData.find(maestroRow => maestroRow[4].toString().trim() == clienteNit); // Asumiendo que la columna E en maestro tiene el NIT del cliente
          
          if (clienteInfo) {
            const cotizacionInfo = {
              fecha: row[0],
              numero: row[4],
              valor: row[5],
              clienteNombre: clienteInfo[6], // Columna F en maestro
              clienteCargo: clienteInfo[8], // Columna H en maestro
              clienteTelefono: clienteInfo[9], // Columna I en maestro
              clienteEmail: clienteInfo[1], // Columna B en maestro
              clienteEmail2: clienteInfo[2] // Columna c en maestro
            };
            cotizacionesInfo.push(cotizacionInfo);
          } else {
            Logger.log(`No se encontró información de contacto para el cliente con NIT: ${clienteNit}`);
          }
        }
      });
      
      Logger.log(`Opciones generadas: ${JSON.stringify(options)}`);
      
      if (options.length === 0) {
        Logger.log('No se encontraron cotizaciones del último mes.');
        return;
      }
      
      try {
        // Añadir nuevas opciones
        checkboxItem.setChoiceValues(options);
        Logger.log('Opciones añadidas al formulario.');
      } catch (e) {
        Logger.log(`Error al añadir opciones: ${e.message}`);
      }
    
      // Obtener la URL pública del formulario
      const formUrl = form.getPublishedUrl();
      Logger.log(`URL del formulario: ${formUrl}`);
      
      return { formUrl, cotizacionesInfo };
    }
    
    function sendWeeklyReminder() {
      
      const { formUrl, cotizacionesInfo } = prefillForm(); // Obtener la URL pública del formulario y la información de las cotizaciones
      
      if (!formUrl || cotizacionesInfo.length === 0) {
        Logger.log('No se pudo obtener la URL del formulario o no hay cotizaciones para enviar.');
        return;
      }
      
      // Crear el cuerpo del correo electrónico
      let emailBody = 'Lista de cotizaciones enviadas durante el último mes:\n\n';
      cotizacionesInfo.forEach(info => {
        emailBody += `Fecha: ${info.fecha}\nNúmero de Cotización: ${info.numero}\nValor: ${info.valor}\n`;
        emailBody += `Cliente: ${info.clienteNombre}\nCargo: ${info.clienteCargo}\nTeléfono: ${info.clienteTelefono}\nEmail: ${info.clienteEmail}\n\n`;
      });
      emailBody += `Puedes seleccionar las cotizaciones para reenviar el recordatorio en el siguiente enlace: ${formUrl}`;
      
     /* MailApp.sendEmail({
        to: personycaEmail1,
        cc: `${personycaEmail2},${personycaEmail3}`,
        subject: 'Recordatorio semanal de cotizaciones',
        body: emailBody,
      })*/;
    }
    
    function setupTrigger() {
      ScriptApp.newTrigger('sendWeeklyReminder')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(9)
        .create();
    }
    