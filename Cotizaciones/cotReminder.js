function prefillForm() {
    const sheetId = '1f5LwW6Ko0o4mUVrhgO5Fa6AiiRTdmENyhx4Oj8r3hK0';
    const formId = '1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c';
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Aprobaciones');
    
    if (!sheet) {
      Logger.log('Hoja no encontrada. Asegúrate de que el nombre de la hoja es correcto.');
      return;
    }
    
    const form = FormApp.openById(formId);
    const items = form.getItems(FormApp.ItemType.CHECKBOX);
    const checkboxItem = items[0].asCheckboxItem();
  
    // Borrar las opciones actuales
    checkboxItem.setChoiceValues([]);
  
    // Obtener las cotizaciones del último mes
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
    const data = sheet.getDataRange().getValues();
  
    if (data.length <= 1) {
      Logger.log('No se encontraron datos o sólo se encontró el encabezado.');
      return;
    }
  
    let options = [];
    data.forEach((row, index) => {
      if (index === 0) return; // Saltar la fila de encabezado
      const dateSent = new Date(row[0]);
      if (dateSent >= oneMonthAgo) {
        options.push(`Fecha: ${row[0]}, Cliente: ${row[3]}, Cotización: ${row[4]}`);
      }
    });
  
    if (options.length === 0) {
      Logger.log('No se encontraron cotizaciones del último mes.');
      return;
    }
  
    // Añadir nuevas opciones
    checkboxItem.setChoiceValues(options);
  }
  
  function sendWeeklyReminder() {
    const emailRecipient = 'autopersonyca@gmail.com';
    const formUrl = 'https://docs.google.com/forms/d/1q0gnfJRANe7t6JEtqlXpvCqm02ooUHoICuRnJ0MeP6c/edit';
  
    prefillForm();
  
    MailApp.sendEmail({
      to: emailRecipient,
      subject: 'Recordatorio semanal de cotizaciones',
      body: `Lista de cotizaciones enviadas durante el último mes. Puedes seleccionar las cotizaciones para reenviar el recordatorio en el siguiente enlace: ${formUrl}`,
    });
  }
  
  function setupTrigger() {
    ScriptApp.newTrigger('sendWeeklyReminder')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
  }
  