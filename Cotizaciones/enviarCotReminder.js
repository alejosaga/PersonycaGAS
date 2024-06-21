function parseCotizacionString(cotizacionString) {
    const SSreminders = SpreadsheetApp.openById(remidersCotId);
    const shetReminders = SSreminders.getSheetByName(Reminders);
    const lastRowRem = shetReminders.getLastRow();
    const lastColumnRem = shetReminders.getLastColumn();

    let  data = shetReminders.getDataRange(lastRowRem, 2)  // Dividir el string por las comas
    const parts = data.split(',');
  
    // Crear un objeto para almacenar las partes
    const cotizacionData = {};
  
    // Recorrer las partes y dividir por ": " para obtener clave y valor
    parts.forEach(part => {
      const [key, value] = part.split(': ').map(item => item.trim());
      cotizacionData[key] = value;
    });
  
    console.log(cotizacionData);
  }
  
 
  