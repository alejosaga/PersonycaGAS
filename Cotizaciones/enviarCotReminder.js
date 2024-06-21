function parseCotizacionString() {
    const SSreminders = SpreadsheetApp.openById(remidersCotId);
    const shetReminders = SSreminders.getSheetByName("Reminders");
    const lastRowRem = shetReminders.getLastRow();
  
    let data = shetReminders.getRange(lastRowRem, 2).getValue();  // Obtener el string de cotizaciones
  
    // Dividir el string completo en cotizaciones individuales
    const cotizaciones = data.split(', Fecha:').map((item, index) => index === 0 ? item : 'Fecha:' + item);
  
    let cotizacionesArray = [];
  
    // Recorrer cada cotización y dividir por ": " para obtener clave y valor
    cotizaciones.forEach(cotizacionString => {
      const parts = cotizacionString.split(',');
      const cotizacionData = {};
  
      parts.forEach(part => {
        const [key, value] = part.split(': ').map(item => item.trim());
        cotizacionData[key] = value;
      });
  
      cotizacionesArray.push(cotizacionData);
    });
  
    return cotizacionesArray;
  }
  
  function buscarEnOtroSheet() {
    const cotizaciones = parseCotizacionString();
  
    const otroSheet = SpreadsheetApp.openById(batPsiServiceId).getSheetByName('Aplicacion Bateria riesgo psico');
  
    if (!otroSheet) {
      return;
    }
  
    // Obtener todos los datos del sheet
    const dataRange = otroSheet.getDataRange();
    const data = dataRange.getValues();
  
    cotizaciones.forEach(cotizacion => {
      const cotizacionValue = cotizacion['Cotización'];
  
      // Buscar el valor de la cotización en la columna correspondiente
      let resultado;
      data.forEach(row => {
        if (row.includes(cotizacionValue)) {
          resultado = row;
        }
      });
  
      if (resultado) {
        // Extraer el enlace del formulario
        const formLink = resultado[6]; // Asumiendo que el link del formulario está en la columna G (índice 6)
        Logger.log(`Enlace del formulario: ${formLink}`);
      }
    });
  }