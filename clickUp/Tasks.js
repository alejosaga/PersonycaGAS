function crearTareasEnClickUp(listId) {
  var apiKey = API_TOKEN;
  let ssTasksId = SpreadsheetApp.openById(tasksSgsst);
  let sheetTasks = ssTasksId.getSheetByName("Tasks");
  var data = sheetTasks.getRange(2, 1, sheetTasks.getLastRow() - 1, sheetTasks.getLastColumn()).getValues();

  for (var i = 1; i < data.length; i++) {  // Cambiado de i=1 a i=0
    var nombre = data[i][0];
    var descripcion = data[i][1];
    var estado = data[i][2];
    var fechaVencimiento = new Date(data[i][3]).getTime();

    // Reemplazar estado por uno válido si no es reconocido
    var estadosValidos = ["to do", "in progress", "complete"]; // Ajusta según los estados válidos obtenidos
    if (!estadosValidos.includes(estado.toLowerCase())) {
      estado = "to do"; // Estado por defecto en caso de no coincidencia
    }

    var payload = {
      "name": nombre,
      "description": descripcion,
      "status": estado,
      "due_date": fechaVencimiento
    };

    var options = {
      'method': 'POST',
      'headers': {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };

    var url = 'https://api.clickup.com/api/v2/list/' + listId + '/task';

    try {
      var response = UrlFetchApp.fetch(url, options);
      Logger.log(response.getContentText());
    } catch (e) {
      Logger.log('Error: ' + e.message);
    }
  }
}
