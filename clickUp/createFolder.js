function createClickUpFolder(spaceId, folderName) {
  
  const baseUrl = `https://api.clickup.com/api/v2/space/${spaceId}/folder`;

  // Primero, obtener la lista de carpetas para verificar si ya existe una con el mismo nombre
  const listOptions = {
    method: 'GET',
    headers: {
      'Authorization': API_TOKEN
    },
    muteHttpExceptions: true
  };

  try {
    const listResponse = UrlFetchApp.fetch(baseUrl, listOptions);
    const listStatusCode = listResponse.getResponseCode();
    const listContent = listResponse.getContentText();
    const listData = JSON.parse(listContent);

    if (listStatusCode !== 200) {
      console.error('Error al obtener la lista de carpetas:', listContent);
      return null;
    }

    if (listData.folders) {
      for (let folder of listData.folders) {
        if (folder.name === folderName) {
          console.log(`Carpeta ya existe: ${folderName}`);
          return folder.id;
        }
      }
    }

    // Si no se encuentra ninguna carpeta con el mismo nombre, crear una nueva
    const createOptions = {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        'Authorization': API_TOKEN
      },
      payload: JSON.stringify({
        name: folderName
      }),
      muteHttpExceptions: true
    };

    const createResponse = UrlFetchApp.fetch(baseUrl, createOptions);
    const createStatusCode = createResponse.getResponseCode();
    const createContent = createResponse.getContentText();
    const createData = JSON.parse(createContent);

    if (createStatusCode !== 200 && createStatusCode !== 201) {
      console.error('Error al crear la carpeta:', createContent);
      return null;
    }

    if (createData.id) {
      return createData.id;
    } else {
      console.error('Error en createFolder: No se pudo crear la carpeta.');
      return null;
    }
  } catch (error) {
    console.error('Error en createFolder:', error.message);
    return null;
  }
}
