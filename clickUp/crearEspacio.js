function main(spaceName, sheet, row, contrato) {
  try {
    const spaceId = getSpaceId(spaceName);
    if (spaceId) {
      console.log(`Espacio encontrado: ${spaceName} (ID: ${spaceId})`);
      sheet.getRange(row, 9).setValue(spaceId);
      const folderId = createFolder(spaceId, contrato);
      console.log(`ID de la carpeta creada: ${folderId}`);
    } else {
      console.log(`Espacio no encontrado, creando nuevo espacio: ${spaceName}`);
      const newSpaceId = createSpace(spaceName);
      if (newSpaceId) {
        console.log(`Espacio creado exitosamente: ${spaceName} (ID: ${newSpaceId})`);
        sheet.getRange(row, 9).setValue(newSpaceId);
        const folderId = createFolder(newSpaceId, contrato);
        console.log(`ID de la carpeta creada: ${folderId}`);
      } else {
        console.error('No se pudo crear el espacio.');
      }
    }
  } catch (error) {
    console.error('Error en el flujo principal:', error);
  }
}

function getSpaceId(spaceName) {
  const url = `https://api.clickup.com/api/v2/team/${TEAM_ID}/space`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': API_TOKEN
    },
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const space = data.spaces.find(space => space.name === spaceName);
    return space ? space.id : null;
  } catch (error) {
    console.error('Error obteniendo el ID del espacio:', error);
    return null;
  }
}

function createSpace(spaceName) {
  const url = `https://api.clickup.com/api/v2/team/${TEAM_ID}/space`;
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': API_TOKEN
    },
    payload: JSON.stringify({
      name: spaceName,
      multiple_assignees: true,
      features: {
        due_dates: {
          enabled: true,
          start_date: false,
          remap_due_dates: true,
          remap_closed_due_date: false
        },
        time_tracking: { enabled: false },
        tags: { enabled: true },
        time_estimates: { enabled: true },
        checklists: { enabled: true },
        custom_fields: { enabled: true },
        remap_dependencies: { enabled: true },
        dependency_warning: { enabled: true },
        portfolios: { enabled: true }
      }
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (data.id) {
      console.log(`Espacio creado: ${spaceName} (ID: ${data.id})`);
      return data.id;
    } else {
      console.error('Error en la respuesta de creación del espacio:', data);
      return null;
    }
  } catch (error) {
    console.error('Error creando el espacio:', error);
    return null;
  }
}

function createFolder(spaceId, folderName) {
  const url = `https://api.clickup.com/api/v2/space/${spaceId}/folder`;

  const options = {
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

  try {
    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    console.log(`Respuesta de la API al crear la carpeta: ${content}`);
    const data = JSON.parse(content);

    if (data.id) {
      console.log(`Folder creado: ${folderName} (ID: ${data.id})`);
      
      // Creamos la lista dentro de la carpeta
      const listId = createList(spaceId, data.id, "Mi Lista");

      return data.id;
    } else {
      console.error('Error creando el folder:', data);
      return null;
    }
  } catch (error) {
    console.error('Error creando el folder:', error);
    return null;
  }
}
function createList(spaceId, folderId, listName) {
  const url = `https://api.clickup.com/api/v2/folder/${folderId}/list`;

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'Authorization': API_TOKEN
    },
    payload: JSON.stringify({
      name: listName
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    console.log(`Respuesta de la API al crear la lista: ${content}`);
    const data = JSON.parse(content);

    if (data.id) {
      console.log(`Lista creada: ${listName} (ID: ${data.id})`);
      return data.id;
    } else {
      console.error('Error creando la lista:', data);
      return null;
    }
  } catch (error) {
    console.error('Error creando la lista:', error);
    return null;
  }
}