function main(spaceName, sheet, row, contrato) {
  try {
    const spaceId = getSpaceId(spaceName);
    if (spaceId) {
      sheet.getRange(row, 9).setValue(spaceId);
      const folderId = createFolder(spaceId, contrato);
      if (folderId) {
        const listId = createList(spaceId, folderId, "Mi Lista");
        console.log(`ID de la lista creada: ${listId}`);
        return listId;
      }
    } else {
      const newSpaceId = createSpace(spaceName);
      if (newSpaceId) {
        sheet.getRange(row, 9).setValue(newSpaceId);
        const folderId = createFolder(newSpaceId, contrato);
        if (folderId) {
          const listId = createList(newSpaceId, folderId, "Mi Lista");
          console.log(`ID de la lista creada: ${listId}`);
          return listId;
        }
      }
    }
  } catch (error) {
    console.error('Error en main:', error);
    return null;
  }
  return null;
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
    console.error('Error en getSpaceId:', error);
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
      return data.id;
    } else {
      console.error('Error en createSpace: No se pudo crear el espacio.');
      return null;
    }
  } catch (error) {
    console.error('Error en createSpace:', error);
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
    const data = JSON.parse(content);
    if (data.id) {
      return data.id;
    } else {
      console.error('Error en createFolder: No se pudo crear la carpeta.');
      return null;
    }
  } catch (error) {
    console.error('Error en createFolder:', error);
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
    const data = JSON.parse(content);
    if (data.id) {
      return data.id;
    } else {
      console.error('Error en createList: No se pudo crear la lista.');
      return null;
    }
  } catch (error) {
    console.error('Error en createList:', error);
    return null;
  }
}
