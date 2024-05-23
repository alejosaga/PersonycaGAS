function main(spaceName, sheet, row, contrato) {
  try {
    const spaceId = getSpaceId(spaceName);
    if (spaceId) {
      
      sheet.getRange(row, 9).setValue(spaceId);
      const folderId = createFolder(spaceId, contrato);
      
    } else {
      
      const newSpaceId = createSpace(spaceName);
      if (newSpaceId) {
        
        sheet.getRange(row, 9).setValue(newSpaceId);
        const folderId = createFolder(newSpaceId, contrato);
        console.log(`ID de la carpeta creada: ${folderId}`);
      } else {
        
      }
    }
  } catch (error) {
    
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
      
      return null;
    }
  } catch (error) {
    
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
      
      
      // Creamos la lista dentro de la carpeta
      const listId = createList(spaceId, data.id, "Mi Lista");

      return data.id;
    } else {
      
      return null;
    }
  } catch (error) {
    
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
      
      return null;
    }
  } catch (error) {
    
    return null;
  }
}