const API_TOKEN = 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z'; // Reemplaza con tu token de ClickUp
const TEAM_ID = '9013247276'; // Reemplaza con el ID de tu equipo en ClickUp

function main(spaceName,folderName,listName) {
  const spaceName = 'NombreDelEspacio';
  const folderName = 'NombreDelFolder';
  const listName = 'NombreDeLaLista';

  const spaceId = getSpaceId(spaceName);

  if (spaceId) {
    createFolderAndList(spaceId, folderName, listName);
  } else {
    const newSpaceId = createSpace(spaceName);
    createFolderAndList(newSpaceId, folderName, listName);
  }
}

function getSpaceId(spaceName) {
  const url = `https://api.clickup.com/api/v2/team/${TEAM_ID}/space`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': API_TOKEN,
    },
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  const space = data.spaces.find(space => space.name === spaceName);
  return space ? space.id : null;
}

function createSpace(spaceName) {
  const url = `https://api.clickup.com/api/v2/team/${TEAM_ID}/space`;

  const options = {
    method: 'POST',
    headers: {
      'Authorization': API_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      name: spaceName,
      multiple_assignees: true,
      features: {
        due_dates: {
          enabled: true,
          start_date: true,
          remap_due_dates: true,
          remap_closed_due_date: true,
        },
        time_tracking: {
          enabled: true,
        },
        tags: {
          enabled: true,
        },
        time_estimates: {
          enabled: true,
        },
        check_unresolved: {
          enabled: true,
        },
        sprints: {
          enabled: true,
        },
        custom_fields: {
          enabled: true,
        },
        dependency_warning: {
          enabled: true,
        },
        multiple_assignees: {
          enabled: true,
        },
      },
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  return data.id;
}

function createFolderAndList(spaceId, folderName, listName) {
  const folderId = createFolder(spaceId, folderName);
  createList(folderId, listName);
}

function createFolder(spaceId, folderName) {
  const url = `https://api.clickup.com/api/v2/space/${spaceId}/folder`;

  const options = {
    method: 'POST',
    headers: {
      'Authorization': API_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      name: folderName,
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  return data.id;
}

function createList(folderId, listName) {
  const url = `https://api.clickup.com/api/v2/folder/${folderId}/list`;

  const options = {
    method: 'POST',
    headers: {
      'Authorization': API_TOKEN,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      name: listName,
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  return data.id;
}
