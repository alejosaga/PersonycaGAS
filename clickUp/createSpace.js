async function createSpace(spaceName) {
  try {
    // Obtener los espacios
    var data = await getSpaces();

    // Verificar si el espacio ya existe
    var exists = findSpaceByName(data, spaceName);
    if (exists) {
      console.log(`El espacio "${spaceName}" ya existe.`);
      return null; // Retorna nulo si el espacio ya existe
    }

    // Si no existe, proceder con la creación

    const teamId = '9013247276';
    const url = `https://api.clickup.com/api/v2/team/${teamId}/space`;
    const apiKey = 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z';
    
    const payload = {
      name: spaceName,
      multiple_assignees: true,
      features: {
        due_dates: {
          enabled: true,
          start_date: false,
          remap_due_dates: true,
          remap_closed_due_date: false
        },
        time_tracking: {enabled: false},
        tags: {enabled: true},
        time_estimates: {enabled: true},
        checklists: {enabled: true},
        custom_fields: {enabled: true},
        remap_dependencies: {enabled: true},
        dependency_warning: {enabled: true},
        portfolios: {enabled: true}
      }
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: apiKey
      },
      payload: JSON.stringify(payload)
    };
    
    const resp = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(resp.getContentText());
    const spaceId = responseData.id;
    
    Logger.log(spaceId);
  
    return spaceId
  } catch (error) {
    console.error('Error creating space:', error);
    throw error; // Propaga el error hacia arriba
  }
  }
  ;
  
  async function createFolder(id,contrato){
  const spaceId = id  
  const resp = await fetch(
    `https://api.clickup.com/api/v2/space/${spaceId}/folder`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z'
      },
      body: JSON.stringify({name: contrato})
    }
  );
  
  const data = await resp.json();
  console.log(data);
  return data
  }
  