async function createSpace(spaceName) {
  try {
    // Get spaces
    const data = await getSpaces();

    // Check if space exists
    var spaceId = findSpaceByName(data, spaceName);
    if (spaceId) {
      console.log(`El espacio "${spaceName}" ya existe.`);
      return spaceId;
    }

    // Create space if not found
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

    Logger.log(`Space created successfully. ID: ${spaceId}`); // More specific logging

    return spaceId;
  } catch (error) {
    console.error('Error creating space:', error);
    console.error('Error details:', error.message || error); // Log more details
    throw error;
  }
}

async function createFolder(id, contrato) {
  const spaceId = id;
  const resp = await UrlFetchApp.fetch(
    `https://api.clickup.com/api/v2/space/${spaceId}/folder`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z'
      },
      payload: JSON.stringify({ name: contrato })
    }
  );

  const data = await resp.json();
  console.log(data);
  return data;
}
