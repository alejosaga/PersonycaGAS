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