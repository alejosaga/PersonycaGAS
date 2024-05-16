async function getList() {
  const query = encodeURIComponent('archived=false');

  const folderId = '90170869257';
  const url = `https://api.clickup.com/api/v2/folder/${folderId}/list?${query}`;
  const options = {
    method: 'GET',
    headers: {
      Authorization: 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z'
    }
  };

  const resp = UrlFetchApp.fetch(url, options);
  const data = resp.getContentText();
  console.log(data);
}

