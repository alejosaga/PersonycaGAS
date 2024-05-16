//View information about a List.
async function getList() {
  const listId = '901702393082';
  const url = `https://api.clickup.com/api/v2/list/${listId}`;
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


