async function getTasks() {
  const query = encodeURIComponent('archived=false&include_markdown_description=true&page=0&order_by=string&reverse=true&subtasks=true&statuses=string&include_closed=true&assignees=string&watchers=string&tags=string&due_date_gt=0&due_date_lt=0&date_created_gt=0&date_created_lt=0&date_updated_gt=0&date_updated_lt=0&date_done_gt=0&date_done_lt=0&custom_fields=string&custom_field=string&custom_items=0');

  const listId = '901702393082';
  const url = `https://api.clickup.com/api/v2/list/${listId}/task?${query}`;
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