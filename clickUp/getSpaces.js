async function getSpaces(){
    var archived = 'false';
    var teamId = '9013247276';

    var url = `https://api.clickup.com/api/v2/team/${teamId}/space?archived=${archived}`;
    var headers = {
        Authorization: 'pk_72795913_ZB3OQD9YF8WSXP83IM288GNHNCMJLP3Z'
    };

    var options = {
        method: 'GET',
        headers: headers
    };

    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    return data

    
    
}


