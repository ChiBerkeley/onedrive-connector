function getConfig(request) {
    var config = cc.getConfig();
  
    return config.build();
  };
  
  // Here we assume there a excel file with 3 columns and 4 data rows -> "a", "b", "c"
  function getFields() {
    var fields = cc.getFields();
    var types = cc.FieldType;
    fields
      .newDimension()
      .setId('aId')
      .setName('a')
      .setGroup("stages")
      .setType(types.NUMBER);
  
    fields
      .newDimension()
      .setId('bId')
      .setName('b')
      .setGroup("stages")
      .setType(types.NUMBER);
  
    fields
      .newDimension()
      .setId('cId')
      .setName('c')
      .setGroup("stages")
      .setType(types.NUMBER);
    
    return fields;
  }
   
  function getSchema(request) {
  
    return {schema: getFields().build()};
  }
   
  function isAdminUser() {
    return true;
  }
  
  
  function getData(request) {
    //Ideally we should create a client secret. This is a quick example of using the refresh token
    //https://learn.microsoft.com/en-us/advertising/guides/authentication-oauth-get-tokens?view=bingads-13
    const clientId = "Enter App Regitry Client ID here";
    const refreshToken = "Refresh token"

    //Search the item here you want here if you have a OneDrive or Sharepoint account.
    //In our case, we choose a specific excel file
    //We can also search the entire folder
    //https://developer.microsoft.com/en-us/graph/graph-explorer
    const itemId = 'Enter the item ID here';
    const driveId = 'Enter the Drive ID here';
    const url = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + itemId + '/workbook/worksheets/Sheet1/range(address=\'A1:C5\')';
    
    var data = {
      'client_id': clientId,
      'refresh_token': refreshToken,
      'redirect_uri': 'https://login.live.com/oauth20_desktop.srf',
      'grant_type': 'refresh_token'
    };
    
    var params = {
      'method': 'post',
      'contentType': 'application/x-www-form-urlencoded',
      'payload': data,
    };
  
    var res = UrlFetchApp.fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', params);
    const tokens = JSON.parse(res.getContentText());
    console.log(tokens['access_token']);
    params = {
      headers: { 
        Authorization: 'Bearer '+ tokens['access_token']
      },
    };
  
  
    try{
      var response = UrlFetchApp.fetch(url, params);
  
      var requestedFieldIds = request.fields.map(function(field) {
        return field.name
        });
  
      var issSchema = getSchema(request).schema
  
      var requestedFields = [];
      requestedFieldIds.forEach(function(id) {
          issSchema.forEach(function(field) {
            if (id === field.name) requestedFields.push(field);
          });
        });
  
      var data = getFormattedData(response, requestedFields);
  
      return {
      schema: requestedFields,
      rows: data
    };
    
    } catch (e) {
      cc.newUserError()
        .setDebugText('Error fetching data from API. Exception details: ' + e + '. Data Schema: ' + JSON.stringify(requestedFields) + " Data: " + JSON.stringify(data))
        .setText(
          'The connector has encountered an unrecoverable error. Please try again later, or file an issue if this error persists.'
        )
        .throwException();
    }
  };
  
  
  function getFormattedData(response, dataSchema) {
    const json = JSON.parse(response.getContentText());
    response = json.values
  
  
    var header = response[0];
    var indices = [];
  
    // Find indices of schema fields in header
    dataSchema.forEach(function(item) {
      indices.push(header.indexOf(item.label));
    });
  
    var data = [];
  
    // Process each row of test_case starting from the second row
    for (var i = 1; i < response.length; i++) {
      var row = [];
      indices.forEach(function(index) {
        row.push(response[i][index]);
      });
      data.push({"values": row});
    }
  
    Logger.log(data);  // Prints the data in Apps Script logs
    return data;
  }