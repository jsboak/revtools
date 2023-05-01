function getAccountFields() {

  var getAccountFieldsURL = "/services/data/v57.0/sobjects/Account/describe/";
  var accountFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getAccountFieldsURL,"get","",false);

  var parsedAccountFields = JSON.parse(accountFields);

  var fieldsList = {};

  for (var i=0; i< parsedAccountFields.fields.length; i++) {

      fieldsList[parsedAccountFields.fields[i].name] = {"label":parsedAccountFields.fields[i].label, "type":parsedAccountFields.fields[i].type}

  }

  return fieldsList;
  
}
