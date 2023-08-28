function getAccountFields() {

 
  var getAccountFieldsURL = "/services/data/v57.0/sobjects/Account/describe/";
   if(isTokenValid()) {

    var accountFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getAccountFieldsURL,"get","",false);
  }
  var parsedAccountFields = JSON.parse(accountFields);

  var fieldsList = {};

  for (var i=0; i< parsedAccountFields.fields.length; i++) {

      fieldsList[parsedAccountFields.fields[i].name] = {"label":parsedAccountFields.fields[i].label, "type":parsedAccountFields.fields[i].type}

  }

  return fieldsList;
  
}

function getOpportunityFields() {

  var getOppFieldsURL = "/services/data/v57.0/sobjects/Opportunity/describe/";
   if(isTokenValid()) {

    var oppFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getOppFieldsURL,"get","",false);
  }
  var parsedOppFields = JSON.parse(oppFields);

  var fieldsList = {};

  for (var i=0; i< parsedOppFields.fields.length; i++) {

      fieldsList[parsedOppFields.fields[i].name] = {"label":parsedOppFields.fields[i].label, "type":parsedOppFields.fields[i].type}

  }

  return fieldsList;


}