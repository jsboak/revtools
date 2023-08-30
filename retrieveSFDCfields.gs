function getAccountFields() {

 
  var getAccountFieldsURL = "/services/data/v57.0/sobjects/Account/describe/";
   if(isTokenValid()) {

    var accountFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getAccountFieldsURL,"get","",false);
  }
  var parsedAccountFields = JSON.parse(accountFields);

  return parsedAccountFields.fields;
  
}

function getOpportunityFields() {

  var getOppFieldsURL = "/services/data/v57.0/sobjects/Opportunity/describe/";
   if(isTokenValid()) {

    var oppFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getOppFieldsURL,"get","",false);
  }
  var parsedOppFields = JSON.parse(oppFields);

  return parsedOppFields.fields;

}