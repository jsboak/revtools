function getAccountFields() {

  var getAccountFieldsURL = "/services/data/v57.0/sobjects/Account/describe/";
  var accountFields = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getAccountFieldsURL,"get","",false);

  var parsedAccountFields = JSON.parse(accountFields);

  var fieldsList = [];

  for (var i=0; i< parsedAccountFields.fields.length; i++) {

      var item = {};
      item['name'] = parsedAccountFields.fields[i].name
      item['label'] = parsedAccountFields.fields[i].label
      item['type'] = parsedAccountFields.fields[i].type
      fieldsList.push(item);

  }

  //alphabetically sort by label
  fieldsList.sort(function (a, b) {

    if (a.label < b.label) {
      return -1;
    }
    if (a.label > b.label) {
      return 1;
    }
    return 0;

  });

  return fieldsList;
  
}
