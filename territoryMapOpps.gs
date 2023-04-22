function retrieveOpenOpps() {

  Logger.log("Retrieving open opportunities");

  if (currentSfdcUser == null) {
    getCurrentSfdcUser();
  }

  var accountIds = `SELECT+Id+from+Account+WHERE+OwnerId='${currentSfdcUser}'`

  //var opps = `SELECT+Id,+Name,+NextStep+FROM+Opportunity+WHERE+IsClosed+=+FALSE+AND+AccountId+IN+(${accountIds})`; 

  var opps = `SELECT+Opportunity.Id,+Opportunity.Name,+Opportunity.NextStep+FROM+Account.Opportunities+WHERE+IsClosed+=+FALSE`; 
  var soql = `SELECT+name+,+Id+,+(${opps})+from+Account+WHERE+OwnerId='${currentSfdcUser}'`;


  var getDataURL = '/services/data/v57.0/query/?q='+soql;

  var opportunities = salesforceEntryPoint(userProperties.getProperty(baseURLPropertyName) + getDataURL,"get","",true);

  /*
  {"totalSize":3,"done":true,"records":[{"attributes":{"type":"Opportunity","url":"/services/data/v57.0/sobjects/Opportunity/0068Y00001PhioIQAR"},"Id":"0068Y00001PhioIQAR","Name":"Grafana Labs - User Growth"},{"attributes":{"type":"Opportunity","url":"/services/data/v57.0/sobjects/Opportunity/0068Y00001Phim7QAB"},"Id":"0068Y00001Phim7QAB","Name":"RapidAPI - CustomerGPT"},{"attributes":{"type":"Opportunity","url":"/services/data/v57.0/sobjects/Opportunity/0068Y00001PhikWQAR"},"Id":"0068Y00001PhikWQAR","Name":"LogicMonitor - CustomerGPT"}]}
  */

  Logger.log(opportunities);

  Logger.log("Retrieved open opportunities");
  
  return opportunities;
}
