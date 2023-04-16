function signOutOfSalesforce() {
  userProperties.deleteProperty('SALESFORCE_OAUTH_TOKEN');
  userProperties.deleteProperty('SALESFORCE_REFRESH_TOKEN');
  userProperties.deleteProperty('SALESFORCE_INSTANCE_URL');


  return onHomepage();

}
