function handleTerritory(e) {
  if(isTokenValid()) {
      try {       
            let nav = CardService.newNavigation().pushCard(goToTerritoryBuilder(e));
              return CardService.newActionResponseBuilder()
              .setNavigation(nav)
              .build();

      }catch(e){
      Logger.log(e);
      }
  } else {
    return onHomepage(e);
  }
}

function goToTerritoryBuilder(e) {

  var title = 'Territory Builder';


  if(isTokenValid()) {

    var sfdcAccountFields = getAccountFields();

    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle(title))
      .addSection(CardService.newCardSection()
      .addWidget(generateFieldsDropdown(sfdcAccountFields, 'accountScore', 'Account Score', ''))
      .addWidget(generateFieldsDropdown(sfdcAccountFields, 'revenue', 'Revenue', ''))
      .addWidget(generateFieldsDropdown(sfdcAccountFields, 'licenseCount', 'License Count', ''))
      .addWidget(generateFieldsDropdown(sfdcAccountFields, 'renewalDate', 'Renewal Date', '')));
    
      builder.addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Create Territory Map')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('createNewTerritoryMap'))
          .setDisabled(false))));
    
  } else {

    builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Authenticate')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOpenLink(CardService.newOpenLink()
            .setUrl(getURLForAuthorization()+"'target='_blank'"))
        .setDisabled(false))));
    return;
  }

  // var cardAction = CardService.newCardAction()
  //   .setText("Go to SFDC Page")
  //   .setOpenLink(CardService.newOpenLink()
  //       .setUrl("https://www.google.com")
  //       .setOpenAs(CardService.OpenAs.OVERLAY)
  //       .setOnClose(CardService.OnClose.NOTHING))

  // builder.addCardAction(cardAction);

  return builder.build();

}

function generateFieldsDropdown(sfdcAccountFields, fieldName, fieldTitle) {
  var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.DROPDOWN);

  sfdcAccountFields.forEach((field, array) => {
    selectionInput.addItem(field.label, field.name, false);
  })

  return selectionInput;
}
