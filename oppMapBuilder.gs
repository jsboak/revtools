function handleOpp(e) {
  if(isTokenValid()) {
      try {       
            let nav = CardService.newNavigation().pushCard(goToOppBuilder(e));
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

function goToOppBuilder(e) {

  var title = 'Opp Builder';

  if(isTokenValid()) {

    var sfdcOppFields = getOpportunityFields();

    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle(title))

      builder.addSection(CardService.newCardSection().setHeader("Opps Fields Selection")
      .addWidget(CardService.newDecoratedText().setText("Select the desired fields for your Opportunity report.\nOpportunity Name will be included by default.").setWrapText(true).setIcon(CardService.Icon.DESCRIPTION))
      )
      builder.addSection(CardService.newCardSection()
      .setCollapsible(true)
      .addWidget(generateFieldsSelector(sfdcOppFields, "sfdc_opp_fields", "")
      ))
    
      builder.addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Create Opp Map')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('createNewOppMap'))
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

function generateFieldsSelector(sfdcOppFields, fieldName, fieldTitle) {
  var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.CHECK_BOX);

  Object.keys(sfdcOppFields).sort().
    forEach((function(v, i) {

      if( sfdcOppFields[v].label != "Name") {
        selectionInput.addItem(sfdcOppFields[v].label, v, false);
      }

    }));

  // selectionInput.addItem("Account Name",'{"label":"Account Name","value":"Account.Name","type":"String"}',false)

  return selectionInput;
}

function generateFieldsDropdown(sfdcOppFields, fieldName, fieldTitle) {
  var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.DROPDOWN);

  Object.keys(sfdcOppFields).sort().
    forEach((function(v, i) {

      selectionInput.addItem(sfdcOppFields[v].label, v, false);
    }));

  return selectionInput;
}
