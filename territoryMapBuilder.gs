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

      builder.addSection(CardService.newCardSection().setHeader("Accounts Fields Selection")
      .addWidget(CardService.newDecoratedText().setText("Select the desired fields for your Territory Map.\nAccount Name will be included by default.").setWrapText(true).setIcon(CardService.Icon.DESCRIPTION))
      )
      builder.addSection(CardService.newCardSection()
      .setCollapsible(true)
      .addWidget(generateTerritoryFieldsSelector(sfdcAccountFields, "sfdc_territory_fields", "")
      ))

      builder.addSection(CardService.newCardSection()
        .addWidget(CardService.newDecoratedText()
        .setText("Include Open Opportunity")
        .setBottomLabel("Column will show associated open opportunity.")
        .setWrapText(true)
        .setSwitchControl(CardService.newSwitch()
            .setFieldName("include_open_opp_key")
            .setValue("include_open_opp_value")
            .setSelected(true)
          )
        )
      )
    
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

function generateTerritoryFieldsSelector(sfdcAccountFields, fieldName, fieldTitle) {
  var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.CHECK_BOX);

  sfdcAccountFields.sort(function (a, b) {
      return a.label.localeCompare(b.label);
  });

  for (var i=0; i< sfdcAccountFields.length; i++) {

    if(sfdcAccountFields[i]["label"].trim() != "Account Name") {

      selectionInput.addItem(sfdcAccountFields[i]["label"], sfdcAccountFields[i].name + ":" + sfdcAccountFields[i].label, false);
    }

  }

  return selectionInput;
}

// function generateFieldsDropdown(sfdcAccountFields, fieldName, fieldTitle) {
//   var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
//     .setFieldName(fieldName)
//     .setType(CardService.SelectionInputType.DROPDOWN);

//   Object.keys(sfdcAccountFields).sort().
//     forEach((function(v, i) {

//       selectionInput.addItem(sfdcAccountFields[v].label, v, false);
//     }));

//   return selectionInput;
// }
