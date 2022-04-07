/** @OnlyCurrentDoc */

function Mail_emprunt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Mail et date');
  var sheet2 = ss.getSheetByName('Déclaration emprunt');
  var emailAddress = sheet2.getRange(21,2).getValue();
  var sujet = sheet1.getRange(2,2).getValue();
  var message = sheet1.getRange(2,3).getValue();
  MailApp.sendEmail(emailAddress, sujet, message);
}

function Mail_prevenir() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Mail et date');
  var emailAddress = ("matthieu.cordon@dewesoft.com");
  var sujet = sheet1.getRange(3,2).getValue();
  var message = sheet1.getRange(3,3).getValue();
  MailApp.sendEmail(emailAddress, sujet, message)
}

function Mail_prevenir_fin_emprunt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Mail et date');
  var emailAddress = ("vincent.vs94200@gmail.com");
  var sujet = sheet1.getRange(4,2).getValue();
  var message = sheet1.getRange(4,3).getValue();

  MailApp.sendEmail(emailAddress, sujet, message)
}

function Mail_fin_emprunt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Mail et date');
  var sheet2 = ss.getSheetByName('Rendre matériel')
  var emailAddress = sheet2.getRange(20,2).getValue()
  var sujet = sheet1.getRange(5,2).getValue();
  var message = sheet1.getRange(5,3).getValue();

  MailApp.sendEmail(emailAddress, sujet, message)
}

function deletestock(){   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Matériel en emprunt');
  var r = s.getRange('A:A');
  var v = r.getValues();
  var valeur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rendre matériel').getRange('B9').getValue();
  var feuille = SpreadsheetApp.getActive();
  var cel1 = feuille.getRange('\'Rendre matériel\'!B9');
  var cel2 = feuille.getRange('\'Rendre matériel\'!B14');
  var cel3 = feuille.getRange('\'Rendre matériel\'!B17');
  var cel4 = feuille.getRange('\'Rendre matériel\'!B20');
  var result = feuille.getRange('\'Rendre matériel\'!B7');

  if(cel1.isBlank()){
    result.setValue("Vous devez renseigner le S/N !");
  }
  else if (cel2.isBlank()){
    result.setValue("Vous devez renseigner votre prénom !");
  }
  else if (cel3.isBlank()){
    result.setValue("Vous devez renseigner la marque et le modèle !");
  }
  else if (cel4.isBlank()){
    result.setValue("Vous devez renseigner votre adresse mail !");
  }
  else 
  {
  for(var i=v.length-1;i>=0;i--)
    if(v[0,i]== valeur){
      s.deleteRow(i+1);
      result.setValue("OK")
      feuille.getRange('B9').activate();
      feuille.setActiveSheet(feuille.getSheetByName('Rendre matériel'), true);
      feuille.getRange('B9').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      feuille.getRange('B11').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      feuille.getRange('B14').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      feuille.getRange('B17').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      feuille.getRange('B20').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      feuille.getRange('B7').activate();
      feuille.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      Mail_fin_emprunt();
      Mail_prevenir_fin_emprunt();
    }
    else{
      result.setValue("Un mail vous a été envoyer pour vous confirmer si votre retour a bien été effectué");
    }
  }
}

function entree_stock() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cel1 = ss.getRange('\'Déclaration emprunt\'!B9');
  var result = ss.getRange('\'Déclaration emprunt\'!B7');
  var cel2 = ss.getRange('\'Déclaration emprunt\'!B15');
  var cel3 = ss.getRange('\'Déclaration emprunt\'!B18');
  var cel4 = ss.getRange('\'Déclaration emprunt\'!B21');
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  
  if(cel1.isBlank()){
    result.setValue("Vous devez renseigner le S/N !");
  }
  else if(cel2.isBlank()){
    result.setValue("Vouse devez renseigner la marque et le modèle !");
  }
  else if(cel3.isBlank()){
    result.setValue("Vouse devez renseigner votre prénom !");
  }
  else if(cel4.isBlank()){
    result.setValue("Vouse devez renseigner votre adresse mail !");
  }
  else{
    Mail_prevenir();
    Mail_emprunt();
    result.setValue("OK");
    spreadsheet.getRange('B9').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Matériel en emprunt'), true);
    spreadsheet.getRange('7:7').activate();
    spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
    spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    spreadsheet.getRange('A7').activate();
    spreadsheet.getRange('\'Déclaration emprunt\'!B9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Déclaration emprunt'), true);
    spreadsheet.getRange('B12').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Matériel en emprunt'), true);
    spreadsheet.getRange('B7').activate();
    spreadsheet.getRange('\'Déclaration emprunt\'!B12').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Déclaration emprunt'), true);
    spreadsheet.getRange('B15').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Matériel en emprunt'), true);
    spreadsheet.getRange('C7').activate();
    spreadsheet.getRange('\'Déclaration emprunt\'!B15').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Déclaration emprunt'), true);
    spreadsheet.getRange('B18').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Matériel en emprunt'), true);
    spreadsheet.getRange('D7').activate();
    spreadsheet.getRange('\'Déclaration emprunt\'!B18').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Déclaration emprunt'), true);
    spreadsheet.getRange('B21').activate();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Matériel en emprunt'), true);
    spreadsheet.getRange('F7').activate();
    spreadsheet.getRange('\'Déclaration emprunt\'!B21').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.getRange('7:7').activate();
    spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#f05c2e', SpreadsheetApp.BorderStyle.DOUBLE)
    .setBackground('#232323')
    .setFontColor('#f05c2e');
    spreadsheet.getRange('\'Matériel en emprunt\'!E7').setValue(date);
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Déclaration emprunt'), true);
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('B18').activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('B15').activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('B12').activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('B9').activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('B7').activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    result.setValue("Un mail vous a été envoyer pour vous confirmer si votre retour a bien été effectué");
  }
};
