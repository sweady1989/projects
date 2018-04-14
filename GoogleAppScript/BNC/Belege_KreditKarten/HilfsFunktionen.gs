//GmailApp.sendEmail("sweady_feet@hotmail.de", "Erinnerung RechnungTest", "Hallo Rico, denkst du bitte noch an den Nachweis ?", {from: "lisasluka@googlemail.com"});
//GmailApp.sendEmail("sweady_feet@hotmail.de", "Erinnerung RechnungTest", "Hallo Rico, denkst du bitte noch an den Nachweis ?", options)



function clearAllHelperColumns()
{

var rowCounter = 2;
while (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue()||
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 3).getValue()|| 
           SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 4).getValue()!=="")
{
           
           SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 11).clearContent();
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 11).setBackgroundRGB(255,255,255);
             SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 12).clearContent();
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 12).setBackgroundRGB(255,255,255);
         
  
rowCounter++;
}



}

