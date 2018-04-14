function MasterImport() 
{

var sheetNameMasterImport = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetName(); //in der MasterImport table aktuellen sheet Namen holen
var LastRowVISA2018 = SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getLastRow();
var currentRowVISA2018 = LastRowVISA2018+1;//letzte Zeilennummer mit Inhalt ermitteln 
//SpreadsheetApp.openById("12Z-HfTiZGvB4IDXuRcOxZw1CxFttSLW78N6NA_RDyvY").getSheetByName("VISA 2018").getRange(34, 3).setValue(sheetID); //schreiben der Zeile in...(nur Test)

rowCounter = 2;
while (SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 1).getValue()||//Schleife durch Master Import Table
         SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 2).getValue()|| 
          SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 3).getValue()!=="")//Schleife, welche durch die Tabelle iteriert beginnend bei B1/Prüfung auf vorhandene (Einträge)
  {                                                                                                                 //der Iterator bleibt stehen sobald in B,C und D kein Eintrag vorhanden ist
           // alle Namen umbenennen für bessere Formatierung 
           
           
           
           if (SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="ANNE-SOPHIE RETTIG")
           {
             SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Anne");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="LISA SLUKA")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Lisa");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="DAVID SCHIEBEL")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("David");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="ENRICO GANASSIN")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Enrico");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="FELIX HENSEL")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Felix");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="MELANIE HILGERS")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Melli");
           }
           else if(SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue()=="JENS MEISSNER")
           {
           SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).setValue("Jens");
           }
           
           
           //[Wertstellung] .csv Master Import  zu [Buchungsdatum] Rechnungs Tabelle kopieren an erste verfügbare Zeile nach letztem Eintrag
       
           var spreadsheet = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 3).getValue();
           
           
           SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getRange(currentRowVISA2018, 5).setValue(Wertstellung);
           
           
           
           //[VWZ4] .csv Master Import zu [Belegdatum] Rechnungs Tabelle ] kopieren an erste verfügbare Zeile nach letztem Eintrag
           var VWZ4 = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 9).getValue();
           
           
           
           SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getRange(currentRowVISA2018, 6).setValue(VWZ4);
           
           //[VWZ1] .csv Master Import zu [Firma|Kunde] Rechnungs Tabelle ] kopieren an erste verfügbare Zeile nach letztem Eintrag
           var VWZ1 = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 6).getValue();
           SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getRange(currentRowVISA2018, 3).setValue(VWZ1);
           
           //[VWZ3] .csv Master Import zu [Firma|Kunde] Rechnungs Tabelle ] kopieren an erste verfügbare Zeile nach letztem Eintrag
           var Betrag = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 8).getValue();
           SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getRange(currentRowVISA2018, 7).setValue(Betrag);
           
           //[VWZ1] .csv Master Import zu [Firma|Kunde] Rechnungs Tabelle ] kopieren an erste verfügbare Zeile nach letztem Eintrag
           var Name = SpreadsheetApp.openById("1Rt9AITe28XcveQ-qjmsHRvIYu_cXIy_DSds3JXl9n1U").getSheetByName(sheetNameMasterImport).getRange(rowCounter, 4).getValue();
           SpreadsheetApp.openById("1gRVjgA4fR5Ekd-hX7c0PHDRR21lXnx1jK-oG_-2YJIo").getSheetByName("VISA 2018").getRange(currentRowVISA2018, 2).setValue(Name);
           
           
           
           
           
           currentRowVISA2018++;
           rowCounter++;
  }
  
}





