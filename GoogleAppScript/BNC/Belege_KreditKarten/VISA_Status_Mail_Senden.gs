function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  spreadsheet.addMenu(
    "Skripte",
    [
      {name: "VISA 2018 Master IMPORT", functionName: "MasterImport"},
      {name: "VISA 2018 Rechnungen STATUS", functionName: "AutomatedBillMemory"},  //Menu im Dokument "Skripte" -> RechnungsErinnerungenSendenVISA2018
      {name: "VISA 2018 Mails SENDEN", functionName:"AutomatedMailSending"},
      {name: "VISA 2018 Spalten LÖSCHEN", functionName: "clearAllHelperColumns"},
      
    ]
  )
      
      //Methode zum iterien der Tabelle und Statusabfrage
}
function AutomatedBillMemory()
{
  
  var rowCounter = 2; 
      
      
  while (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue()||
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 3).getValue()|| 
           SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 4).getValue()!=="")//Schleife, welche durch die Tabelle iteriert beginnend bei B2/Prüfung auf vorhandene Namen(Einträge)
  {                                                                                                                 //der Iterator bleibt stehen sobald in B,C und D kein Eintrag vorhanden ist
      
      
      
      
      
      
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Jens") //wenn Eintrag mit [Name] vorhanden, dann
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
  else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Melli")
  {
     
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
  else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "David")
  {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
    else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Anne")
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
    else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Felix")
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
    else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Max")
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
    else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Lisa")
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    }
    else  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 2).getValue() == "Enrico")
    {
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("✅");  //setze grünen Haken in Spalte "K" bzw. Zelle 10 auf X-Achse, wenn kein Eintrag, dann "keinen Namen gefunden" (else Zweig) 
      
       if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter, 10).getValue()!== "X"&&
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="") //wenn "X" fehlt und die Zelle leer ist ergo keine Belegvorhanden ist
       {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg fehlt"); //schreibe "Beleg fehlt" in Spalte "L"
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255, 0, 0);
       }
       else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!==""&&
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()!=="X") //oder iwas ungleich NULL und ungleich "X"steht
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("anderer Beleg/Eintrag");
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setBackgroundRGB(255,204,0);
      }
      else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,10).getValue()=="X") //oder Belegvorhanden
      {
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,12).setValue("Beleg✅");
           
      }
      
      
    
    }
    else
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setValue("keinen Namen gefunden");
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter,11).setBackgroundRGB(255, 0, 0);
    }
    rowCounter++;
    }
    
     
      
}   
   
      
   
      
      
      
      
      // Methode zum iterien der Tabelle und automatisierten mailSendung zwecks nicht vorhandenen Rechnungen
 function AutomatedMailSending()
 {
    var rowCounter2 = 2;
    //ArrayBlöcke für Rechnungsspalten jeweiliger Mitarbeiter
    //JensBlock
    var JensFirmaArray = [];
    var JensProduktionArray = [];
    var JensBelegDArray = [];
    var JensBetragArray = [];
    //EnricoBlock
    var EnricoFirmaArray = [];
    var EnricoProduktionArray = [];
    var EnricoBelegDArray = [];
    var EnricoBetragArray = [];
    //LisaBlock
    var LisaFirmaArray = [];
    var LisaProduktionArray = [];
    var LisaBelegDArray = [];
    var LisaBetragArray = [];
    //AnneBlock
    var AnneFirmaArray = [];
    var AnneProduktionArray = [];
    var AnneBelegDArray = [];
    var AnneBetragArray = [];
    //MelliBlock
    var MelliFirmaArray = [];
    var MelliProduktionArray = [];
    var MelliBelegDArray = [];
    var MelliBetragArray = [];
    //FelixBlock
    var FelixFirmaArray = [];
    var FelixProduktionArray = [];
    var FelixBelegDArray = [];
    var FelixBetragArray = [];
    //DavidBlock
    var DavidFirmaArray = [];
    var DavidProduktionArray = [];
    var DavidBelegDArray = [];
    var DavidBetragArray = [];
    //MaxBlock
    var MaxFirmaArray = [];
    var MaxProduktionArray = [];
    var MaxBelegDArray = [];
    var MaxBetragArray = [];
    
    
    //gehe solange durch die Schleife wie in den ersten 3 Spalten von VISA 2018 min. ein Wert steht
    while (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2, 2).getValue()||
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2, 3).getValue()|| 
             SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2, 4).getValue()!=="")
    {
         if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Jens gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Jens") 
         {
           JensFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           JensProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           JensBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           JensBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
                                                                                                                                         //bei jedem durchlauf wird ein neuer EIntrag des jeweiligen Elements ans Ende des Arrays angehangen
         }
         else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Enrico gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Enrico") 
         {
           EnricoFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           EnricoProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           EnricoBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           EnricoBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }
         else if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Lisa gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Lisa") 
         {
           LisaFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           LisaProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           LisaBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           LisaBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }   
         else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Anne gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Anne") 
         {
           AnneFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           AnneProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           AnneBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           AnneBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         } 
         else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Melli gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Melli") 
         {
           MelliFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           MelliProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           MelliBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           MelliBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }  
         else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Felix gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Felix") 
         {
           FelixFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           FelixProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           FelixBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           FelixBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }  
         else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile David gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="David") 
         {
           DavidFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           DavidProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           DavidBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           DavidBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }  
         else if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,12).getValue()=="Beleg fehlt"&& //wenn Beleg fehlt & die Zeile Max gehört, dann 
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,2).getValue()=="Max") 
         {
           MaxFirmaArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,3).getValue());     //schreibe FirmaEintrag in Firma Array
           MaxProduktionArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,4).getValue()); //schreibe Produktion Eintrag in Produktion Array
           MaxBelegDArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,6).getValue());   //schreibe Beleg Eintrag in Beleg Array
           MaxBetragArray.push(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VISA 2018").getRange(rowCounter2,7).getValue());   //schreibe Betrag Eintrag in Betrag Array
         }  
         rowCounter2++;
    }
          
          
          //for Schleifen Blöcke um fehlende Rechnungseinträge in die PersonenSheets zu kopieren 
          
          //Jens for Schleifen Block 
          
          for (var i=0; i<JensFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 1).setValue(JensFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<JensProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 2).setValue(JensProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<JensBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 3).setValue(JensBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<JensBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 4).setValue(JensBetragArray[i]);
          
          j++;
          }
          //Enrico for Schleifen Block 
          
          for (var i=0; i<EnricoFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 1).setValue(EnricoFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<EnricoProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 2).setValue(EnricoProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<EnricoBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 3).setValue(EnricoBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<EnricoBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 4).setValue(EnricoBetragArray[i]);
          
          j++;
          }
          //Lisa for Schleifen Block 
          
          for (var i=0; i<LisaFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 1).setValue(LisaFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<LisaProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 2).setValue(LisaProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<LisaBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 3).setValue(LisaBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<LisaBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 4).setValue(LisaBetragArray[i]);
          
          j++;
          }
          //Anne for Schleifen Block 
          
          for (var i=0; i<AnneFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 1).setValue(AnneFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<AnneProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 2).setValue(AnneProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<AnneBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 3).setValue(AnneBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<AnneBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 4).setValue(AnneBetragArray[i]);
          
          j++;
          }
          //Melli for Schleifen Block 
          
          for (var i=0; i<MelliFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 1).setValue(MelliFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MelliProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 2).setValue(MelliProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MelliBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 3).setValue(MelliBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MelliBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 4).setValue(MelliBetragArray[i]);
          
          j++;
          }
          //Felix for Schleifen Block 
          
          for (var i=0; i<FelixFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 1).setValue(FelixFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<FelixProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 2).setValue(FelixProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<FelixBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 3).setValue(FelixBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<FelixBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 4).setValue(FelixBetragArray[i]);
          
          j++;
          }
          //David for Schleifen Block 
          
          for (var i=0; i<DavidFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 1).setValue(DavidFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<DavidProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 2).setValue(DavidProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<DavidBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 3).setValue(DavidBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<DavidBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 4).setValue(DavidBetragArray[i]);
          
          j++;
          }
          //Max for Schleifen Block 
          
          for (var i=0; i<MaxFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 1).setValue(MaxFirmaArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MaxProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 2).setValue(MaxProduktionArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MaxBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 3).setValue(MaxBelegDArray[i]);
          
          j++;
          }
          
            for (var i=0; i<MaxBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 4).setValue(MaxBetragArray[i]);
          
          j++;
          }
     
          // if Anweisungen, die prüfen ob in den Mitarbeitersheets Einträge stehen - wenn ja, dann die jeweilige Tabelle als mail senden 
          
          if (SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA"); //mail schicken 
          GmailApp.sendEmail("meissner@bnc-online.tv", "fehlende Visa Belege", "Lieber Jens,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc");  //mail schicken 
          GmailApp.sendEmail("ganassin@bnc-online.tv", "fehlende Visa Belege", "Lieber Enrico,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg");  //mail schicken 
          GmailApp.sendEmail("sluka@bnc-online.tv", "fehlende Visa Belege", "Liebe Lisa,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM");  //mail schicken 
          GmailApp.sendEmail("rettig@bnc-online.tv", "fehlende Visa Belege", "Liebe Anne,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ");  //mail schicken 
          GmailApp.sendEmail("schueller@bnc-online.tv", "fehlende Visa Belege", "Lieber Alex,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs");  //mail schicken 
          GmailApp.sendEmail("hensel@bnc-online.tv", "fehlende Visa Belege", "Lieber Felix,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY");  //mail schicken 
          GmailApp.sendEmail("schiebel@bnc-online.tv", "fehlende Visa Belege", "Lieber David,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          if (SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(2, 1).getValue()!=="")//prüfen ob Einträge in Mitarbeitertabelle vorhanden, wenn ja 
          {
              
          var file = DriveApp.getFileById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo");  //mail schicken 
          GmailApp.sendEmail("kilian@bnc-online.tv", "fehlende Visa Belege", "Lieber Max,\n \n im Anhang findest du deine Übersicht zu den fehlenden VISA Belegen.\n\nBitte lade sie zeitnah hoch.\n\nLiebe Grüße Lisa",{attachments:[file]});
         
          }
          
          //for Schleifen Block zum löschen der Rechnungseinträge in den Personensheets
          
          for (var i=0; i<JensFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<JensProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<JensBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<JensBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gtJUo6RkhgUl4WJ-b9h0k7UNontWDgxjxi2fseMwPFA").getSheetByName("Jens").getRange(j, 4).clearContent();
          
          j++;
          }
          //Enrico for Schleifen Block 
          
          for (var i=0; i<EnricoFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<EnricoProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 2).clearContent();
          j++;
          }
          
            for (var i=0; i<EnricoBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<EnricoBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1eGve3wseMYZlXG0QMhhNYj0By4LPZE_tqy2WLZnZ_Tc").getSheetByName("Enrico").getRange(j, 4).clearContent();
          
          j++;
          }
          //Lisa for Schleifen Block 
          
          for (var i=0; i<LisaFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<LisaProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<LisaBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<LisaBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("131-5EeQ78Uy-6XRfeoSz2qNY_dMRraZ-En-iSDH_SRg").getSheetByName("Lisa").getRange(j, 4).clearContent();
          
          j++;
          }
          //Anne for Schleifen Block 
          
          for (var i=0; i<AnneFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<AnneProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<AnneBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<AnneBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Lv5x7tZmunDjsLl_ll-uFzsLhBA-jG_SwXMzwWFmCbM").getSheetByName("Anne").getRange(j, 4).clearContent();
          
          j++;
          }
          //Melli for Schleifen Block 
          
          for (var i=0; i<MelliFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<MelliProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 2).clearContent();;
          
          j++;
          }
          
            for (var i=0; i<MelliBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<MelliBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gx4lw-zaFK2kMJlEc2ye8ckjC5fUhZdJn_jIVTULXxQ").getSheetByName("Melli").getRange(j, 4).clearContent();
          
          j++;
          }
          //Felix for Schleifen Block 
          
          for (var i=0; i<FelixFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<FelixProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<FelixBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<FelixBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1YrBXcostSrMNm5QWByFNLnrI1rM6a58IcXHITbiRPjs").getSheetByName("Felix").getRange(j, 4).clearContent();
          
          j++;
          }
          //David for Schleifen Block 
          
          for (var i=0; i<DavidFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<DavidProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<DavidBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<DavidBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1gW_XtphgzFhU6SS_oT_-l89jM7QYFm9ukTN1BuFDJxY").getSheetByName("David").getRange(j, 4).clearContent();
          
          j++;
          }
          //Max for Schleifen Block 
          
          for (var i=0; i<MaxFirmaArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 1).clearContent();
          
          j++;
          }
          
            for (var i=0; i<MaxProduktionArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 2).clearContent();
          
          j++;
          }
          
            for (var i=0; i<MaxBelegDArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 3).clearContent();
          
          j++;
          }
          
            for (var i=0; i<MaxBetragArray.length; i++)
          {
          var j = i+2;
          SpreadsheetApp.openById("1Circe9op_S8n-TVI8UQg0-WGM1a_dW3oJr7qVFVzfSo").getSheetByName("Max").getRange(j, 4).clearContent();
          
          j++;
          }
       
          
          
   }
          
 
          
          
          


  
  
  





