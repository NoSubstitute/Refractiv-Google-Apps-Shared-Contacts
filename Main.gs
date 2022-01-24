//▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅
//Refractiv Google Apps Shared Contacts
//Copyight 2011 Refractiv
//http://www.refractiv.co.uk
//You may use and modify this code but don't remove this credit
//▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅


//http://code.google.com/googleapps/domain/shared_contacts/gdata_shared_contacts_api_reference.html
        
var sAppTitle = "Refractiv Google Apps Shared Contacts";
var oWorkbook = SpreadsheetApp.getActiveSpreadsheet();



    


//------------------------------------------------------------
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ 
    {name: "Setup", functionName: "doSetupUI"}
    ,null
    ,{name: "Upload All Sources", functionName: "doUploadAllSources"}
    ,{name: "Delete All Shared Contacts", functionName: "doDeleteAll"} 
    ,{name: "Download Shared Contacts", functionName: "doListContacts"} 
                    
  ];
  
  ss.addMenu("Contacts Manager", menuEntries);
  
}

//------------------------------------------------------------
function checkSettings(){
  
  //check required settings  
  if (!sGAppsAuthUserName || !sGAppsAuthUserPassword || !sGAppsDomain){  
    doSetupUI();
    return false;
  }else{
    return true;
  }
  
}

//------------------------------------------------------------
//upload each source where Active=yes
function doUploadAllSources(){

  if (!checkSettings()) return false;
  

  var oSheet = oWorkbook.getSheetByName("Sources");
  oSheet.activate();
  
  var iRowsMax = oSheet.getLastRow();

  var iTotal = 0;
  
  for (var iRow = 2; iRow <= iRowsMax; iRow++){
    
    
    //only process if status = yes
    if (oSheet.getRange("F" + iRow).getValue() == "yes"){
      
      var sDocId = oSheet.getRange("B" + iRow).getValue();
      var sSheetName = oSheet.getRange("C" + iRow).getValue();
      
      if (sDocId != ""){
        var ss = SpreadsheetApp.openById(sDocId);
        var sDocName = ss.getName();
      }else{
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sDocName = "[this spreadsheet]";
      }
      
      //set the title
      oSheet.getRange("D" + iRow).setValue(sDocName);
      
      
      //todo - if no sheetname specified, and there's only one, use it and update the sources
    
    
      //todo - check datemodified before uploading
      //Browser.msgBox(DocsList.getFileById(sDocId).getLastUpdated());


      //get the sheet
      var oImportSheet = ss.getSheetByName(sSheetName);
    
      //set the # of rows
      var iRows = oImportSheet.getLastRow();
      //subtract 1 for the header row
      iRows --;
      oSheet.getRange("E" + iRow).setValue(iRows);
      
      
      var now = new Date();
      
      var aColumnIndexes = oSheet.getRange("G" + iRow + ":W" + iRow).getValues();
      
      var sSourceCode = oSheet.getRange("A" + iRow).getValue();
      
      oWorkbook.toast("Uploading " + sSourceCode + "....", sAppTitle, -1);
      
      iTotal += doUploadContacts(oImportSheet, aColumnIndexes, sSourceCode);
      
      var now2 = new Date();
      
      //Logger.log("time:" + ((now2-now)/1000));

      
      
    }
  }
  
  oWorkbook.toast(iTotal + " contacts uploaded, total now " + getContactsCount(), sAppTitle, 2);
  
}
  
//------------------------------------------------------------


function doListContacts() {
  
  
  if (!checkSettings()) return false;
  
  
  var oSheet = oWorkbook.getSheetByName("Test Download");
  var oCell = oSheet.getRange('a1');
  var row = 0;//oSheet.getLastRow();
  
  oSheet.activate();
  
  //delete all but a few rows, and clear them
  var iRowsToDelete = oSheet.getMaxRows() - 6;
  if (iRowsToDelete > 0) oSheet.deleteRows(5, iRowsToDelete);
  oSheet.getRange(2, 1, 5, oSheet.getLastColumn()).clear({contentsOnly:true});  
  
  
  
  var sNextUrl = sContactsApiUrl + "/full";

  //get the total number
  //<openSearch:totalResults xmlns:openSearch="http://a9.com/-/spec/opensearch/1.1/">310</openSearch:totalResults>
  var parsedXml = GoogleHTTPrequest(sNextUrl, "cp", "GET", "", false);
  var sTotal = parsedXml.getElement().getElement("http://a9.com/-/spec/opensearch/1.1/", "totalResults").getText();
    
  if (sTotal == "0"){
    Browser.msgBox(sAppTitle, "No contacts found", Browser.Buttons.OK);
    return false;
      
      
  }else if (Browser.msgBox(sAppTitle, "Are you sure you want to list all " + sTotal + " Shared Contacts?", Browser.Buttons.OK_CANCEL) != "ok"){
    return false;
      
      
  }else{
    
    oWorkbook.toast("Downloading...", sAppTitle, -1);
      
    
    //loop while there's another page of entries
    while (sNextUrl){
      //get a list of entries
      parsedXml = GoogleHTTPrequest(sNextUrl, "cp", "GET", "", false);
      
      //Logger.log(parsedXml.toXmlString());
      
      //get the URL of the next page
      sNextUrl = getChildNodeWithSpecificAttribute(parsedXml, "link", "rel", "next", "href");
      
      
      //get the entries  
      var oElements = parsedXml.getElement().getElements("entry");  
      
      
      //loop through all entries (rows)
      for (var iCount = 0; iCount < oElements.length; iCount++){
        
        row++;
        
        var oElement = oElements[iCount];
        //Logger.log(oElement.toXmlString());
        
        var iCol = 0;
        
        updateCellFromXMLNodeText(oElement, null, "updated", null, oCell.offset(row, iCol++));
        updateCellFromXMLNodeText(oElement, null, "id", null, oCell.offset(row, iCol++));
        
        //names
        var oName = oElement.getElement(sNSurl, "name");
        updateCellFromXMLNodeText(oName, sNSurl, "givenName", null, oCell.offset(row, iCol++));
        updateCellFromXMLNodeText(oName, sNSurl, "familyName", null, oCell.offset(row, iCol++));
        
        var oOrg = oElement.getElement(sNSurl, "organization");
        updateCellFromXMLNodeText(oOrg, sNSurl, "orgName", null, oCell.offset(row, iCol++));
        updateCellFromXMLNodeText(oOrg, sNSurl, "orgTitle", null, oCell.offset(row, iCol++));
        
        
        //cell.offset(row, iCol++).setValue(oElement.getElement("title").getText());
        //oCell.offset(row, iCol++).setValue(oElement.getElement(sNSurl, "email").getAttribute("address").getValue());
        updateCellFromXMLNodeText(oElement, sNSurl, "email", "address", oCell.offset(row, iCol++));
        // + "#work"
        
        //loop through phone numbers and put in the correct columns
        var aPhones = oElement.getElements(sNSurl, "phoneNumber");
        for (var iCount2 = 0; iCount2 < aPhones.length; iCount2++){
          var oPhone = aPhones[iCount2];
          
          var iPhoneIndex;
          if (oPhone.getAttribute("rel").getValue().indexOf("#mobile") > 0) iPhoneIndex = 0;
          if (oPhone.getAttribute("rel").getValue().indexOf("#work") > 0) iPhoneIndex = 1;
          if (oPhone.getAttribute("rel").getValue().indexOf("#home") > 0) iPhoneIndex = 2;
          
          oCell.offset(row, iCol + iPhoneIndex).setValue(oPhone.getText());
        }
        
        iCol = iCol + 3;
        
        //address
        var oAdd = oElement.getElement(sNSurl, "structuredPostalAddress");
        updateCellFromXMLNodeText(oAdd, sNSurl, "street", null, oCell.offset(row, iCol++));
        updateCellFromXMLNodeText(oAdd, sNSurl, "city", null, oCell.offset(row, iCol++)); 
        updateCellFromXMLNodeText(oAdd, sNSurl, "region", null, oCell.offset(row, iCol++)); 
        updateCellFromXMLNodeText(oAdd, sNSurl, "postcode", null, oCell.offset(row, iCol++)); 
        updateCellFromXMLNodeText(oAdd, sNSurl, "country", null, oCell.offset(row, iCol++)); 
        
        
        //notes
        updateCellFromXMLNodeText(oElement, null, "content", null, oCell.offset(row, iCol++));
        //type="text"
        
        //extended Properties
        //updateCellFromXMLNodeText(oElement, sNSurl, "extendedProperty", "name", oCell.offset(row, iCol++)); 
        var sEP = getChildNodeWithSpecificAttribute2(sNSurl, oElement, "extendedProperty", "name", "SourceCode", "value");
        oCell.offset(row, iCol++).setValue(sEP);
          
          
      }
    } 
    
    oWorkbook.toast("Finished", sAppTitle, 1);
  } 
}



//------------------------------------------------------------
