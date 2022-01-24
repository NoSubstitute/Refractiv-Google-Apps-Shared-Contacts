/*
▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅
refractiv shared functions for Google Apps Shared Contacts
version 1.0
▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅▅*/


var sContactsApiUrl;


var bUseBatch = true;
var iBatchSize = 95; //100 is max


//this fixes issue "The user is over quota" on delete and upload
var iSleep = 1000;


var sAtomHeader = "<atom:entry xmlns:atom='http://www.w3.org/2005/Atom'" +
    " xmlns:gd='http://schemas.google.com/g/2005'>" +
    "<atom:category scheme='http://schemas.google.com/g/2005#kind'" + 
    " term='http://schemas.google.com/contact/2008#contact' />";


var sBatchHeader = "<feed xmlns='http://www.w3.org/2005/Atom'" +
    " xmlns:gContact='http://schemas.google.com/contact/2008'" +
    " xmlns:gd='http://schemas.google.com/g/2005'" +
    " xmlns:batch='http://schemas.google.com/gdata/batch'>" +
    "<category scheme='http://schemas.google.com/g/2005#kind'" +
    " term='http://schemas.google.com/g/2008#contact' />"
        
        
var sNSurl = "http://schemas.google.com/g/2005";


//------------------------------------------------------------
function getEncodedSheetValue(aRowValues, iRow, aCols, aColIndex, bEncode){
  
  if (aCols[0][aColIndex]){
    
    //acols stores column letters where data is to be found
    var sColLetters = aCols[0][aColIndex].toLowerCase();
    //but the data is an array for speed, so need to turn the letters into a column number
    //a=97 z=122
    var iColIndex = sColLetters.charCodeAt(0) - 96;
    if (sColLetters.length == 2) iColIndex = (iColIndex * 26) + (sColLetters.charCodeAt(1) - 96);
    
    //now get the value - the array is zero indexed
    var sReturn = aRowValues[iRow - 1][iColIndex - 1];//oSheet.getRange(aCols[0][aColIndex] + iRow).getValue();
    
    //trim
    //sReturn = sReturn.replace(/^\s\s*/, '').replace(/\s\s*$/, '');
    
    if (bEncode) sReturn = htmlEncode(sReturn);
    
    //Browser.msgBox(sReturn);
    //return false;
    return sReturn;
    
  }else{
    return 0;
  }
  
}

//------------------------------------------------------------
function doUploadContacts(oSheet, aColNums, sSourceCode) {

  if (!checkSettings()) return false;
  
  //var oCell = oSheet.getRange('a1');

  var iRowsMax = oSheet.getLastRow();
  
  var iUploaded = 0;
  
  //limit for debugging?
  //iRowsMax = 10;
    
  var iBatchId = 0;
  var sBatchData = sBatchHeader;
  
  //for speed, read whole row into an array(
  //google is generous with memory but not with CPU cycles - 5 mins max execution time)
  var aRowValues = oSheet.getRange(1, 1, iRowsMax, oSheet.getLastColumn()).getValues();
  
  
  for (var iRow = 2; iRow <= iRowsMax; iRow++){
    
    //use Status filter to exclude unwanted rows
    var sFilterValueThis = getEncodedSheetValue(aRowValues, iRow, aColNums, 0, false);
    var sFilterAvoid = aColNums[0][1].toLowerCase();
    if (!sFilterValueThis || !sFilterAvoid || (sFilterValueThis.toLowerCase().indexOf(sFilterAvoid)) == -1){      
      
      var sFirstName = getEncodedSheetValue(aRowValues, iRow, aColNums, 2, true);
      //Logger.log(sFirstName);
      //return false;
      
      var sLastName = getEncodedSheetValue(aRowValues, iRow, aColNums, 3, true);
      var sCompany = getEncodedSheetValue(aRowValues, iRow, aColNums, 4, true);
      var sJobTitle = getEncodedSheetValue(aRowValues, iRow, aColNums, 5, true);
      var sEmailAddress = getEncodedSheetValue(aRowValues, iRow, aColNums, 6, false);
      var sPhoneMobile = getEncodedSheetValue(aRowValues, iRow, aColNums, 7, false);
      var sPhoneWork = getEncodedSheetValue(aRowValues, iRow, aColNums, 8, false);
      var sPhoneHome = getEncodedSheetValue(aRowValues, iRow, aColNums, 9, false);
      var sNotes = getEncodedSheetValue(aRowValues, iRow, aColNums, 10, true);
      
      
      var sPostData = bUseBatch ? "" : sAtomHeader;
      
      sPostData += "<gd:name>";
      sPostData += sFirstName ? " <gd:givenName>" + sFirstName + "</gd:givenName>" : "";
      sPostData += sLastName ? " <gd:familyName>" + sLastName + "</gd:familyName>" : ""; 
      //" <gd:fullName>" + sFirstName + " " + sLastName + "</gd:fullName>" +
      sPostData += "</gd:name>";
      
      sPostData += sEmailAddress ? "<gd:email rel='" + sNSurl+ "#work' primary='true' address='" + sEmailAddress + "' />" : "";
      
      sPostData += sPhoneWork ? "<gd:phoneNumber rel='" + sNSurl+ "#work'>" + sPhoneWork + "</gd:phoneNumber>" : "";
      sPostData += sPhoneMobile ? "<gd:phoneNumber rel='" + sNSurl+ "#mobile'>" + sPhoneMobile + "</gd:phoneNumber>" : "";  
      sPostData += sPhoneHome ? "<gd:phoneNumber rel='" + sNSurl+ "#home'>" + sPhoneHome + "</gd:phoneNumber>" : "";  
      //primary='true'
      
      
      if ((sJobTitle + sCompany) != ""){
        sPostData += "<gd:organization rel='http://schemas.google.com/g/2005#work'>";
        
        sPostData += sCompany ? "<gd:orgName>" + sCompany + "</gd:orgName>" : "";
        sPostData += sJobTitle ? "<gd:orgTitle>" + sJobTitle + "</gd:orgTitle>" : "";
        
        
        sPostData += "</gd:organization>";
      }
      
      var sAddressStreet = getEncodedSheetValue(aRowValues, iRow, aColNums, 11, true);
      if (sAddressStreet.length > 1){
        var sAddressStreet2 = getEncodedSheetValue(aRowValues, iRow, aColNums, 12, true);
        var sAddressCity = getEncodedSheetValue(aRowValues, iRow, aColNums, 13, true);
        var sAddressState = getEncodedSheetValue(aRowValues, iRow, aColNums, 14, true);
        var sAddressZip = getEncodedSheetValue(aRowValues, iRow, aColNums, 15, true);
        var sAddressCountry = getEncodedSheetValue(aRowValues, iRow, aColNums, 16, true);
        
        
        sPostData += "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'>";
        //#work' primary='true'>
        
        sPostData += "<gd:street>" + sAddressStreet;
        if (sAddressStreet2) sPostData += "\r" + sAddressStreet2;
        //space below deliberately added as quick fix for strange problem where street is empty even though
        //Grove Cottage Low Road
        //line above is adjusted, so try now without space
        sPostData += "</gd:street>";
        
        sPostData += sAddressCity ? "<gd:city>" + sAddressCity + "</gd:city>" : "";
        sPostData += sAddressState ? "<gd:region>" + sAddressState + "</gd:region>" : "";
        sPostData += sAddressZip ? "<gd:postcode>" + sAddressZip + "</gd:postcode>" : "";
        sPostData += sAddressCountry ? "<gd:country>" + sAddressCountry + "</gd:country>" : "";
        
        sPostData += "</gd:structuredPostalAddress>";
      }
      
      
      sPostData += "<gd:extendedProperty name='SourceCode' value='" + sSourceCode + "'/>";
      //sPostData += "<gd:extendedProperty name='Uploaded' value='" + new Date() + "'/>";
      
      
      if (bUseBatch){
        sPostData += sNotes ? "<content type='text'>" + sNotes + "</content>" : "";
      }else{
        sPostData += sNotes ? "<atom:content type='text'>" + sNotes + "</atom:content>" : "";
      }
      
      
      if (bUseBatch){
        iBatchId++;
        sBatchData += "<entry><batch:id>" + iBatchId + "</batch:id><batch:operation type='insert' />" + sPostData + "</entry>";
        
        
        if ((iBatchId >= iBatchSize) || ((iRow == iRowsMax))){// && (iBatchId > 1))){
          sBatchData += "</feed>";
          //Browser.msgBox(sBatchData);
          //Logger.log(sBatchData);
          //return false;
          
          //this fixes issue "The user is over quota"
          Utilities.sleep(iSleep);
          
          var sXML = GoogleHTTPrequest(sContactsApiUrl + "/full/batch", "cp", "POST", sBatchData, false);
          //todo - check for errors
          //<feed><entry>batch:interrupted error="0" parsed="0" reason="The prefix &quot;atom&quot; for element &quot;atom:content&quot; 
          
          //count these positive results
          //<batch:status code="201" reason="Created"
          
          //Logger.log(sXML.toXmlString());
          
          iBatchId = 0;
          sBatchData = sBatchHeader;
        }
        
      }else{
        
        sPostData += "</atom:entry>";
        
        //Browser.msgBox(sPostData);
        //return false;
        
        var sXML = GoogleHTTPrequest(sContactsApiUrl + "/full", "cp", "POST", sPostData, false);
      } 
      
      iUploaded ++;
      
    }
  }
  
  //send back the total - very lazily
  return iUploaded;
}


//------------------------------------------------------------
function doDeleteAll(){
  
  if (!checkSettings()) return false;
  
  
  var aEntryIds = new Array(0);
  var aEtags = new Array(0);
  
  var sNextUrl = sContactsApiUrl + "/base?max-results=200";
  
  //loop while there's another page of entries
  while (sNextUrl){
    
    //get a list of entries
    var parsedXml = GoogleHTTPrequest(sNextUrl, "cp", "GET", "", false);
    //Logger.log(parsedXml.toXmlString());
    //return false;
    
    //get the URL of the next page
    sNextUrl = getChildNodeWithSpecificAttribute(parsedXml, "link", "rel", "next", "href");

    //Logger.log("sNextUrl:" + sNextUrl + " aEntryIds.length:" + aEntryIds.length);

    //get the entries  
    var oElements = parsedXml.getElement().getElements("entry");  
    
  
    //loop through all entries (rows)
    for (var iCount = 0; iCount < oElements.length; iCount++){
    
      var oElement = oElements[iCount];
      
      //<entry gd:etag="&quot;Q30zcTVSLit7I2A9WhRXEEgIRwU.&quot;">
      var sEtag = oElement.getAttribute(sNSurl, "etag").getValue();
      aEtags.push(sEtag);
      
      //get the unique id
      var sEntryID = oElement.getElement("id").getText();
      
      sEntryID = sEntryID.replace(/base/g, 'full');
      
      //<link href="http://www.google.com/m8/feeds/contacts/n5ltd.co.uk/full/6a624419882228ee" rel="edit" type="application/atom+xml"/>
      
      //append to the array
      aEntryIds.push(sEntryID);
      
    }
    
  }
  
  
  
  if (aEntryIds.length == 0){
    Browser.msgBox(sAppTitle, "No shared contacts to delete", Browser.Buttons.OK);
    
  }else if (Browser.msgBox(sAppTitle, "Are you sure you want to delete all " + aEntryIds.length + " Shared Contacts?", Browser.Buttons.OK_CANCEL) == "ok"){
  
    oWorkbook.toast("Deleting " + aEntryIds.length + " contacts....", sAppTitle, -1);
    

    var iBatchId = 0;
    var sBatchData = sBatchHeader;
    
    //now we've got all the id's, delete them all
    for (var iCount = 0; iCount < aEntryIds.length; iCount++){
      //delete each entry             		    
      //sXML = GoogleHTTPrequest(aEntryIds[iCount], "cp", "DELETE*", "", false);
      
      iBatchId++;
      sBatchData += "<entry gd:etag='" + aEtags[iCount] + "'>" +
        "<batch:id>" + iBatchId + "</batch:id>" + 
        "<batch:operation type='delete' />" + 
        "<id>" + aEntryIds[iCount] + "</id>" + 
        "</entry>";
        
      
      if ((iBatchId >= iBatchSize) || (iCount == (aEntryIds.length - 1))){
        sBatchData += "</feed>";
          
        //Browser.msgBox("Deleting batch...." + iCount);
        //oWorkbook.toast("Deleting batch...." + iCount, sAppTitle, -1);
        //this fixes issue "The user is over quota"
        Utilities.sleep(iSleep);
        
        var sXML = GoogleHTTPrequest(sContactsApiUrl + "/full/batch", "cp", "POST", sBatchData, false);
        //Logger.log(sXML.toXmlString());
        //todo - check for errors
          
        iBatchId = 0;
        sBatchData = sBatchHeader;
      }
      
    }
    
    oWorkbook.toast("Finished" + getContactsCount() + " remaining", sAppTitle, 1);
    
  
  };
    	
}

//------------------------------------------------------------
//get the total number
function getContactsCount() {

  var parsedXml = GoogleHTTPrequest(sContactsApiUrl + "/base?max-results=1", "cp", "GET", "", false);
  var sTotal = parsedXml.getElement().getElement("http://a9.com/-/spec/opensearch/1.1/", "totalResults").getText();
  
  return parseInt(sTotal);

}
