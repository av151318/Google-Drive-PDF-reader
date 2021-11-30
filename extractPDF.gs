
//#####GLOBALS#####
const FOLDER_ID = "1Ix0yO0a8IIn38tvO0DbLmnCwxhWd6ORN"; //Folder ID of all PDFs
const SS = "1mOMR-YRNlIcOK8vXpXjx6PSKiYP1CSZRSx11CioqaYk";//The spreadsheet ID
const SHEET = "Extracted";//The sheet tab name
 
/*########################################################
* Main run file: extracts student IDs from PDFs and their
* section from the PDF name from multiple documents.
*
* Displays a list of students and sections in a Google Sheet.
*
*/
function extractStudentIDsAndSectionToSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  //Get all PDF files:
  const folder = DriveApp.getFolderById(FOLDER_ID);
  //const files = folder.getFiles();
  const files = folder.getFilesByType("application/pdf");
  
  let reportData = [];
  let reportData2 = [];
  let reportData3 = [];
  let reportData4 = [];
  //Iterate through each folderr
  while(files.hasNext()){
    let file = files.next();
    let fileID = file.getId();
    
    const doc = getTextFromPDF(fileID);
    const descr = extractDesc(doc.text);
    const pOIDs = extractPOIDs(doc.text);
    const costVals = extractCostVals(doc.text);
    const startDate = extractstartDate(doc.text);
    const endDate = extractendDate(doc.text);
    const orderNumber = extractONs(doc.text);
    //const invoiceNum = extractIN(doc.text);
    
    //Add ID to Section name..
    
    
    //const dSc = descr.map( DS => [DS]);
    const pID= pOIDs.map( POID => [POID]); 
    //const sD= startDate.map( SD => [SD]); 
    //const eD= endDate.map( ED => [ED]); 
    //const cV= costVals.map( CV =>[CV]);
    const oN= orderNumber.map( ON =>[ON]);
    //const iN= invoiceNum.map( IN =>[IN]);
    //const oracleData = startDate.map( SD => [doc.name,[],[],cV,SD,eD,[],pID]);
    //const oracleData = endDate.map( ED => [dSc,sN,[],cV,sD,ED,sN,pID,doc.name]);
   
    let oracleData = orderNumber.map( ON => [[],[],[],[],[],[],ON,pID,doc.name]);
    let oracleData2 = startDate.map( SD => [[],[],[],[],SD]);
    let oracleData3 = costVals.map( CV => [[],[],[],CV]);
    let oracleData4 = descr.map( DS => [DS,[doc.name],[pID]]);
  
    
    //Optional: Notify user of process. You can delete lines 33 to 38
    if(oN[0] === "No items found") {
      ss.toast("No items found in " + doc.name, "Warning",2);
    }else{
      ss.toast(doc.name + " extracted");
    };    
    
    //allIDsAndCRNs = allIDsAndCRNs.concat(oracleData);
    reportData = reportData.concat(oracleData);
    reportData2 = reportData2.concat(oracleData2);
    reportData3 = reportData3.concat(oracleData3);
    reportData4 = reportData4.concat(oracleData4);
    
  }
    importToSpreadsheet(reportData);
    importToSpreadsheet2(reportData2);
    importToSpreadsheet3(reportData3);
    importToSpreadsheet4(reportData4);
};
 
 
/*########################################################
* Extracts the text from a PDF and stores it in memory.
* Also extracts the file name.
*
* param {string} : fileID : file ID of the PDF that the text will be extracted from.
*
* returns {array} : Contains the file name (section) and PDF text.
*
*/
function getTextFromPDF(fileID) {
  var blob = DriveApp.getFileById(fileID).getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
    ocr: true,
    ocrLanguage: "en"
  };
  // Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
 
  // Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  var title = doc.getName();
  
  // Deleted the document once the text has been stored.
  Drive.Files.remove(doc.getId());
  
  return {
    name:title,
    text:text
  };
}
 
/*########################################################
* Use the text extracted from PDF and extracts student id based on value parameters.
* Also extracts the file name.
*
* param {string} : text : text of data from PDF.
*
* returns {array} : Of all student IDs found in text.
*
*/

function extractDesc(text){
  const regexp = /((Software Update ).*\n.*)/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};
function extractIN(text){
  const regexp = /(?<=INVOICE\s*NUMBER\s*).*/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

function extractPOIDs(text){
  const regexp = /000\d{7}|(?<=PO )000\d{7}/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

function extractONs(text){
  const regexp = /\S*(?=\s*Joe Carrasco\s*LOS RIOS)|\S*(?=\s*Accounts Payable\s*LOS RIOS)/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

function extractstartDate(text){
  const regexp = /((?<=\w*)\S*(?<!TO)(?= :\d*))/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

function extractendDate(text){
  const regexp =/((?<=\w \S*2021 :\s*)\w\S*)|((?<=\w \S*2020 :\s*)\w\S*)|((?<=\w \S*2021 :\n\s*)\w\S*)|((?<=\w \S*2020 :\n\s*)\w\S*)/gm
  //const regexp =/((?<=END DATE:)\s*\S*)|((?<=21 : )\S*)|((?<=20 : )\S*)|((?<=21 : )\S*\s*\S*)/mg
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

function extractCostVals(text){
  const regexp = /(?<=\sN\s*)(\d*(?=,)(?:.\d*).\d*)|((?<=\sN\s*)(?<!,\d*)\d*(?:\.\d*))/gm
  //const regexp = /(\d*(?=,)(?:.\d*).\d*)|((?<!,\d*)\d*(?:\.\d*))/mg

  
  
  //const regexp = /\d*(?=,)(?:.\d*).\d*/mg
  //const regexp2 = /(?<!,\d*)\d*(?:\.\d*)/mg

  try{
    let array = [...text.match(regexp)];
    return array;
  
  }  
  catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99.
    let array = ["Not items found"]
    return array;
  }
};

 
/*########################################################
* Takes the culminated list of IDs and sections and inserts them into
* a Google Sheet.
*
* param {array} : data : 2d array containing a list of ids and their associated sections.
*
*/
function importToSpreadsheet(data){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  
  const range = sheet.getRange(3,1,data.length,9);
  range.setValues(data);
  range.sort([9,1]);
}

function importToSpreadsheet2(data){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  
  const range = sheet.getRange(3,1,data.length,5);
  range.setValues(data);
  range.sort([5,1]);
}
function importToSpreadsheet3(data){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  
  const range = sheet.getRange(3,1,data.length,4);
  range.setValues(data);
  range.sort([4,1]);
}
function importToSpreadsheet4(data){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  
  const range = sheet.getRange(3,1,data.length,2);
  range.setValues(data);
  range.sort([2,1]);
}
