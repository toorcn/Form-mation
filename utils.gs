const DEFAULT_TYPE_TEMPLATE = {
  EMAIL: {
    name: "Meeting Reminder",
    url: "https://docs.google.com/document/d/1LLRoaCZpEDCcByKSMTpsoav95Y5xhLmK5fLEJkda_d0/edit"
  },
  DOC_TO_DOC: {
    name: "Meeting Notes",
    url: "https://docs.google.com/document/d/1JpqjS33Jl-538x0XhxdLIJCQhOsC0JvnrGc1e4jfxVM/edit"
  },
  DOC_TO_PDF: {
    name: "Internship Contract",
    url: "https://docs.google.com/document/d/1OGE4YggdJnJWDOtPRMetByNULJD_c4HPNX_pEnfX1YM/edit"
  },
  SLIDE_TO_SLIDE: {
    name: "Class Presentation",
    url: "https://docs.google.com/presentation/d/1ChZHf4NxauZrwSibuqM44z4-3bhamH2U9gbLhfJz6Tk/edit"
  },
  SLIDE_TO_PDF: {
    name: "Class Presentation",
    url: "https://docs.google.com/presentation/d/1FYluNikackMSbiIGkrohDwjL-CCqrUTgpx19r2Bo8fo/edit"
  },
  SHEET_TO_SHEET: {
    name: "Income Statement",
    url: "https://docs.google.com/spreadsheets/d/1U9XaBTPsMrn3-A7oQbF2GbtiQ1qzlGfovWnsBLZIgfI/edit"    
  },
  SHEET_TO_PDF: {
    name: "Purchase Order",
    url: "https://docs.google.com/spreadsheets/d/1LKSaWXkjwFKClJZwBccYOhkO22Jw76hV1WkZbYWzHyE/edit?gid=943090944#gid=943090944"    
  }
};

// Regular Expression used to find and retrieve variables
const patternString = `(?<=${VAR_PREFIX}).+?(?=${VAR_SUFFIX})`;
const uvRegex = new RegExp(patternString, 'g');

function setEmailBlank() { addRowBlank(SUPPORTED_TYPE.EMAIL); }
function setDocToPdfBlank() { addRowBlank(SUPPORTED_TYPE.DOC_TO_PDF); }
function setDocToDocBlank() { addRowBlank(SUPPORTED_TYPE.DOC_TO_DOC); }
function setSlideToSlideBlank() { addRowBlank(SUPPORTED_TYPE.SLIDE_TO_SLIDE); }
function setSlideToPdfBlank() { addRowBlank(SUPPORTED_TYPE.SLIDE_TO_PDF); }
function setSheetToSheetBlank() { addRowBlank(SUPPORTED_TYPE.SHEET_TO_SHEET); }
function setSheetToPdfBlank() { addRowBlank(SUPPORTED_TYPE.SHEET_TO_PDF); }
function setEmailConversion() { addRowConversion(SUPPORTED_TYPE.EMAIL); }
function setDocToPdfConversion() { addRowConversion(SUPPORTED_TYPE.DOC_TO_PDF); }
function setDocToDocConversion() { addRowConversion(SUPPORTED_TYPE.DOC_TO_DOC); }
function setSlideToSlideConversion() { addRowConversion(SUPPORTED_TYPE.SLIDE_TO_SLIDE); }
function setSlideToPdfConversion() { addRowConversion(SUPPORTED_TYPE.SLIDE_TO_PDF); }
function setSheetToSheetConversion() { addRowConversion(SUPPORTED_TYPE.SHEET_TO_SHEET); }
function setSheetToPdfConversion() { addRowConversion(SUPPORTED_TYPE.SHEET_TO_PDF); }

function getKeyByValue(object, value) {
  return Object.keys(object).find(key => object[key] === value);
}

function disableRow(boolean, row, noteMsg) {
  if (boolean) {
    const range = SpreadsheetApp.getActiveSheet().getRange(row, 1);
    range.setValue(false);
    range.setNote(noteMsg);
  }
  return boolean;
}

function getControlPanelSetups() {
  var result = [];

  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    const rowData = data[i];
    var varObjList = [];
    for (var j = SETUP_MAIN_COLUMN.length; j < rowData.length; j++) {
      if (rowData[j] == '') continue;
      varObjList.push(rowData[j]);
    }
    var setupObj = {
      Variables: varObjList
    };
    for (var j = 0; j < SETUP_MAIN_COLUMN.length; j++) {
      setupObj[SETUP_MAIN_COLUMN[j]] = rowData[j];
    }
    result.push(setupObj);
  }
  return result;
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function getCPSetupFromFormId(gFormId) {
  var cpDataObj;
  // Get control panel setup data based on formId
  const cpSetups = getControlPanelSetups();
  for (var i = 0; i < cpSetups.length; i++) {
    const cpSetupObj = cpSetups[i];
    if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
      cpDataObj = cpSetupObj;
    }
  }
  return cpDataObj;
}

function getCurrentDateTimeString() {
  return new Date().toLocaleString(undefined, {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    hour12: true
  });
}

function convertToPdf_(doc, folder) {
  const blob = doc.getAs('application/pdf');
  const file = folder.createFile(blob);
  return file;
}

// https://gist.github.com/tanaikech/f84831455dea5c394e48caaee0058b26
var replaceTextToImage = function(body, searchText, image, width) {
  var next = body.findText(searchText);
  while (next) { // slightly modified to replace all instances of "searchText"
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);
    if (width && typeof width == "number") {
      var w = img.getWidth();
      var h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    }
    next = body.findText(searchText, next);
  }
};

function fileExist(fileId) {
  var exist = false;
  try {
    DriveApp.getFileById(fileId);
    exist = true;
  } catch (e) {

  } finally {
    return exist;
  }
}

function getEmails(emailStr) {
  var toEmails = emailStr;
  var ccEmails = "";
  var bccEmails = "";

  if (emailStr.toString().includes(" bcc:")) {
    const str = emailStr.split(" bcc:")[1];
    bccEmails = str;
    // remove before cc
    if (str.toString().includes(" cc:")) {
      bccEmails = str.split(" cc:")[0];
    }
  }
  if (emailStr.toString().includes(" cc:")) {
    const str = emailStr.split(" cc:")[1];
    ccEmails = str;
    // remove before bcc
    if (str.toString().includes(" bcc:")) {
      ccEmails = str.split(" bcc:")[0];
    }
  }

  // toEmails 
  if (ccEmails || bccEmails) {
    if (emailStr.toString().includes(" cc:")) {
      const str = emailStr.split(" cc:")[0];
      toEmails = str;
      if (str.includes(" bcc:")) {
        toEmails = str.split(" bcc:")[0];
      }
    }
    if (emailStr.includes(" bcc:")) {
      const str = emailStr.split(" bcc:")[0];
      toEmails = str;
      if (str.includes(" cc:")) {
        toEmails = str.split(" cc:")[0];
      }
    }
  }
  if (toEmails.startsWith("cc:")) {
    toEmails = "";
    const str = emailStr.split("cc:")[1];
    ccEmails = str;
    // remove before bcc
    if (ccEmails.endsWith(" b")) {
      ccEmails = ccEmails.substring(0, ccEmails.length-2);
    }
  }
  if (toEmails.startsWith("bcc:")) {
    toEmails = "";
    const str = emailStr.split("bcc:")[1];
    bccEmails = str;
    // remove before cc
    if (str.toString().includes(" cc:")) {
      bccEmails = str.split(" cc:")[0];
    }
  }

  console.log({toEmails, ccEmails, bccEmails});
  return {toEmails, ccEmails, bccEmails};
}

function parseIMGVarName(varName) {
  var widthPX = 150; // default size for S
  var width = varName.match(/(?<=\-).+?(?=\_)/g)[0]; // getting first value between "-" and "_"
  width = width.toUpperCase();
  
  if (
    width === "S" ||
    width === "M" ||
    width === "L" 
  ) {
    if (width === "M") widthPX = 250;
    if (width === "L") widthPX = 450;
  } else if (isNumeric(width)) {
    widthPX = parseInt(width);
  }
  
  return widthPX;
}

function isNumeric(str) {
  if (typeof str != "string") return false;
  return !isNaN(str) && !isNaN(parseFloat(str));
}

function getDocHtml(gDocId) {
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + gDocId + "&exportFormat=html";
  var param = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(url, param).getContentText();
}

function strReplaceAll(subject, search, replacement) {
  function escapeRegExp(str) { return str.toString().replace(/[^A-Za-z0-9_]/g, '\\$&'); }
  search = search instanceof RegExp ? search : new RegExp(escapeRegExp(search), 'g');
  return subject.replace(search, replacement);
}

function getLatestFormResponse(gFormId) {
  const formResponses = getFormResponses(gFormId);
  return formResponses[formResponses.length-1];
}

function getFileByTriggerId(triggerId){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i =0; i<triggers.length; i++){
    if(triggers[i].getUniqueId() == triggerId){
      return triggers[i].getTriggerSourceId();
    }
  }
}

function deleteAllTrigger() {
  // Loop over all triggers.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let index = 0; index < allTriggers.length; index++) {
    ScriptApp.deleteTrigger(allTriggers[index]);
  }
}

function getFolder(folderName) {
  var projectFolder = getProjectFolder();

  const drive = DriveApp.getFoldersByName(folderName);
  // create project folder if not exist
  if (!drive.hasNext()) {
    DriveApp.createFolder(folderName);
  } 
  
  var folder, folderParents;

  do {
    folder = DriveApp.getFoldersByName(folderName).next();
    folderParents = folder.getParents().next();
    folder.moveTo(projectFolder);
  } while (folderParents.getName() != projectFolder.getName());

  return folder;
}

function getProjectFolder() {
  const drive = DriveApp.getFoldersByName(PROJECT_FOLDER_NAME);
  // create project folder if not exist
  if (!drive.hasNext()) {
    DriveApp.createFolder(PROJECT_FOLDER_NAME);
  } 
  const folder = DriveApp.getFoldersByName(PROJECT_FOLDER_NAME);

  return folder.next();
}

function extractNotionBlockId(notionUrl) {
  const regex = /#([a-f0-9]{32})/;
  const match = notionUrl.match(regex);
  return match ? match[1] : null;
}

function getNotionUrlType(notionUrl) {
  const id = extractNotionBlockId(notionUrl);
  const url = `https://api.notion.com/v1/blocks/${id}`;

  const headers = {
    'Authorization': `Bearer ${PropertiesService.getScriptProperties().getProperty(PROPERTY_NOTION_API_KEY)}`,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28'
  };

  const options = {
    method: 'get',
    contentType: 'application/json',
    headers: headers,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());
  console.log({responseData})
  if (responseData.status === 400 || responseData.status === 401) {
    return { message: responseData.message };
  }

  var is_toggleable;

  // get is toggleable
  if (responseData.heading_1) is_toggleable = responseData.heading_1.is_toggleable;
  if (responseData.heading_2) is_toggleable = responseData.heading_2.is_toggleable;
  if (responseData.heading_3) is_toggleable = responseData.heading_3.is_toggleable;

  var result = {
    type: responseData.type,
    is_toggleable
  };

  if (responseData?.table?.table_width > 0) {
    result['table_width'] = responseData.table.table_width;
  }

  return result;
}


function isSupportChildBlockNotion(notionUrl) {
  const res = getNotionUrlType(notionUrl);
  console.log({res})
  if (!res.type && res.message) {
    const ui = SpreadsheetApp.getUi();
    switch (res.message) {
      case 'path failed validation: path.block_id should be a valid uuid, instead was `"null"`.':
        return false;
      case "API token is invalid.":
        ui.alert(`Notion Authorization Error`, `Your Notion API Key may be invalid or you have not given your Notion Integration access to your Notion Page. Please correct the error for it to work.`, ui.ButtonSet.OK);
        return res;
    }
  }
  if (res.is_toggleable) return res;
  console.log(`isSupportChildBlockNotion type:${res.type} `);
  if (
    NOTION_SUPPORTED_TYPE.includes(res.type) &&
    res.is_toggleable === undefined
  ) {
    return res;
  }
  return false;
}

function openInputKeyPrompt({
  propertyName,
  keyName,
  getKeyUrl,
  keySample
}) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(`Setting your ${keyName}`, `Key (${scriptProperties.getProperty(propertyName)})\n\nTo remove your ${keyName}, enter 'UNSET'.`, ui.ButtonSet.OK_CANCEL);
  const responseText = response.getResponseText();
  const responseButton = response.getSelectedButton();
  if (responseButton != "OK") return;
  if (!responseText) {
    ui.alert(`Setting your ${keyName}`, `${keyName} can not be empty!\n\nGet your ${keyName} here: ${getKeyUrl}\nIt looks something like this: '${keySample}'`, ui.ButtonSet.OK);
    return;
  }
  if (responseText.length != keySample.length) {
    if (responseText == 'UNSET') {
      scriptProperties.deleteProperty(propertyName);
      ui.alert(`Setting ${keyName} success!`, `Your ${keyName} is removed!`, ui.ButtonSet.OK);
      return;
    }
    ui.alert(`Setting your ${keyName}`, `${keyName} is invalid!\n\nGet your ${keyName} here: ${getKeyUrl}\nIt looks something like this: '${keySample}'`, ui.ButtonSet.OK);
    return;    
  }
  scriptProperties.setProperty(propertyName, responseText);
  return responseButton;
}

function callGemini(prompt, temperature=0.5) {
  const payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          },
        ]
      }
    ], 
    "generationConfig":  {
      "temperature": temperature,
    },
  };

  const options = { 
    'method' : 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(geminiEndpoint, options);
  const data = JSON.parse(response);
  if (data?.error?.message == "API key expired. Please renew the API key.") {
    const ui = SpreadsheetApp.getUi();
    // Override and close Gemini Modal
    var output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
    ui.showModalDialog(output, 'Loading...');
    // Alert user to error and abort
    ui.alert(`Gemini Authorization Error`, `Your Gemini API Key may be invalid or have expired. Please update your API key before trying again.`, ui.ButtonSet.OK);
    PropertiesService.getScriptProperties().deleteProperty(PROPERTY_GEMINI_API_KEY);
    return false;
  }
  const content = data["candidates"][0]["content"]["parts"][0]["text"];
  return content;
}

function getGeminiOutputContent(Keyword, str) {
  const rg = new RegExp(`(?<=<${Keyword}>).+?(?=<\/${Keyword}>)`);
  return str.match(rg);
}

function createAndLinkSpreadsheet(formId) {
  // Create a new spreadsheet
  var spreadsheet = SpreadsheetApp.create('Form Responses');
  var spreadsheetId = spreadsheet.getId();
  var spreadsheetUrl = spreadsheet.getUrl();
  
  // Open the form using the form ID
  var form = FormApp.openById(formId);
  
  // Set the destination of form responses to the new spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  
  return spreadsheetUrl;
}

function getSheetData(gFormId) {
  const sheetUrl = getLinkedSpreadsheetUrl(gFormId);
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sheet = spreadsheet.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.forEach(array => array.splice(0,1));
  data = data.filter(subarray => !subarray.every(element => element === ''));
  console.log({data});
  // delete all
  sheet.getRange(2,1,sheet.getLastRow(), sheet.getLastColumn()).setValue("");
  return data;
}

function getLinkedSpreadsheetUrl(formId) {
  // Open the form using the form ID
  var form = FormApp.openById(formId);
  
  // Get the URL of the linked Google Sheets
  var destinationId = form.getDestinationId();
  if (destinationId) {
    var spreadsheet = SpreadsheetApp.openById(destinationId);
    var spreadsheetUrl = spreadsheet.getUrl();
    return spreadsheetUrl;
  } else {
    return 'No linked spreadsheet found.';
  }
}

function splitMergedCells(rangeA1Notation) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rangeA1Notation);
  var mergedRanges = range.getMergedRanges();
  
  mergedRanges.forEach(function(mergedRange) {
    mergedRange.breakApart();
  });
}

function mergeTwoColumns(rangeA1Notation) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rangeA1Notation);
  var numRows = range.getNumRows();
  
  for (var i = 1; i <= numRows; i++) {
    var cell1 = range.getCell(i, 1);
    var cell2 = range.getCell(i, 2);
    sheet.getRange(cell1.getRow(), cell1.getColumn(), 1, 2).mergeAcross();
  }
}

function splitMergedCellsByRow(row) {
  splitMergedCells(`C${row}:D${row}`);
}

function mergeTwoColumnsByRow(row) {
  mergeTwoColumns(`C${row}:D${row}`);
}

function clearAfterColumn(rowIndex, colIndex) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastColumn = sheet.getLastColumn();

  console.log({colIndex, lastColumn});
  
  if (colIndex < lastColumn) {
    var range = sheet.getRange(rowIndex, colIndex + 1, 1, lastColumn - colIndex);
    console.log({values: range.getValues()})
    range.clear();
  }
}