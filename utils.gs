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

function getProjectFolder() {
  const drive = DriveApp.getFoldersByName(PROJECT_FOLDER_NAME);
  // create project folder if not exist
  if (!drive.hasNext()) {
    DriveApp.createFolder(PROJECT_FOLDER_NAME);
  } 
  const folder = DriveApp.getFoldersByName(PROJECT_FOLDER_NAME);

  return folder.next();
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
    'payload': JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(geminiEndpoint, options);
  const data = JSON.parse(response);
  const content = data["candidates"][0]["content"]["parts"][0]["text"];
  return content;
}