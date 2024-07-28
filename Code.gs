/*
  Made for Google Workspace Hackathon 2024 (GDSC APU)
  Project Name: Form-mation
  Team: SCAC
  Developer: Hong
  Version: 1.8.1
  Last Modified: 28 July 2024 4:00AM GMT+8
*/

const SETUP_MAIN_COLUMN = [
  "Enabled", "Name", "Type",
  "TemplateUrl", "GFormUrl", "GDriveOutputUrl",
];
const VAR_PREFIX = "{";
const VAR_SUFFIX = "}";
const SUPPORTED_TYPE = {
  EMAIL: "Email",
  DOC_TO_PDF: "Doc-to-PDF",
  DOC_TO_DOC: "Doc-to-Doc",
  SLIDE_TO_SLIDE: "Slide-to-Slide",
  SLIDE_TO_PDF: "Slide-to-PDF",
  // SHEET_TO_SHEET: "Sheet-to-Sheet"
};

// To show the menu item to reload
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SCAC')
    .addItem('Reload', 'reload')
    .addToUi();
}

function disableRow(boolean, row) {
  if (boolean) {
    SpreadsheetApp.getActiveSheet().getRange(row, 1).setValue(false);
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

function reload() {
  deleteAllTrigger();

  var controlPanelSetups = getControlPanelSetups();
  // console.log({controlPanelSetups});

  var row = 1;
  for (var i = 0; i < controlPanelSetups.length; i++) {
    row++;
    const cpSetupObj = controlPanelSetups[i];
    if (!cpSetupObj) continue;
    if (!cpSetupObj.Enabled) continue;

    // if row does not have Template GDoc or GSheet, will disable row
    if (
      disableRow(
        (!cpSetupObj.TemplateUrl || cpSetupObj.TemplateUrl == "") ||
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "")
      , row)
    ) continue;

    if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
      // if row does not have "EMAIL" as first variable, will disable row
      if (
        disableRow(
          (!cpSetupObj.Variables[0] || cpSetupObj.Variables[0] != "EMAIL")
        , row)
      ) continue;
      setEmailTrigger(cpSetupObj);
    } else {
      // if row does not have Output GDrive Link, will disable row
      if (
        disableRow(
          (!cpSetupObj.GDriveOutputUrl || cpSetupObj.GDriveOutputUrl == "")
        , row)
      ) continue;

      if (
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC
      ) {
        // (TOOD) Validate GDriveOutputUrl as accessible folder and link
        setDocTrigger(cpSetupObj);
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
      ) {
        setSlideTrigger(cpSetupObj);
      }
      //  else if (
      //   cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET
      // ) {
      //   setSheetTrigger(cpSetupObj);
      // }
    }
  }
}

// For future code refactor use (as of v1.8)
function createTrigger(cpSetupObj, funcName) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger(funcName)
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId();
  Logger.log("Trigger created for '" + cpSetupObj.Name + "' to function '" + funcName + "' with triggerUID '" + triggerId + "'");
}

function setEmailTrigger(cpSetupObj) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger('onEmailTrigger')
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId();
}

function setDocTrigger(cpSetupObj) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger('onDocTrigger')
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId();
}

function setSlideTrigger(cpSetupObj) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger('onSlideTrigger')
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId();
}

// Current not in use as of v1.8
function setSheetTrigger(cpSetupObj) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger('onSheetTrigger')
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId(); 
}

// Current not in use as of v1.8
function onSheetTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);

  var cpDataObj;
  // Get control panel setup data based on formId
  const cpSetups = getControlPanelSetups();
  for (var i = 0; i < cpSetups.length; i++) {
    const cpSetupObj = cpSetups[i];
    if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
      cpDataObj = cpSetupObj;
    }
  }

  // get gsheets template
  const templateSheet = SpreadsheetApp.openByUrl(cpDataObj.TemplateUrl);
  const gSheetsTemplate = DriveApp.getFileById(templateSheet.getId());

  var outputFileName = templateSheet.getName();

  // get drive output location
  // for ref https://drive.google.com/drive/u/0/folders/1yPBt0GbZweFD9wQoqRTA4oReuIU7-Jth
  const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(43);
  const destinationFolder = DriveApp.getFolderById(outputFolderId);

  const copy = gSheetsTemplate.makeCopy(outputFileName, destinationFolder);
  const sheets = SpreadsheetApp.openById(copy.getId());

  sheets.getSheets().forEach(function(sheet) {
    
  });

  copy.setName(outputFileName);

  // if (cpDataObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF) {
  //   var blob = DriveApp.getFileById(sheets.getId()).getBlob();
  //   destinationFolder.createFile(blob);
  //   const sheetsFile = DriveApp.getFileById(sheets.getId());
  //   sheetsFile.setTrashed(true);
  // }
}

function onSlideTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);

  var cpDataObj;
  // Get control panel setup data based on formId
  const cpSetups = getControlPanelSetups();
  for (var i = 0; i < cpSetups.length; i++) {
    const cpSetupObj = cpSetups[i];
    if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
      cpDataObj = cpSetupObj;
    }
  }

  // get gslides template
  const templateSlide = SlidesApp.openByUrl(cpDataObj.TemplateUrl);
  const gSlidesTemplate = DriveApp.getFileById(templateSlide.getId());

  var outputFileName = templateSlide.getName();

  // get drive output location
  // for ref https://drive.google.com/drive/u/0/folders/1yPBt0GbZweFD9wQoqRTA4oReuIU7-Jth
  const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(43);
  const destinationFolder = DriveApp.getFolderById(outputFolderId);

  const copy = gSlidesTemplate.makeCopy(outputFileName, destinationFolder);
  const slides = SlidesApp.openById(copy.getId());

  slides.getSlides().forEach(function(slide) {
    var shapes = (slide.getShapes());
    shapes.forEach(function(shape) {
      for (var j = 0; j < cpDataObj.Variables.length; j++) {
        const variableName = cpDataObj.Variables[j];
        const replacementData = formResponseData[j];
        Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
        shape.getText().replaceAllText(VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
        outputFileName = strReplaceAll(outputFileName, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
      }
    }); 
  });

  slides.setName(outputFileName);
  slides.saveAndClose();

  if (cpDataObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF) {
    var blob = DriveApp.getFileById(slides.getId()).getBlob();
    destinationFolder.createFile(blob);
    const slidesFile = DriveApp.getFileById(slides.getId());
    slidesFile.setTrashed(true);
  }
}

function onDocTrigger(e) {
  // e = {"authMode":"FULL","response":{},"source":{ "triggerUid":"-2540183388906956469"}

  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);

  var cpDataObj;
  // Get control panel setup data based on formId
  const cpSetups = getControlPanelSetups();
  for (var i = 0; i < cpSetups.length; i++) {
    const cpSetupObj = cpSetups[i];
    if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
      cpDataObj = cpSetupObj;
    }
  }

  // get gdoc template
  const templateDoc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);
  const gDocTemplate = DriveApp.getFileById(templateDoc.getId());

  var outputFileName = templateDoc.getName();

  // get drive output location
  // for ref https://drive.google.com/drive/u/0/folders/1yPBt0GbZweFD9wQoqRTA4oReuIU7-Jth
  const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(43);
  const destinationFolder = DriveApp.getFolderById(outputFolderId);

  const copy = gDocTemplate.makeCopy(outputFileName, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  //All of the content lives in the body, so we get that for editing
  const body = doc.getBody();
  
  //In these lines, we replace our replacement tokens with values from our spreadsheet row
  // loop thru all variables specified & replace specified variables with form value
  for (var j = 0; j < cpDataObj.Variables.length; j++) {
    const variableName = cpDataObj.Variables[j];
    const replacementData = formResponseData[j];
    Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
    if (variableName.includes("IMG") && replacementData.toString().length > 30) {
      var image = DriveApp.getFileById(replacementData).getBlob();
      replaceTextToImage(body, VAR_PREFIX + variableName + VAR_SUFFIX, image, parseIMGVarName(variableName));
      continue;
    }
    body.replaceText(VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
    outputFileName = strReplaceAll(outputFileName, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
  }
  
  doc.setName(outputFileName);
  doc.saveAndClose();

  if (cpDataObj.Type === SUPPORTED_TYPE.DOC_TO_PDF) {
    const pdf = convertToPdf_(doc, destinationFolder); // Convert the doc to a PDF file.
    const url = pdf.getUrl(); // Get the URL of the new PDF file.
    const docFile = DriveApp.getFileById(doc.getId()); // Get the temporary Google Docs file.
    docFile.setTrashed(true); // Trash the temporary Google Docs file.
  }
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

function onEmailTrigger(e) {
  // e = {"authMode":"FULL","response":{},"source":{},"triggerUid":"-2540183388906956469"}

  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);

  var cpDataObj;
  // Get control panel setup data based on formId
  const cpSetups = getControlPanelSetups();
  for (var i = 0; i < cpSetups.length; i++) {
    const cpSetupObj = cpSetups[i];
    if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
      cpDataObj = cpSetupObj;
    }
  }

  // Stop operation if first variable is not {EMAIL}
  if (cpDataObj.Variables[0] != "EMAIL") return;

  // get gdoc template
  const templateDoc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);

  // get gdoc as HTML
  var html = getDocHtml(templateDoc.getId());

  var subject = templateDoc.getName();

  // loop thru all variables specified & replace specified variables with form value
  for (var j = 0; j < cpDataObj.Variables.length; j++) {
    const variableName = cpDataObj.Variables[j];
    const replacementData = formResponseData[j];
    Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
    if (variableName.includes("IMG") && replacementData.toString().length > 30) {
      const imgSrcUrl = `https://lh3.googleusercontent.com/d/${replacementData}`;
      const imgHtml = `<img src="${imgSrcUrl}" style="width: ${parseIMGVarName(variableName)}">`;
      html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, imgHtml);
      Logger.log(`full: ${VAR_PREFIX + variableName + VAR_SUFFIX}, imgHtm: ${imgHtml}`);
      continue;
    }
    html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
    subject = strReplaceAll(subject, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
  }

  MailApp.sendEmail(formResponseData[0], subject, html, {
    htmlBody: html
  });
}

function parseIMGVarName(varName) {
  // const varName = "IMG-300_G1";
  // const varName = "IMG-S_G2";

  var widthPX = 150; // default size for S
  const width = varName.match(/(?<=\-).+?(?=\_)/g)[0];
  
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
  
  // console.log(width, widthPX);
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

function getFormResponses(gFormId) {
  var result = []

  const form = FormApp.openById(gFormId);
  const formResponses = form.getResponses();

  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    var responseAnswers = [];
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      responseAnswers.push(itemResponse.getResponse());
    }
    result.push(responseAnswers);
  }
  return result;
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