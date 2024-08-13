/*
  Made for Google Workspace Hackathon 2024 (GDSC APU)
  Project Name: Form-mation
  Team: SCAC
  Developer: Hong, Kar Kin
  Version: 2.5
  Last Modified: 13 August 2024 2:20PM GMT+8
*/

// Changes to these may require update to addRowConversion()
const SETUP_MAIN_COLUMN = [
  "Enabled", "Name", "Type",
  "TemplateUrl", "GDriveOutputUrl", "GFormUrl",
];

const PROJECT_FOLDER_NAME = "Form-mation";

const VAR_PREFIX = "{";
const VAR_SUFFIX = "}";

const patternString = `(?<=${VAR_PREFIX}).+?(?=${VAR_SUFFIX})`;
const uvRegex = new RegExp(patternString, 'g');

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
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Conversion')
      .addItem('Email','setEmailConversion')
      .addItem('Doc To PDF','setDocToPdfConversion')
      .addItem('Doc To Doc','setDocToDocConversion')
      .addItem('Slide To Slide','setSlideToSlideConversion')
      .addItem('Slide To PDF','setSlideToPdfConversion')
      )
    .addToUi();
}

function setEmailConversion() {
  addRowConversion(SUPPORTED_TYPE.EMAIL);
}

function setDocToPdfConversion() {
  addRowConversion(SUPPORTED_TYPE.DOC_TO_PDF);
}

function setDocToDocConversion() {
  addRowConversion(SUPPORTED_TYPE.DOC_TO_DOC);
}

function setSlideToSlideConversion() {
  addRowConversion(SUPPORTED_TYPE.SLIDE_TO_SLIDE);
}

function setSlideToPdfConversion() {
  addRowConversion(SUPPORTED_TYPE.SLIDE_TO_PDF);
}

function addRowConversion(type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var j = 1; j < data.length; j++) {
    if (data[j][1] == "") {
      var lastRow = j
      break;
    }
  }
  var fileName = 'Untitled Setup';

  if (
    type === SUPPORTED_TYPE.DOC_TO_DOC ||
    type === SUPPORTED_TYPE.DOC_TO_PDF ||
    type === SUPPORTED_TYPE.EMAIL
  ) {
    var document = DocumentApp.create(fileName),
      documentId = document.getId();

    DriveApp.getFileById(documentId).moveTo(getProjectFolder());

    var documentUrl = `https://docs.google.com/document/d/${documentId}/edit`;

    var inputs = [fileName, type, documentUrl, getProjectFolder().getUrl()];
    if (type === SUPPORTED_TYPE.EMAIL) {
      inputs.pop();
    }
    sheet.getRange(lastRow + 1, 2, 1, inputs.length).setValues([inputs]);
  } else if (
    type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
    type === SUPPORTED_TYPE.SLIDE_TO_PDF
  ) {
    var slide = SlidesApp.create(fileName),
      slideId = slide.getId();

    DriveApp.getFileById(slideId).moveTo(getProjectFolder());

    var slideUrl = `https://docs.google.com/presentation/d/${slideId}/edit`;

    var inputs = [fileName, type, slideUrl, getProjectFolder().getUrl()];
    sheet.getRange(lastRow + 1, 2, 1, inputs.length).setValues([inputs]);
  }
}

function onEdit(e) {
  const range = e.range;
  if (range.getColumn() === 1) {
    range.setNote('Last modified: ' + new Date() + "\n\nRemember to Reload!");
  }
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

function reload() {
  deleteAllTrigger();

  var controlPanelSetups = getControlPanelSetups();
  // console.log({controlPanelSetups});

  var row = 1;
  for (var i = 0; i < controlPanelSetups.length; i++) {
    row++;
    const cpSetupObj = controlPanelSetups[i];

    // remove hyperlink to publish URL of Name
    SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).setValue(cpSetupObj.Name);

    if (!cpSetupObj) continue;
    if (!cpSetupObj.Enabled) continue;

    // if row has no type, name, or template, will disable row
    if (
      disableRow(
        (!cpSetupObj.Name || cpSetupObj.Name == "") ||
        (!cpSetupObj.Type || cpSetupObj.Type == "") || 
        (!cpSetupObj.TemplateUrl || cpSetupObj.TemplateUrl == "")
      , row, "'Name', 'Type', and 'Template Link' is required!")
    ) continue;

    if (
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC ||
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
      cpSetupObj.Type === SUPPORTED_TYPE.EMAIL
    ) {
      if (
        disableRow(
          !cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/document/d/")
        , row, "'Template Link' contains an invalid Google Docs link.")
      ) {
        continue;
      }
      // check if row does not have GFormUrl and Variables filled
      if (
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "") &&
        cpSetupObj.Variables.length === 0 
      ) {
        // auto find and input variables into CP and generate Google Form to be inserted
        var uVariables = getDocsUniqueVariables(cpSetupObj);

        // set first variable as "EMAIL"
        if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
          uVariables = uVariables.filter(function(value, index, array) {
            return value != "EMAIL";
          });
          uVariables.unshift("EMAIL");
        }

        addVariablesToCP(row, uVariables);
        const newGFormLink = generateGoogleForms(cpSetupObj, uVariables);
        SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("GFormUrl")+1).setValue(newGFormLink);

        // update cpSetupObj for the remainder of this process
        cpSetupObj.GFormUrl = newGFormLink;
        cpSetupObj.Variables = uVariables;
      }
    } else if (
      cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
      cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
    ) {
      if (
        disableRow(
          !cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/presentation/d/")
        , row, "'Template Link' contains an invalid Google Slides link.")
      ) {
        continue;
      }
      // check if row does not have GFormUrl and Variables filled
      if (
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "") &&
        cpSetupObj.Variables.length === 0 
      ) {
        // auto find and input variables into CP and generate Google Form to be inserted
        var uVariables = getSlidesUniqueVariables(cpSetupObj);

        addVariablesToCP(row, uVariables);
        const newGFormLink = generateGoogleForms(cpSetupObj, uVariables);
        SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("GFormUrl")+1).setValue(newGFormLink);

        // update cpSetupObj for the remainder of this process
        cpSetupObj.GFormUrl = newGFormLink;
        cpSetupObj.Variables = uVariables;
      }
    }

    // if row does not have GFormUrl, will disable row
    if (
      disableRow(
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "")
      , row, "'Google Forms Link' is required!")
    ) continue;

    if (
      disableRow(
        !cpSetupObj.GFormUrl.startsWith("https://docs.google.com/forms/d/")
      , row, "'Google Forms Link' contains an invalid Google Forms link.")
    ) {
      continue;
    }

    if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
      // if row does not have "EMAIL" as first variable, will disable row
      if (
        disableRow(
          (!cpSetupObj.Variables[0] || cpSetupObj.Variables[0] != "EMAIL")
        , row, "For Type 'Email', the first variable in the Control Panel must be 'EMAIL'.")
      ) continue;
      createTrigger(cpSetupObj, "onEmailTrigger");
    } else {
      // if row does not have Output GDrive Link, will disable row
      if (
        disableRow(
          (!cpSetupObj.GDriveOutputUrl || cpSetupObj.GDriveOutputUrl == "")
        , row, "'Google Drive Link' is required!")
      ) continue;

      if (
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC
      ) {
        // (TOOD) Validate GDriveOutputUrl as accessible folder and link
        createTrigger(cpSetupObj, "onDocTrigger");
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
      ) {
        createTrigger(cpSetupObj, "onSlideTrigger");
      }
      //  else if (
      //   cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET
      // ) {
      //   createTrigger(cpSetupObj, "onSheetTrigger");
      // }
    }

    // set label name to hyperlink to published form URL
    updateNameWithFormPushlishedUrl(cpSetupObj, row);

    // acknowledge successful setups
    SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Enabled")+1).setNote('Setup Successful!\nAcknowledgement time: ' + new Date());
  }
}

function updateNameWithFormPushlishedUrl(cpSetupObj, row) {
  const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const formPublishedUrl = form.getPublishedUrl();
  const newValue = `=HYPERLINK("${formPublishedUrl}", "${cpSetupObj.Name}")`;
  SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).setValue(newValue);
}

function createTrigger(cpSetupObj, funcName) {
  const gf = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const triggerId = ScriptApp.newTrigger(funcName)
      .forForm(gf)
      .onFormSubmit()
      .create()
      .getUniqueId();
  Logger.log("Trigger created for '" + cpSetupObj.Name + "' to function '" + funcName + "' with triggerUID '" + triggerId + "'");
}

// // Current not in use as of v2.0
// function onSheetTrigger(e) {
//   // get form by triggerUid
//   const gFormId = getFileByTriggerId(e.triggerUid);
//   // Retrieve submitted form data
//   const formResponseData = getLatestFormResponse(gFormId);

//   var cpDataObj;
//   // Get control panel setup data based on formId
//   const cpSetups = getControlPanelSetups();
//   for (var i = 0; i < cpSetups.length; i++) {
//     const cpSetupObj = cpSetups[i];
//     if (cpSetupObj.GFormUrl.toString().includes(gFormId)) {
//       cpDataObj = cpSetupObj;
//     }
//   }

//   // get gsheets template
//   const templateSheet = SpreadsheetApp.openByUrl(cpDataObj.TemplateUrl);
//   const gSheetsTemplate = DriveApp.getFileById(templateSheet.getId());

//   var outputFileName = templateSheet.getName();

//   // get drive output location
//   // for ref https://drive.google.com/drive/u/0/folders/1yPBt0GbZweFD9wQoqRTA4oReuIU7-Jth
  // const outputFolderIdStartingIndex = cpDataObj.GDriveOutputUrl.toString().indexOf("folders/") + 8;
  // const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(outputFolderIdStartingIndex);
//   const destinationFolder = DriveApp.getFolderById(outputFolderId);

//   const copy = gSheetsTemplate.makeCopy(outputFileName, destinationFolder);
//   const sheets = SpreadsheetApp.openById(copy.getId());

//   sheets.getSheets().forEach(function(sheet) {
    
//   });

//   copy.setName(outputFileName);

//   // if (cpDataObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF) {
//   //   var blob = DriveApp.getFileById(sheets.getId()).getBlob();
//   //   destinationFolder.createFile(blob);
//   //   const sheetsFile = DriveApp.getFileById(sheets.getId());
//   //   sheetsFile.setTrashed(true);
//   // }
// }

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
  const outputFolderIdStartingIndex = cpDataObj.GDriveOutputUrl.toString().indexOf("folders/") + 8;
  const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(outputFolderIdStartingIndex);
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
  const outputFolderIdStartingIndex = cpDataObj.GDriveOutputUrl.toString().indexOf("folders/") + 8;
  const outputFolderId = cpDataObj.GDriveOutputUrl.toString().substring(outputFolderIdStartingIndex);
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
    if (variableName.startsWith("IMG") && replacementData.toString().length > 30) {
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
    if (variableName.startsWith("IMG") && replacementData.toString().length > 30) {
      const imgSrcUrl = `https://lh3.googleusercontent.com/d/${replacementData}`;
      const imgHtml = `<img src="${imgSrcUrl}" style="width: ${parseIMGVarName(variableName)}">`;
      html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, imgHtml);
      Logger.log(`full: ${VAR_PREFIX + variableName + VAR_SUFFIX}, imgHtm: ${imgHtml}`);
      continue;
    }
    html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
    subject = strReplaceAll(subject, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
  }

  const emailField = formResponseData[0];
  const emails = getEmails(emailField);

  MailApp.sendEmail({
    to: emails.toEmails,
    cc: emails.ccEmails,
    bcc: emails.bccEmail,
    subject: subject,
    htmlBody: html
  });
}

function getEmails() {
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
  // const varName = "IMG-300_G1";
  // const varName = "IMG-S_G2";

  var widthPX = 150; // default size for S
  const width = varName.match(/(?<=\-).+?(?=\_)/g)[0]; // getting first value between "-" and "_"
  
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

// finds all potential variables from google docs and returns an array of unique variables
function getDocsUniqueVariables(cpDataObj) {
  const doc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);
  const body = doc.getBody();

  // var variables = body.getText().match(/(?<=\{).+?(?=\})/g);
  // const docNameVariables = doc.getName().match(/(?<=\{).+?(?=\})/g);

  var variables = body.getText().match(uvRegex);
  const docNameVariables = doc.getName().match(uvRegex);

  if (!variables || variables.length === 0) {
    // If document name has variables, return those
    if (docNameVariables) {
      return docNameVariables.filter(function(value, index, arr) {
        return arr.indexOf(value) === index;
      });
    }
    return [];
  }

  if (docNameVariables) {
    const uniqueDNVs = docNameVariables.filter(function(value, index, arr) {
      return arr.indexOf(value) === index;
    });
    uniqueDNVs.forEach(function(value) {
      variables.unshift(value);
    });
  }

  // Remove all duplicates
  const uVariables = variables.filter(function(value, index, arr) {
    return arr.indexOf(value) === index;
  });

  return uVariables;
}

function getSlidesUniqueVariables(cpDataObj) {
  const slides = SlidesApp.openByUrl(cpDataObj.TemplateUrl);

  var variables = [];

  // get slide name variables
  const slideNameVariables = slides.getName().match(uvRegex);
  if (slideNameVariables) {
    const uniqueSNVs = slideNameVariables.filter(function(value, index, arr) {
      return arr.indexOf(value) === index;
    });
    uniqueSNVs.forEach(function(value) {
      variables.push(value);
    });
  }

  // get individual slides' variables
  slides.getSlides().forEach(function(slide) {
    var shapes = (slide.getShapes());
    shapes.forEach(function(shape) {
      const shapeVariables = shape.getText().asString().match(uvRegex);
      if (!shapeVariables) return;
      shapeVariables.forEach(function(value) {
        variables.push(value);
      });
    }); 
  });

  const uniqueVars = variables.filter(function(value, index, arr) {
    return arr.indexOf(value) === index;
  });
  return uniqueVars;
}

function addVariablesToCP(row, variables) {
  var column = SETUP_MAIN_COLUMN.length;
  variables.forEach(function(variable) {
    SpreadsheetApp.getActiveSheet().getRange(row, column + 1).setValue(variable);
    column++;
  });
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

function generateGoogleForms(cpDataObj, uVariables) {
  const newForm = FormApp.create(cpDataObj.Name);
  DriveApp.getFileById(newForm.getId()).moveTo(getProjectFolder());
  
  const form = FormApp.openById(newForm.getId());

  uVariables.forEach(function(variable) {
    const formItem = form.addParagraphTextItem();
    formItem.setRequired(true);
    if (variable.startsWith("IMG")) {
      formItem.setTitle(variable + " [CHANGE THIS TO 'File upload' TYPE -> 'Allow only specific file types' -> 'Image' ]");
      return;
    }
    formItem.setTitle(variable);
  });

  return form.getEditUrl();
}