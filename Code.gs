/*
  Made for Google Workspace Hackathon 2024 (GDSC APU)
  Project Name: Form-mation
  Team: SCAC
  Developer: Hong, Kar Kin
  Version: 4.2
  Last Modified: 18 August 2024 2:40AM GMT+8
*/

const SETUP_MAIN_COLUMN = [
  "Enabled", "Name", "Type",
  "TemplateUrl", "GDriveOutputUrl", "GFormUrl",
];

const PROJECT_FOLDER_NAME = "Form-mation";
const LATEST_ACTIVITY_LOG_COUNT = 50;
const VAR_PREFIX = "{{";
const VAR_SUFFIX = "}}";

const PROPERTY_GEMINI_API_KEY = "GEMINI_API_KEY";
const PROPERTY_ACTIVITY_LOG = "ACTIVITY_LOG";

const SUPPORTED_TYPE = {
  EMAIL: "Email",
  DOC_TO_PDF: "Doc-to-PDF",
  DOC_TO_DOC: "Doc-to-Doc",
  SLIDE_TO_SLIDE: "Slide-to-Slide",
  SLIDE_TO_PDF: "Slide-to-PDF",
  SHEET_TO_SHEET: "Sheet-to-Sheet",
  SHEET_TO_PDF: "Sheet-to-PDF"
};

const DEFAULT_TYPE_TEMPLATE = {
  EMAIL: {
    name: "Email Sample",
    url: "https://docs.google.com/document/d/1LLRoaCZpEDCcByKSMTpsoav95Y5xhLmK5fLEJkda_d0/edit"
  },
  DOC_TO_DOC: {
    name: "Doc Sample",
    url: "https://docs.google.com/document/d/1JpqjS33Jl-538x0XhxdLIJCQhOsC0JvnrGc1e4jfxVM/edit"
  },
  DOC_TO_PDF: {
    name: "Doc Sample",
    url: "https://docs.google.com/document/d/1OGE4YggdJnJWDOtPRMetByNULJD_c4HPNX_pEnfX1YM/edit"
  },
  SLIDE_TO_SLIDE: {
    name: "Slide Sample",
    url: "https://docs.google.com/presentation/d/1GaWQQmruGXa-2MM06aHinr5fNhTqsaZ5QDdd7y6MjOo/edit"
  },
  SLIDE_TO_PDF: {
    name: "Slide Sample",
    url: "https://docs.google.com/presentation/d/1ndQrvvW2QWOYijM17yS6mmYdXGAl2uOyCHlT0Vma31k/edit"
  },
  SHEET_TO_SHEET: {
    name: "Sheet Sample",
    url: "https://docs.google.com/spreadsheets/d/1kJpe7FUw8jhjOt7aWm5_y9DvR1KWU21rQG2MSmJw78A/edit"    
  },
  SHEET_TO_PDF: {
    name: "Sheet Sample",
    url: "https://docs.google.com/spreadsheets/d/1aB3CuHHOZZbihWrKrk2ME781DoIxO55luDBNWsd5LYI/edit"    
  }
};

// Regular Expression used to find and retrieve variables
const patternString = `(?<=${VAR_PREFIX}).+?(?=${VAR_SUFFIX})`;
const uvRegex = new RegExp(patternString, 'g');

// To show the menu item to reload
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Form-mation')
    .addItem('✔ Validate', 'reload')
    .addItem('ℹ Information', 'openSidebar')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📄 Create a Blank Process')
      .addItem('Email','setEmailBlank')
      .addItem('Doc To Doc','setDocToDocBlank')
      .addItem('Doc To PDF','setDocToPdfBlank')
      .addItem('Slide To Slide','setSlideToSlideBlank')
      .addItem('Slide To PDF','setSlideToPdfBlank')
      .addItem('Sheet To Sheet', 'setSheetToSheetBlank')
      .addItem('Sheet To PDF', 'setSheetToPdfBlank')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('💡 Create a Sample Process')
      .addItem('Email','setEmailConversion')
      .addItem('Doc To Doc','setDocToDocConversion')
      .addItem('Doc To PDF','setDocToPdfConversion')
      .addItem('Slide To Slide','setSlideToSlideConversion')
      .addItem('Slide To PDF','setSlideToPdfConversion')
      .addItem('Sheet To Sheet', 'setSheetToSheetConversion')
      .addItem('Sheet To PDF', 'setSheetToPdfConversion')
    )
    .addItem('✨ Co-Create a Process with Gemini', 'openGeminiPrompt')
    .addSeparator()
    .addItem('🗨 Help Form-mation improve', 'helpFormmation')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙ Settings')
      .addItem('Gemini API Key', 'openGeminiKeyPrompt')
    ).addToUi();
}

function addRowBlank(type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var j = 1; j < data.length; j++) {
    if (data[j][1] == "") {
      var lastRow = j
      break;
    }
  }

  var processName = "";
  var templateUrl = "";

  if (
    type === SUPPORTED_TYPE.DOC_TO_DOC ||
    type === SUPPORTED_TYPE.DOC_TO_PDF ||
    type === SUPPORTED_TYPE.EMAIL
  ) {
    const newDoc = DocumentApp.create("Document Template");
    DriveApp.getFileById(newDoc.getId()).moveTo(getProjectFolder());
    if (type === SUPPORTED_TYPE.EMAIL) {
      newDoc.getBody().setMarginBottom(0);
      newDoc.getBody().setMarginLeft(0);
      newDoc.getBody().setMarginRight(0);
      newDoc.getBody().setMarginTop(0);
    }

    processName = "Document Process";
    templateUrl = `https://docs.google.com/document/d/${newDoc.getId()}/edit`;
  } else if (
    type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
    type === SUPPORTED_TYPE.SLIDE_TO_PDF
  ) {
    const file = SlidesApp.create("Slide Template");
    DriveApp.getFileById(file.getId()).moveTo(getProjectFolder());
    processName = "Slide Process";
    templateUrl = `https://docs.google.com/presentation/d/${file.getId()}/edit`;
  } else if (
    type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
    type === SUPPORTED_TYPE.SHEET_TO_PDF
  ) {
    const file = SpreadsheetApp.create("Sheet Template");
    DriveApp.getFileById(file.getId()).moveTo(getProjectFolder());
    processName = "Sheet Process";
    templateUrl = `https://docs.google.com/spreadsheets/d/${file.getId()}/edit`;
  }
  
  var inputs = [];

  inputs[SETUP_MAIN_COLUMN.indexOf("Name")] = processName;
  inputs[SETUP_MAIN_COLUMN.indexOf("Type")] = type;
  inputs[SETUP_MAIN_COLUMN.indexOf("TemplateUrl")] = templateUrl;
  if (type !== SUPPORTED_TYPE.EMAIL) {
    inputs[SETUP_MAIN_COLUMN.indexOf("GDriveOutputUrl")] = getProjectFolder().getUrl();
  }

  sheet.getRange(lastRow + 1, 1, 1, inputs.length).setValues([inputs]);
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

  const { url, name } = DEFAULT_TYPE_TEMPLATE[getKeyByValue(SUPPORTED_TYPE, type)];

  var templateFile = DriveApp.getFileById(getIdFromUrl(url));
  var copy = templateFile.makeCopy(templateFile.getName(), getProjectFolder());
  var inputs = [];

  inputs[SETUP_MAIN_COLUMN.indexOf("Name")] = name;
  inputs[SETUP_MAIN_COLUMN.indexOf("Type")] = type;
  inputs[SETUP_MAIN_COLUMN.indexOf("TemplateUrl")] = copy.getUrl();
  if (type !== SUPPORTED_TYPE.EMAIL) {
    inputs[SETUP_MAIN_COLUMN.indexOf("GDriveOutputUrl")] = getProjectFolder().getUrl();
  }

  sheet.getRange(lastRow + 1, 1, 1, inputs.length).setValues([inputs]);
}

function helpFormmation() {
  var ui = HtmlService.createHtmlOutputFromFile('feedback-page')
    .setHeight(500)
    .setWidth(500);
  SpreadsheetApp.getUi().showModelessDialog(ui, "Help Form-mation improve");
}

function onEdit(e) {
  const range = e.range;
  range.clearNote();
}

function reload() {
  deleteAllTrigger();

  var errorMsgs = [];
  var errorMsg = "";

  var controlPanelSetups = getControlPanelSetups();

  var row = 1;
  for (var i = 0; i < controlPanelSetups.length; i++) {
    row++;
    var cpSetupObj = controlPanelSetups[i];

    // remove hyperlink to publish URL of Name
    SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).setValue(cpSetupObj.Name);

    if (!cpSetupObj) continue;
    if (!cpSetupObj.Enabled) continue;

    // if row has no type, name, or template, will disable row
    errorMsg = "'Name', 'Type', and 'Template Link' is required!";
    if (
      disableRow(
        (!cpSetupObj.Name || cpSetupObj.Name == "") ||
        (!cpSetupObj.Type || cpSetupObj.Type == "") || 
        (!cpSetupObj.TemplateUrl || cpSetupObj.TemplateUrl == "")
      , row, errorMsg)
    ) {
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      continue;
    };

    if (
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC ||
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
      cpSetupObj.Type === SUPPORTED_TYPE.EMAIL
    ) {
      errorMsg = "'Template Link' contains an invalid Google Docs link.";
      if (
        disableRow(
          !(
            cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/document/d/")
          )
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
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
      } else {
        // update control panel and form items of new variables found in template
        cpSetupObj = updateVariablesFromTemplate(cpSetupObj, row);
        // update form confirmation message with newest GDriveOutputUrl
        if (cpSetupObj.Type !== SUPPORTED_TYPE.EMAIL) {
          const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
          form.setConfirmationMessage("Thank you for using Form-mation!\n\nGoogle Drive Folder: " + cpSetupObj.GDriveOutputUrl);
        }
      }
    } else if (
      cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
      cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
    ) {
      errorMsg = "'Template Link' contains an invalid Google Slides link.";
      if (
        disableRow(
          !cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/presentation/d/")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
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
    } else if (
      cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
      cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF
    ) {
      errorMsg = "'Template Link' contains an invalid Google Sheets link.";
      if (
        disableRow(
          !cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/spreadsheets/d/")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      // check if row does not have GFormUrl and Variables filled
      if (
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "") &&
        cpSetupObj.Variables.length === 0 
      ) {
        // auto find and input variables into CP and generate Google Form to be inserted
        var uVariables = getSheetsUniqueVariables(cpSetupObj);

        addVariablesToCP(row, uVariables);
        const newGFormLink = generateGoogleForms(cpSetupObj, uVariables);
        SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("GFormUrl")+1).setValue(newGFormLink);

        // update cpSetupObj for the remainder of this process
        cpSetupObj.GFormUrl = newGFormLink;
        cpSetupObj.Variables = uVariables;
      }
    }

    // if row does not have GFormUrl, will disable row
    errorMsg = "'Google Forms Link' is required!";
    if (
      disableRow(
        (!cpSetupObj.GFormUrl || cpSetupObj.GFormUrl == "")
      , row, errorMsg)
    ) {
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      continue;
    }

    errorMsg = "'Google Forms Link' contains an invalid Google Forms link.";
    if (
      disableRow(
        !cpSetupObj.GFormUrl.startsWith("https://docs.google.com/forms/d/")
      , row, errorMsg)
    ) {
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      continue;
    }

    if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
      // if row does not have "EMAIL" as first variable, will disable row
      errorMsg = "For Type 'Email', the first variable in the Control Panel must be 'EMAIL'.";
      if (
        disableRow(
          (!cpSetupObj.Variables[0] || cpSetupObj.Variables[0] != "EMAIL")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      createTrigger(cpSetupObj, "onEmailTrigger");
    } else {
      // if row does not have Output GDrive Link, will disable row
      errorMsg = "'Google Drive Folder Link' is required!";
      if (
        disableRow(
          (!cpSetupObj.GDriveOutputUrl || cpSetupObj.GDriveOutputUrl == "")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }

      errorMsg = "'Google Drive Folder Link' contains an invalid Google Drive folder link.";
      if (
        disableRow(
          !cpSetupObj.GDriveOutputUrl.startsWith("https://drive.google.com/drive/folders/")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }

      if (
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC
      ) {
        createTrigger(cpSetupObj, "onDocTrigger");
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
      ) {
        createTrigger(cpSetupObj, "onSlideTrigger");
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF
      ) {
        createTrigger(cpSetupObj, "onSheetTrigger");
      }
    }

    // set label name to hyperlink to published form URL
    updateNameWithFormPushlishedUrl(cpSetupObj, row);

    // acknowledge successful setups
    SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Enabled")+1).setNote('Process Validated.\n\nAcknowledged at: ' + getCurrentDateTimeString());
  }

  if (errorMsgs.length > 0) {
    SpreadsheetApp.flush();
    const ui = SpreadsheetApp.getUi();
    var alertMsg = "";
    errorMsgs.forEach(function ({ message, row, name }, index) {
      alertMsg += `${index + 1}. Process '${name}' on row ${row}
      Information: ${message}\n\n`;
    })
    ui.alert("Validation Issues Detected", `These processes are currently disabled due to an error.\nHover over the 'Enabled' cell of the error causing process for more information. Resolve the underlying issue, re-enable the process, and re-validate.\n\n` + alertMsg, ui.ButtonSet.OK);
  }

  newActivityLog(`Ran validation.`);
}

function updateNameWithFormPushlishedUrl(cpSetupObj, row) {
  const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
  const formPublishedUrl = form.getPublishedUrl();
  const newValue = `=HYPERLINK("${formPublishedUrl}", "${cpSetupObj.Name}")`;
  SpreadsheetApp.getActiveSheet().getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).setValue(newValue);
}

function updateVariablesFromTemplate(cpSetupObj, row) {
  var uVariables = [];
  // get all unique variables and check to CP
  if (
    cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC ||
    cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
    cpSetupObj.Type === SUPPORTED_TYPE.EMAIL
  ) {
    uVariables = getDocsUniqueVariables(cpSetupObj);
  } else if (
    cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
    cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
  ) {
    uVariables = getSlidesUniqueVariables(cpSetupObj);
  } else if (
    cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
    cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF
  ) {
    uVariables = getSheetsUniqueVariables(cpSetupObj);
  }

  var newVars = [];
  // get all CP variables and check with form
  uVariables.forEach(function (v) {
    const hasVar = cpSetupObj.Variables.includes(v);
    if (!hasVar) {
      newVars.push(v);
      cpSetupObj.Variables.push(v);
    }
  });
  if (newVars.length > 0) {
    addVariablesToCP(row, cpSetupObj.Variables);
    const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
    
    // update form with new items
    newVars.forEach(function (variable) {
      const formItem = form.addParagraphTextItem();
      formItem.setRequired(true);
      if (
        variable.startsWith("IMG") &&
        variable.includes("_")
      ) {
        const varDisplayName = variable.substring(variable.indexOf('_') + 1);
        formItem.setTitle(varDisplayName);
        formItem.setHelpText("[CHANGE THIS TO 'File upload' TYPE -> 'Allow only specific file types' -> 'Image']")
        return;
      }
      formItem.setTitle(variable);
      console.log({formInd: formItem.getIndex()});
    });

    // for if existing form has attachment items, move new item to before attachement items
    if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
      const formItems = form.getItems();
      form.moveItem(formItems.length - 1, cpSetupObj.Variables.length - 1);
    }
    newActivityLog(`Found new placeholder and updated for process '${cpSetupObj.Name}'.`);
  }

  return cpSetupObj;
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

function onSheetTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);
  // Get control panel setup data based on formId
  var cpDataObj = getCPSetupFromFormId(gFormId);

  // get gsheets template
  const templateSheet = SpreadsheetApp.openByUrl(cpDataObj.TemplateUrl);
  const gSheetsTemplate = DriveApp.getFileById(templateSheet.getId());

  var outputFileName = templateSheet.getName();

  // get drive output location
  const destinationFolder = DriveApp.getFolderById(getIdFromUrl(cpDataObj.GDriveOutputUrl));

  const copy = gSheetsTemplate.makeCopy("[FORM-MATION | PROCESSING] " + cpDataObj.Name, destinationFolder);
  const sheets = SpreadsheetApp.openById(copy.getId());

  sheets.getSheets().forEach(function(sheet) {
    for (var j = 0; j < cpDataObj.Variables.length; j++) {
      const variableName = cpDataObj.Variables[j];
      const replacementData = formResponseData[j];
      Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
      sheet.createTextFinder(VAR_PREFIX + variableName + VAR_SUFFIX).replaceAllWith(replacementData);
      outputFileName = strReplaceAll(outputFileName, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
    }
  });

  copy.setName(outputFileName);   

  if (cpDataObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF) {
    copy.setName(copy.getName() + ".pdf");
    var blob = DriveApp.getFileById(sheets.getId()).getBlob();
    destinationFolder.createFile(blob);
    const sheetsFile = DriveApp.getFileById(sheets.getId());
    sheetsFile.setTrashed(true);
  }
  newActivityLog(`Process '${cpDataObj.Name}' ran successfully!`);
}

function onSlideTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);
  // Get control panel setup data based on formId
  var cpDataObj = getCPSetupFromFormId(gFormId);

  // get gslides template
  const templateSlide = SlidesApp.openByUrl(cpDataObj.TemplateUrl);
  const gSlidesTemplate = DriveApp.getFileById(templateSlide.getId());

  var outputFileName = templateSlide.getName();

  // get drive output location
  const destinationFolder = DriveApp.getFolderById(getIdFromUrl(cpDataObj.GDriveOutputUrl));

  const copy = gSlidesTemplate.makeCopy("[FORM-MATION | PROCESSING] " + cpDataObj.Name, destinationFolder);
  const slides = SlidesApp.openById(copy.getId());

  try {
    slides.getSlides().forEach(function(slide) {
      var shapes = (slide.getShapes());

      shapes.forEach(function(shape) {
        for (var j = 0; j < cpDataObj.Variables.length; j++) {
          const variableName = cpDataObj.Variables[j];
          const replacementData = formResponseData[j];

          if (
            variableName.startsWith("IMG") && 
            shape.getText().asString().toString().startsWith(VAR_PREFIX + variableName) &&
            replacementData.toString().length > 30 && 
            fileExist(replacementData)
          ) {
            Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
            var image = DriveApp.getFileById(replacementData).getBlob();
            shape.replaceWithImage(image);
            continue;
          }

          shape.getText().replaceAllText(VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
          outputFileName = strReplaceAll(outputFileName, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
        }
      }); 
    });
  } catch (err) {
    console.log(`Slides replace with image err: ${err}`);
  }

  slides.setName(outputFileName);  
  slides.saveAndClose();

  if (cpDataObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF) {
    slides.setName(slides.getName() + ".pdf");
    var blob = DriveApp.getFileById(slides.getId()).getBlob();
    destinationFolder.createFile(blob);
    const slidesFile = DriveApp.getFileById(slides.getId());
    slidesFile.setTrashed(true);
  }
  newActivityLog(`Process '${cpDataObj.Name}' ran successfully!`);
}

function onDocTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);
  // Get control panel setup data based on formId
  var cpDataObj = getCPSetupFromFormId(gFormId);

  // get gdoc template
  const templateDoc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);
  const gDocTemplate = DriveApp.getFileById(templateDoc.getId());

  var outputFileName = templateDoc.getName();

  // get drive output location
  const destinationFolder = DriveApp.getFolderById(getIdFromUrl(cpDataObj.GDriveOutputUrl));

  const copy = gDocTemplate.makeCopy("[FORM-MATION | PROCESSING] " + cpDataObj.Name, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  //All of the content lives in the body, so we get that for editing
  const body = doc.getBody();
  
  // loop thru all variables specified & replace specified variables with form value
  for (var j = 0; j < cpDataObj.Variables.length; j++) {
    const variableName = cpDataObj.Variables[j];
    const replacementData = formResponseData[j];
    Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
    if (
      variableName.startsWith("IMG") && 
      replacementData.toString().length > 30 && 
      fileExist(replacementData)
    ) {
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
    doc.setName(doc.getName() + ".pdf");
    const pdf = convertToPdf_(doc, destinationFolder); // Convert the doc to a PDF file.
    const docFile = DriveApp.getFileById(doc.getId()); // Get the temporary Google Docs file.
    docFile.setTrashed(true); // Trash the temporary Google Docs file.
  }
  newActivityLog(`Process '${cpDataObj.Name}' ran successfully!`);
}

function onEmailTrigger(e) {
  // get form by triggerUid
  const gFormId = getFileByTriggerId(e.triggerUid);
  // Retrieve submitted form data
  const formResponseData = getLatestFormResponse(gFormId);
  // Get control panel setup data based on formId
  var cpDataObj = getCPSetupFromFormId(gFormId);

  // Stop operation if first variable is not {EMAIL}
  if (cpDataObj.Variables[0] != "EMAIL") return;

  // get gdoc template
  const templateDoc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);

  // get gdoc as HTML
  var html = getDocHtml(templateDoc.getId());

  var subject = templateDoc.getName();

  var inlineImages = {};

  // loop thru all variables specified & replace specified variables with form value
  for (var j = 0; j < cpDataObj.Variables.length; j++) {
    const variableName = cpDataObj.Variables[j];
    const replacementData = formResponseData[j];
    Logger.log("VarName: " + variableName + ", ReplaceData:" + replacementData);
    if (
      variableName.startsWith("IMG") && 
      replacementData.toString().length > 30 && 
      fileExist(replacementData)
    ) {
      const varDisplayName = variableName.substring(variableName.indexOf('_') + 1);
      const imgHtml = `<img src='cid:${variableName}' style='width:${parseIMGVarName(variableName)}px;'>`;
      html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, imgHtml);
      const imageBlob = DriveApp.getFileById(replacementData).getBlob().setName(varDisplayName);
      inlineImages[variableName] = imageBlob;

      Logger.log(`find: ${VAR_PREFIX + variableName + VAR_SUFFIX}, imgHtm: ${imgHtml}`);
      continue;
    }
    html = strReplaceAll(html, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
    subject = strReplaceAll(subject, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
  }

  var attachmentFiles = [];

  // if form consist more than required, Email attachments exist
  if (formResponseData.length > cpDataObj.Variables.length) {
    for (var i = cpDataObj.Variables.length; i < formResponseData.length; i++) {
      const attachmentFileId = formResponseData[i];
      if (!fileExist(attachmentFileId)) continue;
      const file = DriveApp.getFileById(attachmentFileId);
      attachmentFiles.push(file.getBlob());
    }
  }

  const emailField = formResponseData[0];
  const emails = getEmails(emailField);

  MailApp.sendEmail({
    to: emails.toEmails,
    cc: emails.ccEmails,
    bcc: emails.bccEmails,
    subject: subject,
    htmlBody: html,
    inlineImages: inlineImages,
    attachments: attachmentFiles
  });
  newActivityLog(`Process '${cpDataObj.Name}' ran successfully!`);
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

// finds all potential variables from google docs and returns an array of unique variables
function getDocsUniqueVariables(cpDataObj) {
  const doc = DocumentApp.openByUrl(cpDataObj.TemplateUrl);
  const body = doc.getBody();

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

function getSheetsUniqueVariables(cpDataObj) {
  const sheets = SpreadsheetApp.openByUrl(cpDataObj.TemplateUrl);

  var variables = [];

  // get spreadsheet name variables
  const sheetNameVariables = sheets.getName().match(uvRegex);
  if (sheetNameVariables) {
    const uniqueSNVs = sheetNameVariables.filter(function(value, index, arr) {
      return arr.indexOf(value) === index;
    });
    uniqueSNVs.forEach(function(value) {
      variables.push(value);
    });
  }

  // get individual sheets' variables
  sheets.getSheets().forEach(function(sheet) {
    sheet.getDataRange().getValues().forEach(function(row) {
      const rowVariables = row.toString().match(uvRegex);
      if (!rowVariables) return;
      rowVariables.forEach(function(value) {
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

function generateGoogleForms(cpDataObj, uVariables, existingFormUrl = undefined) {
  var form;
  if (existingFormUrl) {
    form = FormApp.openByUrl(existingFormUrl);
  } else {
    const newForm = FormApp.create(cpDataObj.Name);
    DriveApp.getFileById(newForm.getId()).moveTo(getProjectFolder());
    form = FormApp.openById(newForm.getId());
  }

  uVariables.forEach(function(variable) {
    const formItem = form.addParagraphTextItem();
    formItem.setRequired(true);
    if (
      variable.startsWith("IMG") &&
      variable.includes("_")
    ) {
      const varDisplayName = variable.substring(variable.indexOf('_') + 1);
      formItem.setTitle(varDisplayName);
      formItem.setHelpText("[CHANGE THIS TO 'File upload' TYPE -> 'Allow only specific file types' -> 'Image']")
      return;
    }
    formItem.setTitle(variable);
  });

  if (cpDataObj.Type === SUPPORTED_TYPE.EMAIL) {
    form.getItems()[0].setHelpText(`Example: "hong@email.ext cc: doe@mail.ext, jane@www.ext bcc: termi@rock.ext"`);
    form.setConfirmationMessage("Thank you for using Form-mation!");
  } else {
    form.setConfirmationMessage("Thank you for using Form-mation!\n\nGoogle Drive Folder: " + cpDataObj.GDriveOutputUrl);
  }

  newActivityLog(`Auto retrieved placeholders and generated Google Forms for process '${cpDataObj.Name}'.`);
  return form.getEditUrl();
}

function newActivityLog(activity) {
  var date = getCurrentDateTimeString();

  var propValue = PropertiesService.getScriptProperties().getProperty(PROPERTY_ACTIVITY_LOG);
  if (propValue) {
    propValue = JSON.parse(propValue);

    propValue.push({ date, activity });

    propValue.sort((a, b) => {
      return new Date(b.date) - new Date(a.date);
    });

    propValue = propValue.slice(0, LATEST_ACTIVITY_LOG_COUNT);
  } else propValue = [{ date, activity }];

  PropertiesService.getScriptProperties().setProperty(PROPERTY_ACTIVITY_LOG, JSON.stringify(propValue));
}

/*
  Sidebar
*/

function openSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Form-mation Information');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getCurrentDate() { return getCurrentDateTimeString(); }
function getEmailQuota() { return MailApp.getRemainingDailyQuota(); }
function getActivityHistory() {
  var value = PropertiesService.getScriptProperties().getProperty(PROPERTY_ACTIVITY_LOG);
  if (value) value = JSON.parse(value);
  else value = [];

  var output = value.map(({ date, activity }) => `<li><p class="activity-date">${date}</p><p class="activity-details">${activity}</p></li>`).join("");

  return output;
}

function clearActivityHistory() {
  PropertiesService.getScriptProperties().deleteProperty(PROPERTY_ACTIVITY_LOG);
}

/*
  Gemini Integration
*/

function openGeminiKeyPrompt() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const response = SpreadsheetApp.getUi().prompt(`Key (${scriptProperties.getProperty(PROPERTY_GEMINI_API_KEY)})\n\nTo remove your Gemini API Key, enter 'UNSET'.`)
  const responseText = response.getResponseText();
  const responseButton = response.getSelectedButton();
  if (responseButton == "CLOSE") return;
  if (!responseText) {
    SpreadsheetApp.getUi().alert("Gemini API Key can not be empty!\n\nGet your Gemini API Key here: https://aistudio.google.com/app/apikey\nIt looks something like this: 'AIz124CrPasyiPTVcZxsr-dinuertTw-P229bQc'");
    return;
  }
  if (responseText.length != 39) {
    if (responseText == 'UNSET') {
      scriptProperties.deleteProperty(PROPERTY_GEMINI_API_KEY);
      SpreadsheetApp.getUi().alert("Your Gemini API Key is removed!");
      return;
    }
    SpreadsheetApp.getUi().alert("Gemini API Key is invalid!\n\nGet your Gemini API Key here: https://aistudio.google.com/app/apikey\nIt looks something like this: 'AIz124CrPasyiPTVcZxsr-dinuertTw-P229bQc'");
    return;    
  }
  scriptProperties.setProperty(PROPERTY_GEMINI_API_KEY, responseText);
  return true;
}

const properties = PropertiesService.getScriptProperties().getProperties();
const geminiApiKey = properties[PROPERTY_GEMINI_API_KEY];
const geminiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro-latest:generateContent?key=${geminiApiKey}`;

function openGeminiPrompt(hasKey=false) {
  if (!geminiApiKey && !hasKey) {
    const response = openGeminiKeyPrompt();
    if (response) openGeminiPrompt(true);
    return;
  }

  var ui = HtmlService.createTemplateFromFile('gemini-input')
    .evaluate()
    .setHeight(265)
    .setWidth(450)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi().showModalDialog(ui, 'Create Process with Gemini');
}

function runGemini(selectedType, textDescription, rerun = 0) {
  var prePrompt = `Create a document template with a hard maximum of 5 placeholders within the template content, surrounded by double curly brackets ({{}}). Placeholders should be named descriptively (e.g., {{Date}}, {{Company Name}}). Adhere to the specified template type (e.g., letter, report, email). Use double asterisks (** **) for bold and double underscores (__ __) for underline formatting.

Generate a suitable file name without extensions, preferably with spaces, or a subject line for email templates. Create a process label name based on the provided description.

Output format:
<TEMPLATE_CONTENT>[template content]</TEMPLATE_CONTENT>
<FILE_NAME>[suggested file name]</FILE_NAME>
<EMAIL_SUBJECT>[if type is "Email", suggested email subject]</EMAIL_SUBJECT>
<PROCESS_NAME>[suggested template label]</PROCESS_NAME>
`;

  if (rerun > 0 && rerun < 5) {
    prePrompt = prePrompt + " It seems like the previous generation output did not satisfy the output format requirements, ensure that does not occur this time.";
  } else if (rerun >= 5) return false;

  const prompt = `${prePrompt} Type: "${selectedType}", Description: "${textDescription}"`;
  const output = callGemini(prompt);
  Logger.log({selectedType, textDescription, rerun})
  Logger.log({output});

  var processName, fileName, subjectTitle, templateContent;

  templateContent = output.match(/(?<=<TEMPLATE_CONTENT>).*(?=<\/TEMPLATE_CONTENT>)/s);
  processName = getGeminiOutputContent("PROCESS_NAME", output);
  fileName = getGeminiOutputContent("FILE_NAME", output);
  subjectTitle = getGeminiOutputContent("EMAIL_SUBJECT", output);

  if (!(templateContent && processName && fileName)) {
    return runGemini(selectedType, textDescription, rerun + 1);
  }

  templateContent = templateContent.toString().trim();
  processName = processName.toString().trim();
  fileName = fileName.toString().trim();

  if (subjectTitle && !subjectTitle.includes("Not Applicable")) {
    subjectTitle = subjectTitle.toString().trim();
    fileName = subjectTitle;
  }

  Logger.log({processName, fileName, templateContent});

  var fileUrl = "";

  if (
    selectedType === SUPPORTED_TYPE.DOC_TO_DOC ||
    selectedType === SUPPORTED_TYPE.DOC_TO_PDF ||
    selectedType === SUPPORTED_TYPE.EMAIL
  ) {
    const doc = DocumentApp.create(fileName);
    const docId = doc.getId();
    DriveApp.getFileById(docId).moveTo(getProjectFolder());
    fileUrl = `https://docs.google.com/document/d/${doc.getId()}/edit`;

    if (selectedType === SUPPORTED_TYPE.EMAIL) {
      doc.getBody().setMarginBottom(0);
      doc.getBody().setMarginLeft(0);
      doc.getBody().setMarginRight(0);
      doc.getBody().setMarginTop(0);
    }

    doc.getBody().setText(templateContent);

    // Document Formatting
    var underlineText = templateContent.match(/(?<=\_\_).+?(?=\_\_)/g);
    if (underlineText) {
      for (var i = 0; i < underlineText.length; i++) {
        const text = underlineText[i];
        // const textElement = doc.getBody().findText(text).getElement();
        // textElement.asText().setUnderline(true);
        doc.getBody().replaceText("__" + text + "__", text);
      }
    }
    var boldText = templateContent.match(/(?<=\*\*).+?(?=\*\*)/g);
    if (boldText) {
      for (var i = 0; i < boldText.length; i++) {
        const text = boldText[i];
        const textElement = doc.getBody().findText(text).getElement();
        textElement.asText().setBold(true);
        doc.getBody().replaceText("\\*\\*" + text + "\\*\\*", text);
      }
    }
  }

  geminiInsert(selectedType, processName, fileUrl);
  return fileUrl;
}

function geminiInsert(type, name, templateUrl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var j = 1; j < data.length; j++) {
    if (data[j][1] == "") {
      var lastRow = j
      break;
    }
  }

  var inputs = [];

  inputs[SETUP_MAIN_COLUMN.indexOf("Name")] = name;
  inputs[SETUP_MAIN_COLUMN.indexOf("Type")] = type;
  inputs[SETUP_MAIN_COLUMN.indexOf("TemplateUrl")] = templateUrl;
  if (type !== SUPPORTED_TYPE.EMAIL) {
    inputs[SETUP_MAIN_COLUMN.indexOf("GDriveOutputUrl")] = getProjectFolder().getUrl();
  }

  sheet.getRange(lastRow + 1, 1, 1, inputs.length).setValues([inputs]);
}