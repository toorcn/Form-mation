/*
  Made for Google Workspace Hackathon 2024 (GDSC APU)
  Project Name: Form-mation
  Team: SCAC
  Developer: Hong, Kar Kin
  Version: 4.6
  Last Modified: 21 August 2024 11:00AM GMT+8
*/

const SETUP_MAIN_COLUMN = [
  "Enabled", "Name", "Type",
  "TemplateUrl", "GDriveOutputUrl", "NotionUrl", "GFormUrl",
];

const PROJECT_FOLDER_NAME = "Form-mation";
const LATEST_ACTIVITY_LOG_COUNT = 50;
const VAR_PREFIX = "{{";
const VAR_SUFFIX = "}}";

const PROPERTY_GEMINI_API_KEY = "GEMINI_API_KEY";
const PROPERTY_NOTION_API_KEY = "NOTION_API_KEY";
const PROPERTY_ACTIVITY_LOG = "ACTIVITY_LOG";

// https://developers.notion.com/reference/block#block-types-that-support-child-blocks
const NOTION_SUPPORTED_TYPE = ['bulleted_list_item', 'callout', 'child_database', 'child_page', 'column', 'numbered_list_item', 'paragraph', 'quote', 'synced_block', 'template', 'to_do', 'toggle', 'table'];

const SUPPORTED_TYPE = {
  EMAIL: "Email",
  DOC_TO_DOC: "Doc-to-Doc",
  DOC_TO_PDF: "Doc-to-PDF",
  SLIDE_TO_SLIDE: "Slide-to-Slide",
  SLIDE_TO_PDF: "Slide-to-PDF",
  SHEET_TO_SHEET: "Sheet-to-Sheet",
  SHEET_TO_PDF: "Sheet-to-PDF"
};

// To show the menu item to reload
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Form-mation')
    .addItem('✔ Validate', 'reload')
    .addItem('ℹ Information (Help)', 'openSidebar')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📄 Create a Blank Process')
      .addItem('Email','setEmailBlank')
      .addItem('Doc-to-Doc','setDocToDocBlank')
      .addItem('Doc-to-PDF','setDocToPdfBlank')
      .addItem('Slide-to-Slide','setSlideToSlideBlank')
      .addItem('Slide-to-PDF','setSlideToPdfBlank')
      .addItem('Sheet-to-Sheet', 'setSheetToSheetBlank')
      .addItem('Sheet-to-PDF', 'setSheetToPdfBlank')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('💡 Create a Sample Process')
      .addItem('Email','setEmailConversion')
      .addItem('Doc-to-Doc','setDocToDocConversion')
      .addItem('Doc-to-PDF','setDocToPdfConversion')
      .addItem('Slide-to-Slide','setSlideToSlideConversion')
      .addItem('Slide-to-PDF','setSlideToPdfConversion')
      .addItem('Sheet-to-Sheet', 'setSheetToSheetConversion')
      .addItem('Sheet-to-PDF', 'setSheetToPdfConversion')
    )
    .addItem('✨ Co-Create a Process with Gemini', 'openGeminiPrompt')
    .addSeparator()
    .addItem('🗨 Help Form-mation improve', 'openFeedback')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙ Settings')
      .addItem('Gemini API Key', 'openGeminiKeyPrompt')
      .addItem('Notion API Key', 'openNotionKeyPrompt')
    ).addToUi();
}

// adds blank process & template
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
    var newDoc;
    if (type === SUPPORTED_TYPE.EMAIL) {
      newDoc = DocumentApp.create("Email Template");
      processName = "Email Process";
    } else {
      newDoc = DocumentApp.create("Document Template");
      processName = "Document Process";
    }
    DriveApp.getFileById(newDoc.getId()).moveTo(getProjectFolder());
    if (type === SUPPORTED_TYPE.EMAIL) {
      newDoc.getBody().setMarginBottom(0);
      newDoc.getBody().setMarginLeft(0);
      newDoc.getBody().setMarginRight(0);
      newDoc.getBody().setMarginTop(0);
    }

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

// adds sample process & template from making a copy of a predefined document
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

function openFeedback() {
  const title = 'Help Form-mation improve';
  var template = HtmlService.createTemplateFromFile('iframe-page');
  template.pageUrl = "https://forms.gle/zXCRRJ6qjaBcoLWG7";
  template.title = title;
  var htmlOutput = template.evaluate()
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

function onEdit(e) {
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();

  range.clearNote();

  if (column === 1 && range.getValue() === false) {
    removeProcessNameHyperlink(row);
    range.setNote('Heads up! To deactivate a process, validate again.');
  }
}

// remove hyperlink to publish URL of Name
function removeProcessNameHyperlink(row) {  
  const sheet = SpreadsheetApp.getActiveSheet();
  const processName = sheet.getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).getValue();
  sheet.getRange(row, SETUP_MAIN_COLUMN.indexOf("Name")+1).setValue(processName);
}

// Main function for validation
function reload() {
  var errorMsgs = [];
  var errorMsg = "";
  var processSuccessCount = 0;
  var processFailCount = 0;

  deleteAllTrigger();

  var controlPanelSetups, row;

  controlPanelSetups = getControlPanelSetups();
  row = 1;
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
      processFailCount++;
      continue;
    };

    // test if user has permission to links (template, folder)
    try {
      if (
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC ||
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
        cpSetupObj.Type === SUPPORTED_TYPE.EMAIL
      ) { DocumentApp.openByUrl(cpSetupObj.TemplateUrl); }
      if (
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
      ) { SlidesApp.openByUrl(cpSetupObj.TemplateUrl); }
      if (
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF
      ) { SpreadsheetApp.openByUrl(cpSetupObj.TemplateUrl); }
      if (cpSetupObj.Type != SUPPORTED_TYPE.EMAIL) {
        DriveApp.getFolderById(getIdFromUrl(cpSetupObj.GDriveOutputUrl));
      }
    } catch (error) {
      processFailCount++;
      errorMsg = `You do not have permission to use one of the links of this process, please check the file permission and try again.`;
      if (
        disableRow(error == 'Exception: Action not allowed'
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      if (
        disableRow(error == 'Exception: No item with the given ID could be found. Possibly because you have not edited this item or you do not have permission to access it.'
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      errorMsg = `Unexpected error when checking file access. Error: ${error}`;
      disableRow(true, row, errorMsg);
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      Logger.log(`Uncaught error when checking file access for process: ${cpSetupObj.Name}, error: ${error}`);
      continue;
    }

    if (
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC ||
      cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
      cpSetupObj.Type === SUPPORTED_TYPE.EMAIL
    ) {
      errorMsg = "'Template Link' contains an invalid Google Docs link.";
      if (
        disableRow(!cpSetupObj.TemplateUrl.startsWith("https://docs.google.com/document/d/"), row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        processFailCount++;
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
        const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
        if (cpSetupObj.Type !== SUPPORTED_TYPE.EMAIL) {
          form.setConfirmationMessage("Thank you for using Form-mation!\n\nGoogle Drive Folder: " + cpSetupObj.GDriveOutputUrl);
        } else {
          form.setConfirmationMessage("Thank you for using Form-mation!\n\nExpect the email to be sent within the next 3 minutes.");
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
        processFailCount++;
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
      } else {
        // update control panel and form items of new variables found in template
        cpSetupObj = updateVariablesFromTemplate(cpSetupObj, row);
        // update form confirmation message with newest GDriveOutputUrl
        const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
        form.setConfirmationMessage("Thank you for using Form-mation!\n\nGoogle Drive Folder: " + cpSetupObj.GDriveOutputUrl);
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
        processFailCount++;
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
      } else {
        // update control panel and form items of new variables found in template
        cpSetupObj = updateVariablesFromTemplate(cpSetupObj, row);
        // update form confirmation message with newest GDriveOutputUrl
        const form = FormApp.openByUrl(cpSetupObj.GFormUrl);
        form.setConfirmationMessage("Thank you for using Form-mation!\n\nGoogle Drive Folder: " + cpSetupObj.GDriveOutputUrl);
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
      processFailCount++;
      continue;
    }

    errorMsg = "'Google Forms Link' contains an invalid Google Forms link.";
    if (
      disableRow(
        !cpSetupObj.GFormUrl.startsWith("https://docs.google.com/forms/d/")
      , row, errorMsg)
    ) {
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      processFailCount++;
      continue;
    }

    // test if user has permission to links (forms)
    try {
      FormApp.openByUrl(cpSetupObj.GFormUrl).getEditors();
    } catch (error) {
      processFailCount++;
      errorMsg = `You do not have permission to use one of the links of this process, please check the file permission and try again.`;
      if (
        disableRow(error == 'Exception: Action not allowed'
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      if (
        disableRow(error == 'Exception: No item with the given ID could be found. Possibly because you have not edited this item or you do not have permission to access it.'
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        continue;
      }
      errorMsg = `Unexpected error when checking file access. Error: ${error}`;
      disableRow(true, row, errorMsg);
      errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
      Logger.log(`Uncaught error when checking file access for process: ${cpSetupObj.Name}, error: ${error}`);
      continue;
    }

    if (cpSetupObj.Type === SUPPORTED_TYPE.EMAIL) {
      // if row does not have "EMAIL" as first variable, will disable row
      errorMsg = "For Type 'Email', the first placeholder of the process must be 'EMAIL'.";
      if (
        disableRow(
          (!cpSetupObj.Variables[0] || cpSetupObj.Variables[0] != "EMAIL")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        processFailCount++;
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
        processFailCount++;
        continue;
      }

      errorMsg = "'Google Drive Folder Link' contains an invalid Google Drive folder link.";
      if (
        disableRow(
          !cpSetupObj.GDriveOutputUrl.startsWith("https://drive.google.com/drive/folders/")
        , row, errorMsg)
      ) {
        errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
        processFailCount++;
        continue;
      }

      if (cpSetupObj.NotionUrl) {
        // has key
        const notionApiKey = PropertiesService.getScriptProperties().getProperty(PROPERTY_NOTION_API_KEY);
        if (!notionApiKey) {
          openNotionKeyPrompt();
        }
        // is support
        errorMsg = `'(Optional) Notion Link' contains an invalid or unsupported Notion Block link. Supported Notion block types: ${NOTION_SUPPORTED_TYPE.join(", ")}`;
        if (
          disableRow(!isSupportChildBlockNotion(cpSetupObj.NotionUrl)
          , row, errorMsg)
        ) {
          errorMsgs.push({ message: errorMsg, row, name: cpSetupObj.Name });
          processFailCount++;
          continue;
        }
      }

      if (
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_PDF ||
        cpSetupObj.Type === SUPPORTED_TYPE.DOC_TO_DOC
      ) {
        createTrigger(cpSetupObj, "onDocTrigger");
        processSuccessCount++;
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
        cpSetupObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF
      ) {
        createTrigger(cpSetupObj, "onSlideTrigger");
        processSuccessCount++;
      } else if (
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_SHEET ||
        cpSetupObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF
      ) {
        createTrigger(cpSetupObj, "onSheetTrigger");
        processSuccessCount++;
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
    ui.alert("Validation Issues Detected", `These processes are currently disabled due to an error.\nHover over the 'Active' cell of the error causing process for more information. Resolve the underlying issue, re-enable the process, and re-validate.\n\n` + alertMsg, ui.ButtonSet.OK);
  }

  newActivityLog(`Validation Status: ${processSuccessCount} (Validated), ${processFailCount} (Invalidated)`);
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

/*
  Google Forms on submit to trigger background process
*/

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

  var outputFile = copy;   

  if (cpDataObj.Type === SUPPORTED_TYPE.SHEET_TO_PDF) {
    copy.setName(copy.getName() + ".pdf");
    var blob = DriveApp.getFileById(sheets.getId()).getBlob();
    outputFile = destinationFolder.createFile(blob);
    const sheetsFile = DriveApp.getFileById(sheets.getId());
    sheetsFile.setTrashed(true);
  }
  
  // if has notion url, perform insert
  const notionUrl = cpDataObj.NotionUrl;
  if (notionUrl) {
    appendBlockToNotion(notionUrl, outputFile, undefined, cpDataObj);
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

  slides.getSlides().forEach(function(slide) {
    var shapes = (slide.getShapes());

    var imageObj = {};

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
          imageObj[variableName] = DriveApp.getFileById(replacementData).getBlob();
          continue;
        }

        shape.getText().replaceAllText(VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
        outputFileName = strReplaceAll(outputFileName, VAR_PREFIX + variableName + VAR_SUFFIX, replacementData);
      }
    });

    var currentVar;
    try {
      shapes.forEach(function(s) {
        var text = s.getText().asString();
        currentVar = text;
        text = strReplaceAll(text, '\n', '');
        text = text.match(uvRegex);
        if (text) {
          text = text[0];
          currentVar = text;
          if (imageObj[text]) s.replaceWithImage(imageObj[text]);
        }
      });
    } catch (err) {
      console.log(`Slides replace '${currentVar}' with image err: ${err}`);
    }
  });

  slides.setName(outputFileName);  
  slides.saveAndClose();

  var outputFile = slides;

  if (cpDataObj.Type === SUPPORTED_TYPE.SLIDE_TO_PDF) {
    slides.setName(slides.getName() + ".pdf");
    var blob = DriveApp.getFileById(slides.getId()).getBlob();
    outputFile = destinationFolder.createFile(blob);
    const slidesFile = DriveApp.getFileById(slides.getId());
    slidesFile.setTrashed(true);
  }
  
  // if has notion url, perform insert
  const notionUrl = cpDataObj.NotionUrl;
  if (notionUrl) {
    appendBlockToNotion(notionUrl, outputFile, undefined, cpDataObj);
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

  var outputFile = doc;

  if (cpDataObj.Type === SUPPORTED_TYPE.DOC_TO_PDF) {
    doc.setName(doc.getName() + ".pdf");
    outputFile = convertToPdf_(doc, destinationFolder); // Convert the doc to a PDF file.
    const docFile = DriveApp.getFileById(doc.getId()); // Get the temporary Google Docs file.
    docFile.setTrashed(true); // Trash the temporary Google Docs file.
  }

  // if has notion url, perform insert
  const notionUrl = cpDataObj.NotionUrl;
  if (notionUrl) {
    appendBlockToNotion(notionUrl, outputFile, undefined, cpDataObj);
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

  // if has notion url, perform insert
  const notionUrl = cpDataObj.NotionUrl;
  if (notionUrl) {
    var recipent = emails.toEmails;
    if (emails.ccEmails) recipent += ` cc: ` + emails.ccEmails;
    if (emails.bccEmails) recipent += ` bcc: ` + emails.bccEmails;
    appendBlockToNotion(notionUrl, undefined, `${recipent} (${subject})`, cpDataObj);
  }

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

/*
  finds all potential variables from Google Docs/Sheets/Slides templates and returns an array of unique variables
*/

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
    form.setConfirmationMessage("Thank you for using Form-mation!\n\nExpect the email to be sent within the next 3 minutes.");
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

  if (value.length == 0) {
    return `<i style="font-size:12px;">There seems to be no history. Create a process if you haven't already and validate!</i>`;
  }

  return output;
}

function clearActivityHistory() {
  PropertiesService.getScriptProperties().deleteProperty(PROPERTY_ACTIVITY_LOG);
}

function openInstructionPage(pageUrl) {
  const title = 'Help Guide';
  var template = HtmlService.createTemplateFromFile('iframe-page');
  template.pageUrl = pageUrl;
  template.title = title;
  var htmlOutput = template.evaluate()
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}

/*
  Notion Integration
*/

function openNotionKeyPrompt() {
  return openInputKeyPrompt({
    propertyName: PROPERTY_NOTION_API_KEY,
    keyName: "Notion API Key",
    getKeyUrl: "https://www.notion.so/profile/integrations",
    keySample: "secret_kuEMMGafwt2kympeBjK8ttxwm78uhgtdMlkCLfEsYEC"
  })
}

function appendBlockToNotion(notionUrl, file, stringText = undefined, cpDataObj = undefined) {
  const notionApiKey = PropertiesService.getScriptProperties().getProperty(PROPERTY_NOTION_API_KEY);
  var blockId = extractNotionBlockId(notionUrl);

  if (!notionApiKey) {
    openNotionKeyPrompt();
    return;
  }

  if (!blockId) {
    Logger.log(`Notion URL invalid: ${notionUrl}`);
    return false;
  }

  const res = isSupportChildBlockNotion(notionUrl);
  if (res.message) return; //there was an unexpected error, abort operation

  const blockType = res.type;

  const endpoint = `https://api.notion.com/v1/blocks/${blockId}/children`;

  var outputName;
  var outputLinkObj = null;

  if (file) {
    outputName = file.getName();
    outputLinkObj = { url: file.getUrl() };
  } else if (stringText) {
    outputName = stringText;
  }

  const headers = {
    'Authorization': `Bearer ${notionApiKey}`,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28'
  };

  var data = {
    children: [
      {
        object: 'block',
        type: 'paragraph',
        paragraph: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: outputName,
                link: outputLinkObj
              }
            }
          ]
        }
      }
    ]
  };

  if (blockType === 'table' && res.table_width) {
    var data_colums = [];

    if (res.table_width >= 2) {
      data_colums.push([
        {
          type: 'text',
          text: {
            content: getCurrentDateTimeString(),
            link: null
          }
        }
      ]);
    }

    if (res.table_width >= 3) {
      const typeDict = {
        "Email": "Email",
        "Doc-to-Doc": "Google Docs",
        "Doc-to-PDF": "PDF",
        "Slide-to-Slide": "Google Slides",
        "Slide-to-PDF": "PDF",
        "Sheet-to-Sheet": "Google Sheets",
        "Sheet-to-PDF": "PDF"
      }
      data_colums.push([
        {
          type: 'text',
          text: {
            content: typeDict[cpDataObj.Type],
            link: null
          }
        }
      ]);
    }

    data_colums.push([
      {
        type: 'text',
        text: {
          content: outputName,
          link: outputLinkObj
        }
      }
    ]);

    if (res.table_width > 3) {
      for (var i = 3; i < res.table_width; i++) {
        data_colums.push([
          {
            type: 'text',
            text: {
              content: '',
              link: null
            }
          }
        ]);
      }
    }

    console.log({ blockType, tw: res.table_width, data_colums })

    data = {
      children: [
        {
          object: 'block',
          type: 'table_row',
          table_row: {
            cells: data_colums
          }
        }
      ]
    };
  }

  const options = {
    method: 'patch',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseData = JSON.parse(response.getContentText());
    console.log('Block added:', responseData);
    return true;
  } catch (error) {
    console.error('Error adding block:');
    console.log(JSON.stringify(error));
    newActivityLog(`Notion insertion for process '${cpDataObj.Name}' failed!`);
    return false;
  }
}

/*
  Gemini Integration
*/

function openGeminiKeyPrompt() {
  return openInputKeyPrompt({
    propertyName: PROPERTY_GEMINI_API_KEY,
    keyName: "Gemini API Key",
    getKeyUrl: "https://aistudio.google.com/app/apikey",
    keySample: "AIz124CrPasyiPTVcZxsr-dinuertTw-P229bQc"
  });
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
  
  SpreadsheetApp.getUi().showModalDialog(ui, 'Co-Create Process with Gemini');
}

function runGemini(selectedType, textDescription, rerun = 0) {
  var prePrompt, prompt;
  if (
    selectedType === SUPPORTED_TYPE.DOC_TO_DOC ||
    selectedType === SUPPORTED_TYPE.DOC_TO_PDF ||
    selectedType === SUPPORTED_TYPE.EMAIL
  ) {
    prePrompt = `Create a document template with up to a maximum of 10 placeholders within the template content, surrounded by double curly brackets ({{}}). Placeholders should be named descriptively (e.g., {{Date}}, {{Company Name}}). Adhere to the specified template type (e.g., letter, report, email). Use double asterisks (** **) for bold and double underscores (__ __) for underline formatting.

Generate a suitable file name without extensions, preferably with spaces, or a subject line for email templates. Create a process label name based on the provided description.

Output format:
<TEMPLATE_CONTENT>[template content (Do not include subject name)]</TEMPLATE_CONTENT>
<FILE_NAME>[suggested file name]</FILE_NAME>
<EMAIL_SUBJECT>[if type is "Email", suggested email subject]</EMAIL_SUBJECT>
<PROCESS_NAME>[suggested short and concise process label name]</PROCESS_NAME>
`;
  prompt = `${prePrompt} Type: "${selectedType}", Description: "${textDescription}"`;
  } else if (
    selectedType === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
    selectedType === SUPPORTED_TYPE.SLIDE_TO_PDF
  ) {
    prompt = `Prepare a slide deck for the purpose of ${textDescription}. Generate between 5 and 15 slides, adjusting the number based on the complexity of the topic. Please generate main bullet points (up to 5 per slide) with between 1-3 placeholders each, surrounded by double curly brackets ({{}}). Placeholders should be embedded within sentences, not standalone. Placeholders should be named descriptively, preferably with space. Avoid nested placeholders. Keep the title of each slide short and concise, must not include or indicate which slide the title is for. Generate a suitable file name with at least 1 placeholder for this deck, without extensions, preferably with spaces. Create a process label name based on the purpose. Create a suitable topic as text shown before the deck. Please produce the result as a valid JSON (e.g., {"topic": "[topic]", "fileName": "[file name]", "processLabel": [process label name]", "slides": [{"title": "[slide 1 title]","body": []},{"title": "[slide 2 title]","body": []}]}) so that I can pass it to other APIs.`;
  }

  if (rerun > 0 && rerun < 5) {
    prePrompt = prePrompt + " It seems like the previous generation output did not satisfy the output format requirements, ensure that does not occur this time.";
  } else if (rerun >= 5) return false;

  var geminiOutput = callGemini(prompt);
  if (callGemini === false) return false;
  Logger.log({selectedType, textDescription, rerun})
  Logger.log({geminiOutput});

  var fileUrl, processName;

  if (
    selectedType === SUPPORTED_TYPE.DOC_TO_DOC ||
    selectedType === SUPPORTED_TYPE.DOC_TO_PDF ||
    selectedType === SUPPORTED_TYPE.EMAIL
  ) {
    var fileName, subjectTitle, templateContent;

    templateContent = geminiOutput.match(/(?<=<TEMPLATE_CONTENT>).*(?=<\/TEMPLATE_CONTENT>)/s);
    processName = getGeminiOutputContent("PROCESS_NAME", geminiOutput);
    fileName = getGeminiOutputContent("FILE_NAME", geminiOutput);
    subjectTitle = getGeminiOutputContent("EMAIL_SUBJECT", geminiOutput);

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

    const doc = DocumentApp.create(fileName);
    const docId = doc.getId();
    DriveApp.getFileById(docId).moveTo(getProjectFolder());
    fileUrl = `https://docs.google.com/document/d/${docId}/edit`;

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
  } else if (
    selectedType === SUPPORTED_TYPE.SLIDE_TO_SLIDE ||
    selectedType === SUPPORTED_TYPE.SLIDE_TO_PDF
  ) {
    // The Gemini model likes to enclose the JSON with ```json and ```
    geminiOutput = geminiOutput.replace(/```(?:json|)/g, "");
    // Remove potential bolding attempts by Gemini
    geminiOutput = geminiOutput.replace(/\*\*(?:|)/g, "");

    try {
      var geminiOutputObj = JSON.parse(geminiOutput);
      var { fileName, processLabel, topic, slides } = geminiOutputObj;
    } catch (e) {
      return runGemini(selectedType, textDescription, rerun + 1);
    }

    console.log(JSON.stringify(geminiOutputObj))

    processName = processLabel;

    // Create a Google Slides presentation.
    const presentation = SlidesApp.create(fileName);
    const fileId = presentation.getId();
    DriveApp.getFileById(fileId).moveTo(getProjectFolder());
    fileUrl = `https://docs.google.com/presentation/d/${fileId}/edit`;

    // Set up the opening slide.
    var slide = presentation.getSlides()[0]; 
    var shapes = slide.getShapes();
    shapes[0].getText().setText(topic);

    var body;
    for (var i = 0; i < slides.length; i++) {
      slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      shapes = slide.getShapes();
      // Set title.
      shapes[0].getText().setText(slides[i]['title']);
  
      // Set body.
      body = "";
      for (var j = 0; j < slides[i]['body'].length; j++) {
        body += '' + slides[i]['body'][j] + '\n';
      }
      shapes[1].getText().setText(body);
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