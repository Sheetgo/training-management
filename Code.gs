/*================================================================================================================*
  Training Management by Sheetgo
  ================================================================================================================
  Version:      1.0.0
  Project Page: 
  Copyright:    (c) 2018 by Sheetgo
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  ----------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release
 *================================================================================================================*/

/**
 * Template file id and names.
 * This configuration changes after the script copy the template files
 * @type
 */
Files = {
  // Form linked
  Form_Training_Management: { id: null, name: "Schedule Your Training" },
  
  // Main Spreadsheet
  Ss_Training_Management: { id: null, name: "Training Management" }
}

/**
 * Creates the 'Training Management' Menu in the spreadsheet. This function is fired every time a spreadsheet is open
 * @param {JSON} e User/Spreadsheet basic parameters 
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  var menu = ui.createMenu('Training Management')
  if (e && e.authMode == ScriptApp.AuthMode.LIMITED) {
    menu.addItem('Create system', 'createSystem');
  } else {
    menu.addItem('Update Training Form', 'updateTrainingForm');
    menu.addItem('Set trainings to calendar', 'setTrainingsToCalendar');
  }
  menu.addToUi();
}


/**
 * Create the Training Management Invoice system by copying the template files and moving into an local 
 * user folder within Google Drive
 */
function createSystem() {

  // Set Main Spreadsheet ID to Files settings
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Files.Ss_Training_Management.id = spreadsheet.getId();
  
  // Set Main Form ID to Files settings
  var formURL = spreadsheet.getFormUrl();
  Files.Form_Training_Management.id = FormApp.openByUrl(formURL).getId();
  var x = Files.Form_Training_Management.id;
  
  spreadsheet.toast("Creating & Configuring Solution. Please wait...");

  // Create the Solution folder on users Drive 
  var folder = DriveApp.createFolder("Training Management");

  // Move the current Dashboard spreadsheet into the Solution folder
  var file = DriveApp.getFileById(Files.Ss_Training_Management.id);
  file.setName(Files.Ss_Training_Management.name);
  moveFile(file, folder);
  
  // Move the current Dashboard spreadsheet into the Solution folder
  var form = DriveApp.getFileById(Files.Form_Training_Management.id);
  form.setName(Files.Form_Training_Management.name);
  moveFile(form, folder);

  // Update menu
  toggleTrigger();
  onOpen();
}   

/**
 * Update all form data base on Settings sheet data
 */
function updateTrainingForm() {
  updateSubjects();
  updateLocations();
}

/**
 * Update dropdown of Subjects on linked form based on Settings sheet data
 */
function updateSubjects() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var items = ss.getSheetByName("Trainings").getDataRange().getValues().map(function(item){
    return item[0];
  })
  
  items.shift();
  
  var form = FormApp.openByUrl(ss.getFormUrl()); 
  var subjects = form.getItems()[4]; 
 
  var listItem = subjects.asListItem(); 
  
  var choices = listItem.getChoices();
  choices.length = 0;
  
  for (var i in items){
    choices.push(listItem.createChoice(items[i]))  
  }
  
  listItem.setChoices(choices)
}
 
/**
 * Update dropdown of Locations on linked form based in Settings sheet data
 */
function updateLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var items = ss.getSheetByName("Locations").getDataRange().getValues().map(function(item){
    return item[0];
  })
  
  items.shift();
  
  var form = FormApp.openByUrl(ss.getFormUrl()); 
  var locations = form.getItems()[5]; 
 
  var listItem = locations.asListItem(); 
  
  var choices = listItem.getChoices();
  choices.length = 0;
  
  for (var i in items){
    choices.push(listItem.createChoice(items[i]))  
  }
  
  listItem.setChoices(choices)
}

/**
 * Send email to manager with a list of new trainings requests
 */
function sendEmailToManager(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_timezone = ss.getSpreadsheetTimeZone();
  var email_manager = ss.getSheetByName('Settings').getRange("F8").getValue();
  var template = HtmlService.createTemplateFromFile("email_template");
 
  var ss_form_responses = ss.getSheetByName('Training Approval');
  var form_data = ss_form_responses.getDataRange().getValues();

  var new_requests = new Array();
  for (i=1; i<form_data.length; i++) {
    var row = form_data[i];
    var approved_at = row[8];
    if (approved_at !== "") {
      continue;
    }
    var training = row[1];
    var start_date = row[2];
    var end_date = row[3]; 
    var location = row[4];
    var requester = row[5];
    
    new_requests.push({
      training: training,
      start_date: Utilities.formatDate(start_date, ss_timezone, 'MM/dd/yyyy HH:mm'),
      end_date: Utilities.formatDate(end_date, ss_timezone, 'MM/dd/yyyy HH:mm'),
      location: location,
      requester: requester,
    });
  }
  
  if (new_requests.length > 0){
    template.trainings_requests = new_requests;
    template.spreadsheet_link = ss.getUrl();
    var email_content = template.evaluate().getContent();
    
    if (validateEmail(email_manager)) {
      MailApp.sendEmail(email_manager, "Training Request", "", {'htmlBody': email_content});
    }
  }
  
}
  
 
/**
 * Access data from a worksheet to create events in Calendar-related dates
 */
function setTrainingsToCalendar() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var ss_form_responses = ss.getSheetByName('Training Approval');
  var form_data = ss_form_responses.getDataRange().getValues();
  
  var ss_settings = ss.getSheetByName('Settings');
  var calendar_id = ss_settings.getRange('F11').getValue();
  var calendar = CalendarApp.getCalendarById(calendar_id);
  
  // Fetch values for each row in the Range.
  for (i=1; i < form_data.length; i++) {
    var row = form_data[i];
    var updated = row[8];    
    var is_approved = row[7];
    
    if (updated == "" && is_approved == "Yes") {
      var training = row[1];
      var start_date = row[2]; 
      var end_date = row[3]; 
      var location = row[4];
      var event_title = training + " - " + location;
      
      var event = calendar.createEvent(
        event_title, 
        start_date,
        end_date,
        {
          location: location
        }
      );
       
      var split_event_id = event.getId().split('@');
      var eventURL = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(split_event_id[0] + " " + calendar_id);
      var event_url_cell = ss_form_responses.getRange(1 + i, 10);
      event_url_cell.setValue(eventURL);
    }
    
    var updated_at_cell = ss_form_responses.getRange(1 + i, 9);
    updated_at_cell.setValue(new Date());
  }
}

/**
 * Validate if a email is valid
 * @param {String} email A valid email
 */
function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  if (!re.test(email)) {
    return false;
  } else {
    return true;
  }
}

/**
 * Move a file from one folder into another
 * @param {Object} file A file object in Google Drive
 * @param {Object} dest_folder A folder object in Google Drive 
 */
function moveFile(file, dest_folder) {
    dest_folder.addFile(file);
    var parents = file.getParents();
    while (parents.hasNext()) {
        var folder = parents.next();
        if (folder.getId() != dest_folder.getId()) {
            folder.removeFile(file)
        }
    }
}


/**
 * Switch on the trigger that will run on form submit to send e-mail to Manager
 */
function toggleTrigger() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet()
        ScriptApp.newTrigger('sendEmailToManager')
            .forSpreadsheet(sheet)
            .onFormSubmit()
            .create()

    } catch (e) {
        Logger.log(e);
    }
}