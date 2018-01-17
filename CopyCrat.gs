/*
 * This app explores methods of automatically processing user-submitted form data in order to generate a personalized merge document, which is then emailed to the user
 * Heavy reliance on documentation from Google Apps Script site & Mukesh Chapagain's blog:
 * http://blog.chapagain.com.np/create-form-using-formapp-class-publish-form-as-webapp-email-save-responses-to-spreadsheet/
 */

var global = getGlobalFormObj();

function doGet() { // this function required for launching as Web App
    
    /*
     * The first steps involve locating the starter files:
     * Google Form, Google Sheet response destination, Google Doc Template
     * Below code also demonstrates how to change file names by ID.
     * Fellow developers of this app, feel free to personalize with your own files & naming conventions.
     */
    
    var formFile = DriveApp.getFileById(global.formId); // locating Form by ID
    formFile.setName('STF Solution: Team Account Expense Request'); // renaming Form
    
    var sheetFile = DriveApp.getFileById('SHEETID'); // locating Form-associated Sheet, per Amy's example
    sheetFile.setName('STF Solution: Team Account Expense Request Coder Challenge (Responses)'); //renaming Sheet
    
    var docFile = DriveApp.getFileById('DOCID'); // locating template Doc
    docFile.setName('STF Solution: Doc Template Team Account Expense Request'); // renaming Doc
    
    /*
     * This section comprises:
     * programmatically updating & personalizing/refining some form properties, e.g., title
     * looping through the form items in order to access each item's Title, ID and index
     * using the above item data to target the invoice url item with text validation
     * the item data is used later, as well, to point to which data the Doc template merge fields will pull in
     */
    
    var form = FormApp.openById(global.formId);
    // updating properties via chaining
    form.setTitle('Team Account Expense Request to School Athletic Booster Club (BABC)')
    .setDescription('This request will be reviewed by the AD before any purchase can be made.')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setConfirmationMessage(
                            'Thank you. Your request will be forwarded to the AD and, if approved, submitted to the BABC. Check your email for a courtesy copy of your request.');
    // if necessary to change Form destination to a different spreadsheet, one can use form.setDestination(FormApp.DestinationType.SPREADSHEET, '123xxxxx');
    
    // used this function to get the list of choices in the Sport qustion sorted alphabetically
    var sportList = form.getItems()[1];
    var sportChoices = sportList.asListItem().getChoices();
    var unsortedChoices = [];
    for(var i in sportChoices){
        // Logger.log(sportChoices[i].getValue());
        unsortedChoices.push([sportChoices[i].getValue(),sportChoices[i]]);// add all existing sport choices
    }
    // Logger.log(unsortedChoices);
    
    unsortedChoices.sort(function(a,b){
        if (a[0] === b[0]) {
            return 0;
        } else {
            return (a[0] < b[0]) ? -1 : 1;
        }
    });
    
    var resultChoices = [];
    for(var n in unsortedChoices){
        resultChoices.push(unsortedChoices[n][1]);// creates a new array with sorted sport choices
    }
    //  Logger.log(resultChoices.toString());// check in logger
    sportList.asListItem().setChoices(resultChoices);// update the form
    
    var items = form.getItems();
    var docItemsArr = []; // created array to hold all form items apart from headers
    for (var i = 0; i < items.length; i++) { // looping through all the items in Form
        if (items[i].getType() !== FormApp.ItemType.SECTION_HEADER) { // excluding the section headers from the items array
            docItemsArr.push({
            title: items[i].getTitle(),
            id: items[i].getId(),
            index: docItemsArr.length
            });
        }
    }
    var urlItemId = docItemsArr[8].id;
    // Logger.log(urlItemId);
    var urlItem = form.getItemById(urlItemId);
    var textValidation = FormApp.createTextValidation()
    .requireTextIsUrl()
    .build();
    urlItem.asTextItem().setValidation(textValidation);
    
    var alphaSport = docItemsArr[0].id;
    
    var url = form.getPublishedUrl(); // our app will render the live Form for the user
    
    // Deletes all triggers in the current project - necessary before adding any new ones
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    ScriptApp.newTrigger('formSubmitToDoc').forForm(form).onFormSubmit().create(); // a new trigger to call formSubmitToDoc function when user submits Form data
    
    // Rendering Form via HtmlService
    // HUGE thanks to http://blog.chapagain.com.np/create-form-using-formapp-class-publish-form-as-webapp-email-save-responses-to-spreadsheet/
    return HtmlService.createHtmlOutput(
                                        "<form action='" + url + "' method='get' id='form'></form>" +
                                        "<script>document.getElementById('form').submit();</script>");
}

/*
 * Fetch submitted response from the Form
 * send to Doc template
 * save responses to Sheet
 */

function formSubmitToDoc(e) {
    // var formResponses = FormApp.openById(global.formId).getResponses();
    // Logger.log(formResponses);
    var eItemResponses = e.response.getItemResponses(); // when form is submitted, responses will be gathered and then fed into the document template
    var eResponseArray = [];
    for (var i = 0; i < eItemResponses.length; i++) {
        eResponseArray.push({ // adding properties to the response to make it easier to assign data to the corresponding template fields
        name: eItemResponses[0].getResponse(),
        sport: eItemResponses[1].getResponse(),
        payable: eItemResponses[2].getResponse(),
        reason: eItemResponses[3].getResponse(),
        amount: eItemResponses[4].getResponse(),
        account: eItemResponses[5].getResponse(),
        vendor: eItemResponses[6].getResponse(),
        description: eItemResponses[7].getResponse(),
        invoice: eItemResponses[8].getResponse()
        });
        
        Logger.log(eResponseArray);
    }
    
    var response = eResponseArray[0];
    
    var respondEmail = e.response.getRespondentEmail();
    
    var timestamp = e.response.getTimestamp();
    var timezone = Session.getScriptTimeZone();
    var formattedTime = Utilities.formatDate(timestamp, timezone, "MM/dd/yyyy" + '   ' + "h:mm a");
    
    var templateid = 'TEMPLATEID';
    var docid = DriveApp.getFileById(templateid).makeCopy().getId(); // preparing copy of the template doc which will be contain the form submission data
    var doc = DocumentApp.openById(docid);
    var body = doc.getActiveSection();
    
    var element = body.findText("<<receipt>>");
    var invoiceUrl = response.invoice;
    
    if(element){ // looks for the receipt url and transforms into a hyperlink
        var start = element.getStartOffset();
        var text = element.getElement().asText();
        text.replaceText("<<receipt>>", "Link to Invoice/Receipt");
        text.setLinkUrl(start, start+"Link to Invoice/Receipt".length-1, invoiceUrl);
    }
    
    body.replaceText('<<date>>', formattedTime)
    .replaceText("<<requester>>", response.name)
    .replaceText("<<sport>>", response.sport)
    .replaceText("<<payable to>>", response.payable)
    .replaceText("<<reason>>", response.reason)
    .replaceText("<<amount>>", response.amount)
    .replaceText("<<team account>>", response.account)
    .replaceText("<<vendor information>>", response.vendor)
    .replaceText("<<description>>", response.description)
    
    doc.saveAndClose();
    
    var sportName = response.sport;
    var submitName = response.name;
    
    
    doc.setName(sportName + ' Team Account Expense Request - ' + submitName + ' '+ formattedTime);
    
    var pdfFILE = DriveApp.getFileById(doc.getId());
    var blobFile = pdfFILE.getBlob().getAs('application/pdf').setName(doc.getName() + ".pdf");
    
    //  var pdfConfirmation = DriveApp.createFile(blobFile).getUrl(); if desired to save the pdf file that was emailed out
    
    MailApp.sendEmail({
    to: respondEmail,
    subject: response.sport + ' Team Account Expense Request',
    htmlBody: '<br>Dear ' +response.name + ', </br> <br>Attached is a copy of your Team Account Expense Request. </br> <br> Your request will be forwarded to the AD and, if approved, submitted to the BABC.</br> <br>Thank you for your submission!',
    attachments: [blobFile]
    });
    
    /*
     * Delete previously created triggers
     */
    function deleteTrigger() {
        var allTriggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < allTriggers.length; i++) {
            if (allTriggers[i].getHandlerFunction() == 'formSubmitToDoc') {
                ScriptApp.deleteTrigger(allTriggers[i]);
            }
        }
    }
    
}
