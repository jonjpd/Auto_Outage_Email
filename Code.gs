// Section 1. describes the addition of the spreadsheet toolbar menu

function onOpen() {                       // Native function to google apps that runs as the name implies - once it is opened.
  var ui = SpreadsheetApp.getUi();      

  ui.createMenu('Send NNI Outage Email')  // Sets the name of the Toolbar menu button.
  .addItem('Test NNI Emails', 'TEST_NNI') // First argument specifies the menu's text. Second argument denotes the name of the function that the menu calls. 
    .addSeparator()
  .addSubMenu(ui.createMenu('DUB')        // Example of a sub menu allowing grouped options for multiple NNIs per PoP
    .addItem('LDN', 'DUB_LDN')
    .addItem('AMS', 'DUB_AMS')
    .addItem('FRA', 'DUB_FRA')
    .addItem('PAR', 'DUB_PAR'))
    .addSeparator()                       // For Style: Adds a line separating each option.
  .addToUi();
}


// Section 2. lists the functions that the user's Menu selection calls.

function TEST_NNI() {
  confirmWindow('TEST_NNISend');          // Calls the function confirmWindow explained in the next section.
}

function TEST_NNISend() {
  sendMail('TEST_NNI', "TestA-TestZ");    // The first argument calls the specific 'Named Range' in the spreadsheet. 
}                                         // The second argument calls the NNI name for the specific entry. This will be used in the email subject and body.

function DUB_LDN() {
  confirmWindow('DUB_LDNSend');
}

function DUB_LDNSend() {
  sendMail('DUB_LDN', "DUB-LDN");         // Example shows Dublin to London NNI. Each subsequent NNI will require a function pair.
}


// Section 3. explains the confirmation window function.

function confirmWindow(NNIsend) {
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Confirm').setHeight('100').setWidth('400');     // These lines creates a confirmation dialog window to prevent accidental emails being sent through user error.
  var panel = app.createVerticalPanel().add(app.createHTML('Please confirm that you entered all information correctly.'));
  var grid = app.createGrid(1,2).setWidth('400');
  var closeHandler = app.createServerHandler('close');
  var approvalHandler = app.createServerHandler(NNIsend);                                       // Calls on the next function if the user selects the 'Send' button created on the next line of code.
  var b1 = app.createButton("Yes, send out this NNI email.", approvalHandler).setTitle('Send');
  var b2 = app.createButton("Cancel, let me recheck my details.",closeHandler).setTitle('Close this window');
  var G1 = app.createVerticalPanel().add(b1).add(b2);
  grid.setWidget(0,0,G1).setWidget(0,1,b2);
  app.add(panel).add(grid);
  doc.show(app);

}

function close(){            // Closes the window and stops the function if the user selects the 'Close this window' button.
  return UiApp.getActiveApplication().close();
}


// Section 4. presents how the email is built and distributed to the correct customers.

function sendMail(range, NNI) {                               // Function calls the arguments specified in section two. Arg 1. The manually defined 'named ranged' of the spreadsheet. 
                                                              // Arg 2. The non-named range version of the affected locations. Example '_111-165' would become '111-165' for customer eligibility and clarity.
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var update = ss.getRangeByName('latestUpdate').getValue();  // Grabs the 'latestUpdate' named range from the spreadsheet. 
  var noc = "noc@yourcompany.com";                            // Set the email you would like CC'd - for example your NOC's email address.
  var rangeValue = ss.getRangeByName(range).getValues();      // Creates an array of the specified email values.

  var LogoCompany = UrlFetchApp.fetch("http://imgur.com/gallery/y7Hm9").getBlob().setName("LogoBlob"); // This photo will be used as the example company's logo :)

  var joinedEmails = rangeValue.join(", ");                   // Joins every email in the named range with a comma and space
  
  var ticketNum = ss.getRangeByName('ticketUpdate').getValue();          // Grabs the ticket and update number from 'ticketUpdate' named range.
  var subject = "$Company NOC Notification " + NNI + " Outage " + ticketNum; // Sets the subject of the email
   
  var body = "<br/><br/><img src='cid:Logo'><br/><br/><h1>" + subject + "</h1><br/><br/><br/>Dear Customer,<br/><br/><b>Please be advised that $Company NOC is currently alerted to an outage on the " + NNI + 
    " path. We are investigating this issue for you.</b><br/><br/><br/><br/>The latest update will appear below:<br/><br/>" + "<b><font color='red'>" + update + 
    "</b></font>" + "<br/><br/><br/><br/>" + ss.getRangeByName('oldUpdates').getValue() + 
    "<br/><br/><br/>Regards,<br/>$Company NOC<br/>"
  
  
  var bodyPlain = body.replace(/(<([^>]+)>)/ig, "");          // Gets rid of HTML tags in case the recipient device cannot process HTML.
 
  GmailApp.sendEmail(noc, subject, bodyPlain, {bcc: joinedEmails, htmlBody: body, inlineImages: 
         { Logo: LogoCompany }});                                                         // Calls the sendEmail function which sends the emails to the joinedEmails array using a BCC format.
  
  var app = UiApp.createApplication().setTitle('Confirm').setHeight('100').setWidth('400');
  var panel = app.createVerticalPanel().add(app.createHTML('Your email has been successful.')); // Notifies the user that the emails have been successfully sent.
  var grid = app.createGrid(1,2).setWidth('400');
  var closeHandler = app.createServerHandler('close');
  var b2 = app.createButton("Okay",closeHandler).setTitle('Close this window');
  var G1 = app.createVerticalPanel().add(b2);
  grid.setWidget(0,0,G1).setWidget(0,1,b2);
  app.add(panel).add(grid);
  ss.show(app);
}
