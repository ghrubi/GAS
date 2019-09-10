// -----------------------------------------------------
// Globals, contants
// -----------------------------------------------------
LABEL_PENDING = "newOrder";  // Comment for testing
LABEL_DONE = "newOrderDone";  // Comment for testing

//LABEL_PENDING = "newOrderTest";  // Uncomment for testing
//LABEL_DONE = "newOrderDoneTest";  // Uncomment for testing

REPEAT_CALL_DELAY = 1000 * 60 * 5; // Define 5 minutes in milliseconds
ADVANCE_ORDER_DELAY = 1000 * 60 * 30; // Define 15 minutes in milliseconds

// Get active sheet
var sheet = SpreadsheetApp.getActiveSheet();

// Get current time.
var currTime = new Date().getTime();
 
// Get the script properties.
var scriptProperties = PropertiesService.getScriptProperties();

function myFunction() {
  // Define email labels
  var label_pending = GmailApp.getUserLabelByName(LABEL_PENDING);
  var label_done = GmailApp.getUserLabelByName(LABEL_DONE);
  
  // The threads currently assigned to the 'pending' label
  var threads = label_pending.getThreads();

  // Process each one in turn, assuming there's only a single
  // message in each thread
  for (var t in threads) {
    var skipEmail = false;  // Boolean Flag for orders placed in the future
    var thread = threads[t];
   
    var emailElements = parseEmail(thread);
    
    var orderType = emailElements[0];
    var phoneNumber = emailElements[1];
    var orderDueTime = emailElements[2];  // Only for ADVANCE orders.
    
//    sheet.appendRow([emailElements[0],emailElements[1],emailElements[2]]);

    // Check to see if the ADVANCE order is due in less than ADVANCE order
    // delay period. If so, make the call. Otherwise, just skip it.
    // ASAP orders just call.
    if(orderType == 'ADVANCE') {
      // Check the time diff. Skip if more than order delay
      var deltaTime = orderDueTime - currTime;
      if(deltaTime > ADVANCE_ORDER_DELAY) {
        skipEmail = true;
      }
    }
    
    // If email isn't skipped, move to done label and make call.
    if(!skipEmail) {
      // Set to 'done' by exchanging labels
      thread.removeLabel(label_pending);
      thread.addLabel(label_done);
    
      // Initiate phone call notification
      sendCall(phoneNumber);
    }
  }
}

// --------------------------------------------------------
// Looks to see if phone number has been called within
// the REPEAT_CALL_DELAY. If not, update delay time for 
// phone number and call. Otherwise, just don't call.
// --------------------------------------------------------
function sendCall(phoneNumber) {
//  sheet.appendRow(['currTime:', currTime]);
//  sheet.appendRow(['sendNotification:', phoneNumber, scriptProperties.getProperty(phoneNumber)]);
  
  // Call delay passed or haven't initiated a call yet to number?
  // Otherwise, skip it
  if(scriptProperties.getProperty(phoneNumber) < currTime) {
  //  sheet.appendRow(['Calling:', phoneNumber, scriptProperties.getProperty(phoneNumber)]);
    
    // Update or set next call delay time for phone number
    scriptProperties.setProperty(phoneNumber, currTime + REPEAT_CALL_DELAY);
  //  sheet.appendRow(['Updated:', phoneNumber, scriptProperties.getProperty(phoneNumber)]);
    
    // Initiate phone call notification
    var response = UrlFetchApp.fetch("https://api.tropo.com/1.0/sessions?action=create&token=6577757373717649707753477464436c7373636e68666878706a67637a69636e4d79546d6f67436f61446447&numberToDial=" + phoneNumber);
  }
}

// --------------------------------------------------------
// Parses out elements of the email body.
// 1) order type 2) store phone number
// If it's and ADVACE order, 3) is the due time. Otherwise,
// due time is set to 0.
// --------------------------------------------------------
function parseEmail(thread) {
   // Get the first email message of a thread. There should only be one per thread.
    var tmp,
      message = thread.getMessages()[0],
      subject = message.getSubject(),
      body = message.getBody();
//      sheet.appendRow([subject]);
  
  var orderType = emailOrderType(body);
  var phoneNumber = emailOrderPhoneNumber(body);
//  sheet.appendRow([orderType, phoneNumber]);
  
  var timeDue = 0;
  if(orderType == 'ADVANCE') {
    timeDue = emailOrderDueTime(body);
//    sheet.appendRow([timeDue]);
  }
 
  // Return as array
  return ([orderType, phoneNumber, timeDue]);
}

// --------------------------------------------------------
// Looks for *ADVANCE* in email body. Otherwise, the 
// order is for ASAP. Returns which type it is.
//
// --------------------------------------------------------
function emailOrderType(emailBody) {
    // Look for *ADVANCE* in email body first
    var tmp;

    tmp = emailBody.match(/\*ADVANCE\*/);

    // If found, set to ADVANCE. Otherwise, the other order type of ASAP.
    if(tmp != null) {
      tmp = 'ADVANCE';
    }
    else {
      tmp = 'ASAP';
    }
  
    return tmp;
}

// --------------------------------------------------------
// Extracts the phone number for the Store from the 
// email body. '-'s are removed and a '1' is 
// prepended.
// --------------------------------------------------------
function emailOrderPhoneNumber(emailBody) {
  // Get the first phone number within email body. This is only one we need.
  var tmp;
  var tmpArray;
  var phoneNumber;
  
  // Try Phone Number with '-'
  tmp = emailBody.match(/[0-9]{3}\-[0-9]{3}\-[0-9]{4}/);
//  sheet.appendRow(['In phone number function: ', tmp[0]]);
  
  // If null, must be Phone Number with '.'
  if(tmp == null) {
    tmp = emailBody.match(/[0-9]{3}\.[0-9]{3}\.[0-9]{4}/);  
    
    // Split on '.'
    tmpArray = tmp[0].split('.');
  }
  else {
    // Split on '-'
    tmpArray = tmp[0].split('-');
  }

  
  // Reassemble phone number with prepended 1.
  phoneNumber = '1' + tmpArray[0] + tmpArray[1] + tmpArray[2];
  
  return phoneNumber;
}

// --------------------------------------------------------
// Extracts the order due time from an ADVANCED order
// email body. The date and time is converted to proper
// UNIX time and returned.
// --------------------------------------------------------
function emailOrderDueTime(emailBody) {
  // Find time stamp within email. The first one is the Pickup Time. Only grab the first one.
  var tmStamp = emailBody.match(/\s*[0-9]{2}\-[0-9]{2}\-[0-9]{4}\s[0-9]{2}\:[0-9]{2}\s[AaPp][Mm]/);
  var arrayTmStamp = tmStamp[0].split(' ');
//      sheet.appendRow([subject,tmp[0],tmStamp[0],arrayTmStamp[0],arrayTmStamp[1],arrayTmStamp[2]]);
      
  // Separate out date elements and time elements. Convert to integers from strings.
  var arrayDate = arrayTmStamp[0].split('-');
  var tsMonth = parseInt(arrayDate[0],10);
  tsMonth += -1; // Jan is 0
  var tsDay = parseInt(arrayDate[1],10);
  var tsYear = parseInt(arrayDate[2]);
      
  var arrayTime = arrayTmStamp[1].split(':');
  var tsHour = parseInt(arrayTime[0],10);
  var tsMin = parseInt(arrayTime[1],10);
      
  var amOrPm = arrayTmStamp[2];

//      sheet.appendRow(['After split',arrayDate[0],arrayDate[1],arrayDate[2],arrayTime[0],arrayTime[1],amOrPm]);
//      sheet.appendRow(['Split vars',tsMonth,tsDay,tsYear,tsHour,tsMin,amOrPm]);
      
  // PM and not noon, add 12 for 24hour time
  if (amOrPm == 'PM' && tsHour != 12) {
//        sheet.appendRow(['In hour add 12',tsHour]);
    tsHour += 12;
//        sheet.appendRow(['In hour add 12',tsHour]);
  }
  var orderDueTime = new Date(tsYear, tsMonth, tsDay, tsHour, tsMin, 0).getTime();
//  sheet.appendRow(['Order Due Time: ', orderDueTime]);
  return orderDueTime;
}

