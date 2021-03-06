function every_two_hours() {
  var data = readData();
  var reqs = buildRequests(data);
  for (i in reqs)
  {
    var req = reqs[i];
  }
  action(reqs, false);
}


function every_day() {
  var data = readData();
  var reqs = buildRequests(data);
  var team = action(reqs, true);
  sendReminders(team);
}


function readData() 
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var startCol = 1;
  var lastRow = sheet.getLastRow(); // Last row of data to process
  var lastColumn = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, 1, lastRow, lastColumn);
  Logger.log('OK, read '+lastRow+' rows')
  return dataRange.getValues();
}


function buildRequests(rows) 
{
  var requests = [];
  for (i in rows) 
  {
    if (rows[i][2] == ''){continue;}
    var row = rows[i];    
    var request = {}
    request['input'] = row[0]; //A
    request['notes'] = row[1]; //B
    request['applicationDate'] = row[2]; //C
    request['username'] = row[3]; //D
    request['userInitial'] = row[3].slice(1, 2);
    request['userID'] = row[4]; //E
    request['usersName'] = row[5]; //F
    request['creditHours'] = row[6]; //G
    request['term'] = row[7]; //H
    request['SAPAck'] = row[11]; //L
    request['fin'] = row[12]; //M
    request['finAck'] = row[13]; //N
    request['plansAck'] = row[14]; //O
    request['nationalityStatus']; // Deprecated
    request['isEmailSent'] = row[15]; //P
    request['reminderEmailSentTo'] = row[16]; //Q
    if (request['input'] == 'Y') {
      request['decision'] = 'Approved';
    }
    else if (request['input'] == 'N') {
      request['decision'] = 'Denied';
    }
    request['owner'] = 'Jennifer';
    requests.push(request);
    Logger.log('Added student '+request['usersName']+'; Response: '+request['decision']);
  }
  return requests;
}



function action(requests, remindersMode) {
  var startRow = 2; // First row of data to process
  var sheet = SpreadsheetApp.getActiveSheet();
  var owners = {
    'Sarah': {
      'email': 'sfroberg@berklee.edu',
      'reminder': false},
    'Andrea': {
      'email': 'atikofsky@berklee.edu',
      'reminder': false},
    'Gosia': {
      'email': 'mtorzecka@berklee.edu',
      'reminder': false},
     'Jennifer': {
      'email': 'jbattista@berklee.edu',
      'reminder': false}
  }
  for (i in requests) 
  {
    var req = requests[i];
    if (req['decision'] && req['isEmailSent'] == '') {
      var email = writeResponse(req);
      Logger.log('Sending response to '+req['username'])
      MailApp.sendEmail(req['username'], email[0], email[1], email[2]);
      var d = new Date();
      var currentDate = d.toLocaleDateString(); //"December 19, 2014" for instance  
      var currentTime = d.toLocaleTimeString(); // "12:35 PM", for instance
      sheet.getRange(startRow+parseInt(i, 10), 16).setValue(req['decision']+' - sent to '+req['username']+' on ' +
                                                            currentDate + ' at ' + currentTime);
    }
    else if (remindersMode)
    {
      if (req['input'] == '' && req['reminderEmailSentTo'] == '') 
      {
        var d = new Date();
        var currentDate = d.toLocaleDateString(); //"December 19, 2014" for instance  
        var currentTime = d.toLocaleTimeString(); // "12:35 PM", for instance
        sheet.getRange(startRow + parseInt(i, 10), 17).setValue(req['owner']+
                                                              ' on '+currentDate+
                                                              ' at '+currentTime);
        Logger.log('Sending reminder to'+req['owner']);
        owners[req['owner']]['reminder'] = true;
      }
    }
    else 
    {
      Logger.log('Skipping request from '+req['username']+' because a '+req['decision']+' response had already been sent');
    }
    SpreadsheetApp.flush();
  }
  return owners
}


function writeResponse(record) {
  if (record['decision'] == 'Approved' ) {
    var subject = "Part-Time Request Approved";
    var options = {
      cc: "registrar@berklee.edu, counselingcenter@berklee.edu, "+
      "financialaid@berklee.edu, scholarships@berklee.edu, success@berklee.edu, "+
      "enrollment@berklee.edu, bursar@berklee.edu",
    }
    // Prevents sending duplicate emails and emails to non-approved rows. 
    var message = 'Student ID# '+record['userID']+
                '\n\nDear '+record['usersName']+',\n\nYour request for part-time status ('+
                record['creditHours']+' credit-hours) beginning in '+record['term']+
                ' has been approved.\n\nAs part of your request for part-time status you agreed to or acknowledged '+
                'the following:\n\n1. Satisfactory Academic Progress (SAP) Policy acknowledgment: '+record['SAPAck']+
                '\n\n2. Stated financial aid status: '+record['fin']+'\n\n3. Statement of financial responsibility: '+
                record['finAck']+'\n\n4. Notification when plans change: '+record['plansAck']+'\n\n5. Berklee status: '+
                'U.S. citizen or permanent resident.\n\nIf you anticipate a problem fulfilling any of the above, please '+
                'contact the appropriate department immediately.\n\nBest wishes for a successful semester!\n\nMichael '+
                'Hagerty\n\nRegistrar | Berklee College of Music\nEnrollment Division\n1140 Boylston Street MS-921-OREG\n'+
                'Boston, MA 02215 | 617-747-2240'

  }
  else if (record['decision'] == 'Denied') {
    var options = {
      cc: 'registrar@berklee.edu',
      bcc: 'enrollment@berklee.edu',
    }
    var subject = "Part-Time Request Denied";
    var message = 'Student ID# '+record['userID']+'\n\nDear '+
          record['usersName']+',\n\nYour request for part-time status '+
          'for the '+record['term']+' semester was not approved. This decision '+
          'was based on the following: \n\n'+record['notes']+'\n\nIf you '+
          'believe that your application was rejected in error or if you require '+
          'further information, please stop by the Office of the Registrar '+
          '(921 BOYL, Suite 120) or email us at (registrar@berklee.edu).\n\nBest '+
          'wishes for a successful semester.\n\nMichael Hagerty\n\nRegistrar | '+
          'Berklee College of Music\nEnrollment Division\n1140 Boylston Street '+
          'MS-921-OREG\nBoston, MA 02215 | 617-747-2240'
  }
  return [subject, message, options];
}


function sendReminders(owners) {
  var subject = "Part-Time Authorization Requests - Action Required";
  var options = {}
  for (i in owners) {
    var owner = owners[i];
    if (owner['reminder']) {
      Logger.log(owner['email'])
      var d = new Date();
      var currentDate = d.toLocaleDateString(); //"December 19, 2014" for instance  
      var currentTime = d.toLocaleTimeString(); // "12:35 PM", for instance
      var message = 'To the Office of the Registrar:\n\n'+
        'This is an automated notification from the Part-Time '+
        'Student Authorization Form response center that you have '+
        'part-time authorization requests to approve. Please navigate '+
        'to the responses spreadsheet and enter either \'Y\' or \'N\' in '+
        'the Approved column for every student request. \n\nThis email '+
        'reminder will not be sent again until there are new form submissions. '+
        'Please be sure to keep the Approved column updated for each authorization '+
        'request that arrives. \n\nKind regards, \n\nPSM';
      MailApp.sendEmail(owner['email'], subject, message, options);
      SpreadsheetApp.flush();
      // Make sure the cell is updated right away in case the script is interrupted
    }
  }
}
