/* 
The functions below do three things:

1. onOpen creates a User Interface in the _23-24 OC FALL_SPRING CR/CA DATA. The user interface has one option, which is to run the updateCompletedCredits function.

2. showConfirmationDialog is a simple confirmation to the user who selects to update the Completed Credits sheet.

3. updateCompletedCredits looks for rows in the "Fall/Spring CR-CA Data" sheet that contain checks (are TRUE) in column A. If a check is there, then it will look to see if that row doesn't already exist in "Completed Credits". If it doesn't exist then it will insert the row into "Completed Credits". At the end of this function is calls the sendCounselorNotification function.

4. sendCounselorNotification sends an email to the student's counselor to let them know that the student completed a course.

Point of contact: Alvaro Gomez, Special Campuses Academic Technology Coach, at alvaro.gomez@nisd.net or 210-363-1577.
*/

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('UPDATE the Completed Credits Sheet')
    .addItem('Import new students from Fall/Spring CR-CA Data', 'showConfirmationDialog')
    .addItem('Watch instructions video', 'openInstructionsVideo')
    .addToUi();
}

function showConfirmationDialog() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert(
    'UPDATE the Completed Credits sheet',
    'Click yes to begin inserting new rows from the Fall/Spring CR-CA Data sheet into the Completed Credits sheet.',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    updateCompletedCredits();
    ui.alert('Completed Credits updated successfully.');
  }
}

function openInstructionsVideo() {
  let videoUrl = 'https://watch.screencastify.com/v/MuRWm5sSjy8R3bMQZPlI';
  let htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + videoUrl + '");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Tutorial Video');
}

function updateCompletedCredits() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let autumnSpringDatasheet = spreadsheet.getSheetByName("Fall/Spring CR-CA Data");
  let completedCreditsSheet = spreadsheet.getSheetByName("Completed Credits");
  let autumnSpringDataRange = autumnSpringDatasheet.getRange("A2:T" + autumnSpringDatasheet.getLastRow());
  let existingDataRange = completedCreditsSheet.getRange("A2:M" + completedCreditsSheet.getLastRow());
  let existingData = existingDataRange.getValues();
  let newData = [];
  let counselorNotification = [];

  let autumnData = autumnSpringDataRange.getValues();

  for (let i = 0; i < autumnData.length; i++) {
    let row = autumnData[i];
    let isChecked = row[0]; // Checkbox in column A
    let studentName = row[5].toString().toLowerCase(); // Student Name in column F
    let studentID = row[6]; // Student ID in column G
    let courseName = row[7].toString().toLowerCase(); // Course Name in column H
    let courseNo = row[9]; // Course No in column J
    let courseDateStart = row[10]; // Course Date Start in column K
    let courseEndStart = row[11]; // CourseEndDate Start in column L
    let courseGrade = row[12]; // CourseGrade in column M
    let timeoncourse = row[16]; // timeoncourse in column Q
    let url = row[17]; // LOC link in column R

    if (!isChecked) {
      continue; // Skip if not checked
    }

    let isDuplicate = false;

    for (let j = 0; j < existingData.length; j++) {
      let existingRow = existingData[j];
      let existingStudentID = existingRow[2];
      let existingCourseName = existingRow[3].toString().toLowerCase();

      if (
        existingStudentID === studentID &&
        existingCourseName === courseName
      ) {
        isDuplicate = true;
        break;
      }
    }

    if (isDuplicate) {
      continue; // Skip if duplicate
    }

    let newRow = [
      row[5].toString().toUpperCase(), // Student Name
      studentID, // Student ID
      courseName.toUpperCase(), // Course Name
      parseInt(courseNo,10), // Course ID
      courseDateStart, // Course Date Start
      row[11], // Course Date Credit Earned
      row[12] !== "" ? parseInt(row[12]) : "", // Course Grade Average
      row[15], // Teacher of Record
      row[16], // Hours on course if CA-MT completion
      url, // LOC link
      "", // LS
      "" // NOTES
    ];

    newData.push(newRow);
  }

  newData.sort(function(a, b) {
    let studentA = String(a[1]).toLowerCase();
    let studentB = String(b[1]).toLowerCase();
    return studentA.localeCompare(studentB);
  });

  let insertIndex = 2;

  for (let m = 0; m < newData.length; m++) {
    let studentID = String(newData[m][1]).toLowerCase();
    let nextStudentID = (m + 1 < newData.length) ? String(newData[m + 1][1]).toLowerCase() : "";

    if (studentID.localeCompare(nextStudentID) < 0) {
      insertIndex++;
    }

    completedCreditsSheet.insertRowBefore(insertIndex);
    completedCreditsSheet.getRange(insertIndex, 2, 1, newData[m].length).setValues([newData[m]]).setBackground(null).setBorder(true, true, true, true, true, true);
    insertIndex++;
  }

  completedCreditsSheet.getRange("A2:M" + completedCreditsSheet.getLastRow()).sort({ column: 2, ascending: true });

  // Renumber rows in column A starting from 1
  let lastRow = completedCreditsSheet.getLastRow();
  let rowNumbers = completedCreditsSheet.getRange("A2:A" + lastRow).getValues();
  let updatedRowNumbers = rowNumbers.map(function (value, index) {
    return [index + 1];
  });
  completedCreditsSheet.getRange("A2:A" + lastRow).setValues(updatedRowNumbers);

  sendCounselorNotification()

}

function sendCounselorNotification(student, course) {
  // Map of counselor names to their corresponding emails
  // let counselorEmails = {
  //   '(A-C) Appleby': 'janelle.appleby@nisd.net',
  //   '(D-Ha) Hewgley': 'shanna.hewgley@nisd.net',
  //   '(He-Mi) Ramos': 'elizabeth.ramos@nisd.net',
  //   '(Mo-R) Clarke': 'darrell.clarke@nisd.net',
  //   '(S-Z) Pearson': 'samantha.pearson@nisd.net',
  //   '(ASTA-All Ag Students) Zablocki': 'pamela.zablocki@nisd.net',
  //   '(Head Counselor) Matta': 'orlando.matta@nisd.net'
  // };

  // Map of counselor names to Alvaro's email. For testing purposes.
  let counselorEmails = {
    '(A-C) Appleby': 'alvaro.gomez@nisd.net',
    '(D-Ha) Hewgley': 'alvaro.gomez@nisd.net',
    '(He-Mi) Ramos': 'alvaro.gomez@nisd.net',
    '(Mo-R) Clarke': 'alvaro.gomez@nisd.net',
    '(S-Z) Pearson': 'alvaro.gomez@nisd.net',
    '(ASTA-All Ag Students) Zablocki': 'alvaro.gomez@nisd.net',
    '(Head Counselor) Matta': 'alvaro.gomez@nisd.net'
  };

  // Email
  // Good afternoon,
  // We are happy to report <<Noah Garza>> completed <<Economics for CR>>! 
  // Thank you,
  // Ms. Guajardo and Ms. Bleier
}