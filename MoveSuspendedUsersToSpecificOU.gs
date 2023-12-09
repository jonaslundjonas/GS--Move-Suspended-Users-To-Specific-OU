// Created by Jonas Lund, 2023. This script is a part of scripts aimed for handling offboarding processes.
// It moves any suspended user in any OU to a OU named Suspended.
// You will need to supply the code with a super admin address email account and your customer ID.
// It is recomended to create a time based trigger for the moveSuspendedAccounts function.
// The Script can be runned from the script menu if requiered
// Logs will be added to the sheet

// Requires an OU named Suspended Users
// Enable the following APIs in services to the right in the editor:
// 1. Admin SDK Directory API
// 2. Google Sheets API

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Script Menu')
    .addItem('Move Suspended Accounts', 'moveSuspendedAccounts')
    .addToUi();
}

function moveSuspendedAccounts() {
  const targetOU = "/Suspended Users";
  const suspendedUsers = getSuspendedUsers(targetOU);
  const targetOuPath = getOuPath(targetOU);

  if (targetOuPath) {
    if (suspendedUsers.length > 0) {
      logToSheet("Target OU found, moving suspended users...");
      moveUsersToOu(suspendedUsers, targetOuPath);
    } else {
      logToSheet("No suspended users to move.");
    }
  } else {
    logToSheet("Target OU not found.", true);
  }
}

function getSuspendedUsers(targetOU) {
  const suspendedUsers = [];

  // Add your customer ID here
  const customerId = 'xxxxxxxxx'; // Replace with your customer ID
  const response = AdminDirectory.Users.list({
    customer: customerId,
    maxResults: 500,
    query: 'isSuspended=true'
  });
  const users = response.users;

  if (users) {
    for (const user of users) {
      if (user.orgUnitPath !== targetOU) {
        suspendedUsers.push(user);
      }
    }
  }

  logToSheet("Amount of users to move: " + suspendedUsers.length);
  return suspendedUsers;
}

function getOuPath(targetOU) {
  // Add your Customer ID here
  const customerId = 'XXXXXX'; // Replace with your customer ID
  const orgUnits = AdminDirectory.Orgunits.list(customerId).organizationUnits;

  for (const orgUnit of orgUnits) {
    if (orgUnit.orgUnitPath === targetOU) {
      logToSheet("Target OU Path: " + orgUnit.orgUnitPath);
      return orgUnit.orgUnitPath;
    }
  }

  return null;
}

function moveUsersToOu(users, targetOuPath) {
  for (const user of users) {
    try {
      AdminDirectory.Users.update({orgUnitPath: targetOuPath}, user.primaryEmail);
      logToSheet("Moved " + user.primaryEmail + " to " + targetOuPath);
    } catch (error) {
      logToSheet("Error moving " + user.primaryEmail + ": " + error.message, true);
    }
  }
}

function logToSheet(message, isError = false) {
  const sheetName = "Logs";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(["Timestamp", "Type", "Message"]);
  }

  const timestamp = new Date();
  const logType = isError ? "Error" : "Info";
  sheet.appendRow([timestamp, logType, message]);
}
