function exportDataToAHKFile() {
  var ownerEmail = "RainvilleT1@verifone.com"; // Replace with the actual owner's email

  // Get the email of the current user
  var currentUser = Session.getActiveUser().getEmail();

  // Check if the current user is the owner
  if (currentUser !== ownerEmail) {
  
    Browser.msgBox('Chill,', 'pm mo muna ako kung need mo latest version :3', Browser.Buttons.OK);
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();
  var text = "";

  // dito nakalagay yung version
  var currentVersion = sheet.getRange("A1").getValue();

  // based sa column E, calculate next version
  var changelogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2");
  var versions = changelogSheet.getRange("E2:E" + changelogSheet.getLastRow()).getValues(); // Assuming version numbers start from row 2
  var maxVersion = currentVersion;


var changelogVersion = maxVersion; // Initialize changelog version with maxVersion
var customName = "Lazy Agent "; //include space
var folderId = "1VCbLaWnryp33Ldirln3S5jhnR6wf2qSI";
var currentCategory = null;

  // GUI

  for (var i = 2; i < data.length; i++) { // start at row 3
    var category = data[i][0];
    var keyword = data[i][1];
    var hotstring = data[i][2];

    // Skip empty rows
    if (category === "" && keyword === "" && hotstring === "") {
      continue;
    }

    // If the category is not empty, update the current category
    if (category != "") {
      currentCategory = category;
      text += `; Category: ${currentCategory}\n\n`;
    }

    if (keyword === "commandernotonboardedcn") {
      // Insert the special format for these keywords
      text += `:*:${keyword}::\nGui, Add, Text, , Has the site already signed up for Verifone C-site Management?\n`;
      text += `Gui, Add, Button, y30 x50 w100 h30 gYesButton, Yes\n`;
      text += `Gui, Add, Button, y30 x+10 w100 h30 gNoButton, No\n`;
      text += `Gui, Show,, C-site Management\nreturn\n\n`;
      text += `YesButton:\nGui, Submit\nGui, Destroy\nSendInput,\n(\n${specialFormatTextForYes}\n)\nreturn\n\n`;
      text += `NoButton:\nGui, Submit\nGui, Destroy\nSendInput,\n(\n${specialFormatTextForNo}\n)\nreturn\n\n`;

    } else if (keyword === "resetpwcn") {
      text += `:*:${keyword}::\n`
      text += `resetpwUserName := ""\n\n`;
      text += `Gui, Add, Text, x20 y20 w130 h20, Username to reset pw:\n`;
      text += `Gui, Add, Edit, x140 y20 w150 h20 vresetpwUserName, manager\n`;
      text += `Gui, Add, Button, x20 y60 w80 h30 gSubmitpwuserNameButton, Submit\n`;
      text += `Gui, Show\n\n`;
      text += `return\n\n`;
      text += `SubmitpwuserNameButton:\n`;
      text += `Gui, Submit\n`;
      text += `resetpwUserName := resetpwUserName\n`;
      text += `Gui, Destroy\n`;
      text += `SendInput,\n(\n${hotstring}\n)\nreturn\n\n`;
   
    } else if (keyword === "lazyagentgui") {
      //writeToC3.writeToC3();
      text += `${hotstring}`;

    } else {
      // Default format for other keywords
      text += `:*:${keyword}::\nSendInput,\n(\n${hotstring}\n)\nreturn\n\n`;
    }

    // ensure return has no params
    if (i < data.length - 1 && (data[i + 1][0] !== "" || data[i + 1][1] !== "" || data[i + 1][2] !== "")) {
      text += "\n\n"; // Add a blank line if the next row is not empty
    }
  }

  // sa ilalim:
  text += "/************Changelogs***********\n";

  // Include changelog entries from Sheet2
  var changelogData = changelogSheet.getDataRange().getValues();
  for (var j = 1; j < changelogData.length; j++) {
    var date = changelogData[j][0];
    var action = changelogData[j][1];
    var keyword = changelogData[j][2];
    var hotstring = changelogData[j][3];

    // changelog version
    var changelogVersion = changelogData[j][4];

    // date format
    var formattedDate = new Date(date).toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

    text += `${formattedDate}::${action}::${keyword}::${hotstring}::version ${changelogVersion}\n`;
  }

  // ilalim ng changelog
  text += "***********************************/\n";

  // saya ka?
  text += `/*---------------------Author's Note [${currentVersion}]---------------------\n`;
  text += `If you're interested in submitting your entries/requests, don't hesitate to contact me at RainvilleT1@verifone.com\n`;
  text += `dasvidanya!\n`;
  text += `--------------------------------------------------------------------------------*/\n`;

  // Increment the version number
  var versionParts = maxVersion.split('.');
  var majorVersion = parseInt(versionParts[0]);
  var minorVersion = parseInt(versionParts[1]);
  var patchVersion = parseInt(versionParts[2]) + 1;

  if (patchVersion > 9) {
    patchVersion = 0;
    minorVersion++;
  }

  if (minorVersion > 9) {
    minorVersion = 0;
    majorVersion++;
  }

  currentVersion = `${majorVersion.toString().padStart(2, '')}.${minorVersion.toString().padStart(2, '0')}.${patchVersion}`;



  // Update current version
  sheet.getRange("A1").setValue(currentVersion);
    for (var i = 2; i <= changelogSheet.getLastRow(); i++) {
    var versionCell = changelogSheet.getRange("E" + i);

    if (versionCell.isBlank()) {
      // Set the empty version cell to the currentVersion
      versionCell.setValue(currentVersion);
    }
  }

  // Create a file with name, version, and extension
  DriveApp.getFolderById(folderId).createFile(customName + currentVersion + ".ahk", text);
}
