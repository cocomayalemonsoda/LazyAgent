var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
var data = sheet.getDataRange().getValues();
var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3");
var commonTsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Common TS");
var resourcesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resources");
var dataResources = resourcesSheet.getDataRange().getValues();
var dataCommonTs = commonTsSheet.getDataRange().getValues();
var dataSheet3 = sheet3.getDataRange().getValues();
var textSheet3 = "";
var escalationLeadNames = [];
var supportNames = [];
var dicNames = [];
var teamLeadNames = [];
var outputText = "";
var bashScript_1 = ""; //Bash script#1
var bashScript_2 = ""; //Bash script#2
//var bashScript_3 = ""; //Bash script#3
var commonTS = ""; //Common TS
var resourceText = "";
var resourceText2 = "";

function writeToC3() {

  sheet.getRange("C3").setValue(textSheet3);
}

for (var i = 0; i < dataSheet3.length; i++) {
    var name = dataSheet3[i][0];
    var escalationLead = dataSheet3[i][1];
    var support = dataSheet3[i][2];
    var dic = dataSheet3[i][3];
    var teamLead = dataSheet3[i][4];

    // Check if the current name has "Yes" in escalation lead
    if (escalationLead === "Yes") {
        escalationLeadNames.push(`workingWithList = "${name}"`);
    } 

    // Check if the current name has "Yes" in support
    if (support === "Yes") {
        supportNames.push(`workingWithList = "${name}"`);
    }

    if (dic === "Yes") {
      dicNames.push(`workingWithList = "${name}"`);
    }

    if (teamLead === "Yes"){
      teamLeadNames.push(`workingWithList = "${name}"`);
    }
}

// Construct the outputText based on the accumulated names
if (escalationLeadNames.length > 0) {
    outputText += "If (";
    outputText += escalationLeadNames.join(" or ");
    outputText += ") {\n";
    outputText += "    WorkingWith := \"Escalation Lead\"\n";
    outputText += "} ";
}

if (supportNames.length > 0) {
    if (outputText !== "") {
        outputText += "else ";
    }
    outputText += "if (";
    outputText += supportNames.join(" or ");
    outputText += ") {\n";
    outputText += "    WorkingWith := \"Support\"\n";
    outputText += "} ";
}

if (dicNames.length > 0) {
    if (outputText !== "") {
      outputText += "else ";
    }
    outputText += "if (";
    outputText += dicNames.join(" or ");
    outputText += ") {\n";
    outputText += "    WorkingWith := \"DIC\"\n";
    outputText += "} ";

}

if (teamLeadNames.length > 0) {
    if (outputText !== "") {
      outputText += "else ";
    }
    outputText += "if (";
    outputText += teamLeadNames.join(" or ");
    outputText += ") {\n";
    outputText += "    WorkingWith := \"Team Leader\"\n";
    outputText += "} ";

}

if (outputText !== "") {
    outputText += "else {\n";
    outputText += "    ;do nothing\n";
    outputText += "}\n";
}

textSheet3 += `#NoEnv\n`;
textSheet3 += `; #Warn\n`;
textSheet3 += `SendMode Input\n`;
textSheet3 += `SetWorkingDir %A_ScriptDir%\n\n`;
textSheet3 += `; Variables\n`;
textSheet3 += `UserSelection := ""\n`;
textSheet3 += `UserName := ""\n`;

textSheet3 += `Gui, Add, Text, x20 y20 w100 h20,Primary Support?\n`;
textSheet3 += `Gui, Add, DropDownList, x130 y20 w150 h180 vUserSelection, `;
for (var i = 1; i < dataSheet3.length; i++) {
  var item = dataSheet3[i][0];
  if (item.trim() !== "") {
  textSheet3 += "|" + item;
  }
}
textSheet3 = textSheet3.slice(0, -1); // Remove the trailing dot from the last item
textSheet3 += `|\n`;
textSheet3 += `Gui, Add, Button, x20 y60 w80 h30 gSubmitSelectionButton, Submit\n`;
textSheet3 += `Gui, Show\n`;
textSheet3 += `return\n\n`;

textSheet3 += `SubmitSelectionButton:\n`;
textSheet3 += `    Gui, Submit\n`;
textSheet3 += `    Gui, Destroy\n`;
textSheet3 += `    GoTo SubmitName\n\n`;

textSheet3 += `SubmitName:\n`;
textSheet3 += `    Gui, Add, Text, x20 y20 w100 h20, Enter your name:\n`;
textSheet3 += `    Gui, Add, Edit, x130 y20 w150 h20 vUserName, Juan D.C.\n`;
textSheet3 += `    Gui, Add, Button, x20 y60 w80 h30 gSubmitNameButton, Submit\n`;
textSheet3 += `    Gui, Show\n`;
textSheet3 += `    return\n\n`;

textSheet3 += `SubmitNameButton:\n`;
textSheet3 += `    Gui, Submit\n`;
textSheet3 += `    UserName := UserName\n`;
textSheet3 += `    MultilineMessage =\n`;
textSheet3 += `    (\n`;
textSheet3 += `Hi %UserName%, your Support for the day is %UserSelection%\n`;
textSheet3 += `Please fill out the template completely and accurately.\n`;
textSheet3 += `\n`;
textSheet3 += `Lazy Agent was developed by Rainville Tobias.\n`;
textSheet3 += `Special Thanks to Ms. Eiz, Kyle, Jeth, Jarry, Sheera, JV, Roechile\n`;
textSheet3 += `Dasvidanya!\n`;
textSheet3 += `)\n`;
textSheet3 += `MsgBox, % MultilineMessage\n`;
textSheet3 += `Gui, Destroy\n`;
textSheet3 += `;Variables\n`;
textSheet3 += `brandList := "" ;brand\n`;
textSheet3 += `registerTypeList := "" ;register\n`;
textSheet3 += `commanderTypeList := "" ;commander\n`;
textSheet3 += `contractTypeList := "" ;contract type\n`;
textSheet3 += `applicationList := "" ;application\n`;
textSheet3 += `supportForTheDayList := UserSelection ;master\n`;
textSheet3 += `workingWithList := "" ;working with\n`;
textSheet3 += `commonTsList := "" ;common ts\n`;
textSheet3 += `resourceList := "" ; resource\n`;
textSheet3 += `HardwareType := ""\n`;
textSheet3 += `hardwareCount := ""\n\n`;

textSheet3 += `Gui, Add, Text, x10 y13 , Contact Person:\n`;
textSheet3 += `Gui, Add, Text, x10 y38 , Callback:\n`;
textSheet3 += `Gui, Add, Text, x10 y63 , Service ID:\n`;
textSheet3 += `Gui, Add, Text, x10 y88 , Store Name & Brand:\n`;
textSheet3 += `Gui, Add, Text, x10 y113 , Hardware Type:\n`;
textSheet3 += `Gui, Add, Text, x10 y138 , Contract Type:\n`;
textSheet3 += `Gui, Add, Text, x10 y163 , App & Version:\n`;
textSheet3 += `Gui, Add, Text, x10 y188 , Issue:\n`;
textSheet3 += `Gui, Add, Text, x10 y208 , Steps to Recreate Issue: (Steps done by Site)\n`;
textSheet3 += `Gui, Add, Text, x10 y267 , Steps to Recreate Issue: (Answers to Probing Questions)\n`;
textSheet3 += `Gui, Add, Text, x10 y338 , KB to use:\n`;
textSheet3 += `Gui, Add, Text, x10 y363 , When did it last work: \n`;
textSheet3 += `Gui, Add, Text, x10 y388 , What Happened/Changed: \n`;
textSheet3 += `Gui, Add, Text, x10 y410 , Insert Sub-Template:\n`;
textSheet3 += `Gui, Add, Text, x10 y473 , Troubleshooting:   \n`;
textSheet3 += `Gui, Add, Text, x10 y605 , Results: \n`;
textSheet3 += `Gui, Add, Text, x260 y8 , Main Support:\n`;
textSheet3 += `Gui, Add, Text, x260 y45 , Working With:\n`;
textSheet3 += `Gui, Add, Text, x403 y8 , Templates\n`;
textSheet3 += `Gui, Add, Text, x400 y90 , Common TS:\n\n`;

textSheet3 += `Gui, Add, Edit, vcontactPerson w100 h20 x90 y10\n`;
textSheet3 += `Gui, Add, Edit, voutputField x10 y655 w475 h20 ; Full Template\n`;
textSheet3 += `Gui, Add, Edit, vcbNumber w100 h20 x90 y35 ; Callback input field\n`;
textSheet3 += `Gui, Add, Edit, vserviceId w100 h20 x90 y60  ; Service ID input field\n`;
textSheet3 += `Gui, Add, Edit, vstoreName w150 h20 x110 y85  ; Store Name & Brand input field\n`;
textSheet3 += `Gui, Add, Edit, vhardwareCount w40 h20 x100 y110 Number ; number of register\n`;
textSheet3 += `Gui, Add, Edit, vappVersion x200 y160 w80 h20 \n`;
textSheet3 += `Gui, Add, Edit, vissueReported x50 y185 w240 h20\n`;
textSheet3 += `Gui, Add, Edit, vdoneBySite x10 y222 w360 h45\n`;
textSheet3 += `Gui, Add, Edit, vkbArticle x65 y335 w40 h19\n`;
textSheet3 += `Gui, Add, Edit, vlastWork x120 y363 w250 h20\n`;
textSheet3 += `Gui, Add, Edit, vwhatChanged x150 y388 w220 h20\n`;
textSheet3 += `Gui, Add, Edit, vsubTemplate x150 y413 w220 h60\n`;
textSheet3 += `Gui, Add, Edit, vtroubleShooting x10 y488 w360 h105\n`;
textSheet3 += `Gui, Add, Edit, vanswersToProbing x10 y283 w360 h45\n`;
textSheet3 += `Gui, Add, Edit, vtsResults x50 y600 w320 h20\n`;
textSheet3 += `Gui, Add, Edit, vbaseVersion x290 y160 w80 h20\n`;
textSheet3 += `Gui, Add, Edit, vextraNotes x380 y488 w100 h105, token / otp / smsuser here\n`;
textSheet3 += `Gui, Add, Edit, vCustomerID x380 y160 w100 h20, Customer ID\n\n`;

// DROPDOWNS
textSheet3 += `Gui, Add, DropDownList, x270 y85 w100 h200 vbrandList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][6];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }

}
textSheet3 += `\n`;
textSheet3 += `brandList := brandList\n\n`;

textSheet3 += `Gui, Add, DropDownList, x150 y110 w80 h100 vregisterTypeList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][7];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `registerTypeList := registerTypeList\n`;


textSheet3 += `Gui, Add, DropDownList, x240 y110 w60 h100 vcommanderTypeList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][8];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `commanderTypeList := commanderTypeList\n\n`;

textSheet3 += `Gui, Add, DropDownList, x100 y135 w150 h100 vcontractTypeList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][9];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `contractTypeList := contractTypeList\n\n`;

textSheet3 += `    Gui, Add, DropDownList, x100 y160 w90 h100 vapplicationList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][10];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `applicationList := applicationList\n\n`;

textSheet3 += `    Gui, Add, DropDownList, x220 y23 w150 h180 vsupportForTheDayList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][11];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `    supportForTheDayList := UserSelection\n\n`;

textSheet3 += `    Gui, Add, DropDownList, x220 y59 w150 h200 vworkingWithList, `;
for (var i=1; i < dataSheet3.length; i++){
  var item = dataSheet3[i][12];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `    workingWithList := workingWithList\n\n`;

textSheet3 += `Gui, Add, DropDownList, x380 y105 w100 h200 vcommonTsList, `;
for (var i=1; i < dataCommonTs.length; i++){
  var item = dataCommonTs[i][0];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `    commonTsList := commonTsList\n\n`;

textSheet3 += `Gui, Add, DropDownList, x110 y335 w195 h200 vresourceList, Choose a resource to open..`;
for (var i=1; i < dataResources.length; i++){
  var item = dataResources[i][0];
  if (item.trim() !== "") {
  textSheet3 += "|" + item ;
  }
}
textSheet3 += `\n`;
textSheet3 += `resourceList = resourceList\n\n`;

textSheet3 += `Gui, Add, Button, x380 y185 w100 h20 gShellSubject, Shell Esca Subject\n`
textSheet3 += `Gui, Add, Button, x300 y185 w70 h20 gSflTitle, SFL Subject\n`
textSheet3 += `Gui, Add, Button, x10 y425 w60 h20 gFuelSub, Fuel\n`
textSheet3 += `Gui, Add, Button, x75 y425 w60 h20 gPinSub, Pinpad\n`
textSheet3 += `Gui, Add, Button, x10 y630 gCopyText1, Support Template\n`
textSheet3 += `Gui, Add, Button, x120 y630 gCopyText2, Sales Force Template\n`
textSheet3 += `Gui, Add, Button, x250 y630 gCopyClipboard, Copy to Clipboard\n`
textSheet3 += `Gui, Add, Button, x425 y630 w60 h20 gResetFields, RESET\n`
textSheet3 += `Gui, Add, Button, x380 y25 w100 h20 gDictemplate, DIC\n`
textSheet3 += `Gui, Add, Button, x380 y45 w100 h20 gShellTemplate, Shell\n`
textSheet3 += `Gui, Add, Button, x380 y65 w100 h20 gDispatchTemplate, Dispatch\n`
textSheet3 += `Gui, Add, Button, x380 y125 w100 h20 gGenerate, Generate\n`
textSheet3 += `Gui, Add, Button, x310 y335 w60 h20 gResource, Open\n`

textSheet3 += `screenWidth := A_ScreenWidth\n`
textSheet3 += `screenHeight := A_ScreenHeight\n`
textSheet3 += `centerX := screenWidth / 2-250\n`
textSheet3 += `centerY := screenHeight / 2-300\n`
textSheet3 += `xPosition := centerX\n`
textSheet3 += `yPosition := centerY\n`

textSheet3 += `Gui, Show, x%xPosition% y%yPosition% w495 h700 , Lazy Agent\n\n`

textSheet3 += `Generate:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControlGet, ExistingTS, , troubleShooting\n`;
var tsName1 = dataCommonTs[1][0];
var tsSteps1 = dataCommonTs[1][1];
var kbArticleUsed1 = dataCommonTs[1][2];
if (tsName1.trim() !== "" && tsSteps1.trim() !== "") {
  bashScript_1 += `    If (commonTsList = "${tsName1}") {\n`;
  bashScript_1 += `    GuiControl,, outputField, ${tsSteps1}\n`;
  bashScript_1 += `    GuiControl,, kbArticle, ${kbArticleUsed1}\n`;
  bashScript_1 += `     return\n\n`
  bashScript_1 += `    }\n`;
}

var tsName2 = dataCommonTs[2][0];
var tsSteps2 = dataCommonTs[2][1];
var kbArticleUsed2 = dataCommonTs[2][2];
if (tsName2.trim() !== "" && tsSteps2.trim() !== "") {
  bashScript_2 += `    else If (commonTsList = "${tsName2}") {\n`;
  bashScript_2 += `    if (applicationList == "Buypass"){\n`
  bashScript_2 += `    GuiControl,, outputField, ${tsSteps2}\n`;
  bashScript_2 += `    GuiControl,, kbArticle, ${kbArticleUsed2}\n`;
  bashScript_2 += `     return\n`
  bashScript_2 += `    } `
  bashScript_2 += `    else {\n`
  bashScript_2 += `    MsgBox, Warning: This will only work for Buypass Application!\n`
  bashScript_2 += `     return\n\n`
  bashScript_2 += `    }\n`
  bashScript_2 += `    }\n`;
}


/*
var tsName3 = dataCommonTs[3][0];
var tsSteps3 = dataCommonTs[3][1];
var kbArticleUsed3 = dataCommonTs[3][2];
if (tsName3.trim() !== "" && tsSteps3.trim() !== "") {
  bashScript_3 += `    else If (commonTsList = "${tsName3}") {\n`;
  bashScript_3 += `    if (applicationList == "Buypass"){\n`
  bashScript_3 += `    GuiControl,, outputField, ${tsSteps3}\n`;
  bashScript_3 += `    GuiControl,, kbArticle, ${kbArticleUsed3}\n`;
  bashScript_3 += `     return\n`
  bashScript_3 += `    }\n`;
}
*/



for (var i = 3; i < dataCommonTs.length; i++) {
  var tsNameNext = dataCommonTs[i][0];
  var tsStepsNext = dataCommonTs[i][1];
  var kbArticleUsedNext = dataCommonTs[i][2];
  
  if (tsNameNext.trim() !== "" && tsStepsNext.trim() !== "") {
    commonTS += `    else If (commonTsList = "${tsNameNext}") {\n`;
    commonTS += `        newTS := ExistingTS . "${tsStepsNext}"\n`;
    commonTS += `    GuiControl,, troubleShooting, %newTS%\n`;
    commonTS += `    GuiControl,, kbArticle, ${kbArticleUsedNext}\n`;
    commonTS += `     return\n\n`;
    commonTS += `    }\n`;
  }
}
textSheet3 += bashScript_1;
textSheet3 += bashScript_2;
//textSheet3 += bashScript_3;
textSheet3 += commonTS;


textSheet3 += `Resource:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
var resourceName1 = dataResources[1][0];
var resourceLink1 = dataResources[1][1];

if (resourceName1.trim() !== "" && resourceLink1.trim() !== "") {
    resourceText += `if (resourceList = "${resourceName1}"){\n`;
    resourceText += `URL := "${resourceLink1}"\n`;
    resourceText += `WinActivate, ahk_exe chrome.exe\n`;
    resourceText += `WinWaitActive, ahk_exe chrome.exe\n`;
    resourceText += `Send, ^t\n`;
    resourceText += `Sleep, 500\n`;
    resourceText += `Send, %URL%{Enter}\n`;
    resourceText += `return\n`;
    resourceText += `}\n`;
	
}

for (var i = 2; i < dataResources.length; i++) {
  var resourceNameNext = dataResources[i][0];
  var resourceLinkNext = dataResources[i][1];
  
  if (resourceNameNext.trim() !== "" && resourceLinkNext.trim() !== "") {
    resourceText2 += `else if (resourceList = "${resourceNameNext}"){\n`;
    resourceText2 += `URL := "${resourceLinkNext}"\n`;
    resourceText2 += `WinActivate, ahk_exe chrome.exe\n`;
    resourceText2 += `WinWaitActive, ahk_exe chrome.exe\n`;
    resourceText2 += `Send, ^t\n`;
    resourceText2 += `Sleep, 500\n`;
    resourceText2 += `Send, %URL%{Enter}\n`;
    resourceText2 += `return\n`;
    resourceText2 += `} `;
  }
}
textSheet3 += resourceText;
textSheet3 += resourceText2;
textSheet3 += `else {\n`;
textSheet3 += `;do nothing\n`;
textSheet3 +=  `return\n`;
textSheet3 += `}\n`;

// FUNCTIONS
textSheet3 += `FuelSub:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, subTemplate, Number of FPs & DCRs:\`r\`nAffected FPs & DCRs:\`r\`nDispenser Brand: \`r\`nDCR Brand: \`r\`nDCR Driver: (set to serial or ethernet?)\`r\`nInside/Outside/Both: \nreturn\n\n`;

textSheet3 += `PinSub:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, outputField, if [ $(awk -F'=' '/Suite_Name/{print $2}' /home/\${PLATFORM_APPUSER}/config/suiteinfo.dat) = "shell" ]; then\`r\`necho "NOT AVAILABLE FOR SHELL SITES!"\`r\`nelse\`r\`nfind /home/pmc/data/pop -type f -name "[0-9][0-9][0-9].properties" | sort | while IFS= read -r i; do\`r\`nPinpad=\$(basename \$i | cut -c 1-3)\`r\`nModel=\$(awk -F'=' '/POP_Model=/{print \$2}' \$i)\`r\`nApp_Version=\$(awk -F'=' '/App_Version=/{print \$2}' \$i)\`r\`nOS_Version=\$(awk -F'=' '/OS_Version=/{print \$2}' \$i)\`r\`nSerialNumber=\$(awk -F'=' '/Pinpad_SerialNumber=/{print \$2}' \$i)\`r\`necho "Pinpad: \$Pinpad\`r\`nModel: \$Model\`r\`nApp & Version: \$App_Version\`r\`nOS Version: \$OS_Version\`r\`nSerial Number: \$SerialNumber\`r\`n"\`r\`ndone\`r\`nfi\nreturn\n\n`;


textSheet3 += `Dictemplate:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, outputField, Sales Force Case # : \`r\`nType of Connection: BASTION / SUNOCO / TESORO\`r\`nIP ADDRESS / TUNNEL IP : \`r\`nTOKEN : \`r\`nBase version CMDR : \`r\`nPiNpad IP ADD : \`r\`nIssue : PiNpad 001 stuck on Blue Screen / Apply New Patch 4.07.10\`r\`nPiNpad stats : Waiting for Download (Netloader)\`r\`nCallback # : %cbNumber%\`r\`nAlternate CB #: \`r\`nApproved by : ESCA LEAD %workingWithList%\`r\`nTested By : %UserName%\nreturn\n\n`;

textSheet3 += `ShellTemplate:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, outputField, VFI HELPDESK TICKET #: \`r\`nAGENT NAME: %UserName% \`r\`nSITE NAME: %storeName% \`r\`nSITE PHONE: %cbNumber%\`r\`nPROBLEM REPORTED: \`r\`nSHELL MERCHANT ID: \`r\`nTRANSACTION DATE AND TIME: \`r\`nPOS TYPE/VERSION: %hardwareCount% %registerTypeList% %commanderTypeList% / %applicationList% %appVersion%\`r\`nTYPE OF TRANSACTION (INDOOR, OUTDOOR, PREPAY, POSTPAY): \`r\`nCURRENT SITE STATUS (HARDDOWN, OPERATIONAL): \`r\`nSITE REQUIREMENTS (SERVICER NEEDED, N/A): \`r\`nHARDWARE FAILURE (LIST WHICH HARDWARE NEEDS REPLACEMENT IF ANY - OTHERWISE N/A): \nreturn\n\n`;

textSheet3 += `DispatchTemplate:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, outputField, DISPATCH TYPE: (Emergency / Routine, See SOP for Definition)\`r\`nISSUE DESCRIPTION or SYMPTOMS:\`r\`nVERIFIED SOFTWARE VERSION OF SITE CONTROLLER AND/OR POS: %hardwareCount% %registerTypeList% %commanderTypeList% / %applicationList% %appVersion% \`r\`nPOSSIBLE PARTS NEEDED AFTER TROUBLESHOOTING and SERVICES REQUESTED: (Must Have POS Model)\`r\`nSITE HOURS:\`r\`nPoint of Contact:\`r\`nCallback#: %cbNumber%\`r\`nAlternate#:\`r\`nHELPDESK TICKET #:\`r\`nFLASH: (Copy Flash Here only if relevant to dispatch)\`r\`nAPPROVED BY: Escalation Lead %workingWithList%\nreturn\n\n`;

textSheet3 += `SflTitle:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    GuiControl,, outputField, %applicationList% %appVersion% / Base %baseVersion% / %issueReported%\nreturn\n\n`;

textSheet3 += `ShellSubject:\n`;
textSheet3 += `    Gui, Submit, NoHide\n`;
textSheet3 += `    if (applicationList == "Shell"){\n`;
textSheet3 += `        GuiControl,, outputField, SHELL %appVersion%:: VFI: %CustomerID%:: %issueReported%\n`;
textSheet3 += `    } else {\n`;
textSheet3 += `        MsgBox, Warning: This will only work for Shell Application!\n`;
textSheet3 += `    }\n`;
textSheet3 += `return\n\n`;

textSheet3 += `CopyText1:\n`;
textSheet3 += `    Gui, Submit, NoHide ; Get the values fron Inputs\n`;
textSheet3 += `    If (registerTypeList=="Ruby2" && commanderTypeList =="CI"){\n`;
textSheet3 += `        HardwareType = %hardwareCount% Ruby2 / CI\n`;
textSheet3 += `    } else if (registerTypeList=="Topaz" && commanderTypeList =="CI") {\n`;
textSheet3 += `        MsgBox, Warning: Topaz could not have a CI :3\n`;
textSheet3 += `    } else if (registerTypeList=="C18" && commanderTypeList =="CI") {\n`;
textSheet3 += `        MsgBox, Warning: C18 could not have a CI :3\n`;
textSheet3 += `    } else {\n`;
textSheet3 += `        HardwareType = %hardwareCount% %registerTypeList%, 1  %commanderTypeList%\n`;
textSheet3 += `    }\n\n`;
textSheet3 += `    GuiControl,, outputField, SUPPORT STATS: %supportForTheDayList% | PRIMARY SUPPORT \`r\`nService ID: %serviceId%\`r\`nStore Name & Brand: %storeName% / %brandList%\`r\`nHardware Type: %HardwareType%\`r\`nContract Type: %contractTypeList%\`r\`nApp & Version: %applicationList% %appVersion% / Base %baseVersion%\`r\`nIssue: %issueReported%\`r\`nSteps to Recreate Issue:\`n%doneBySite%\`n%answersToProbing%\`r\`nWhat I intend to do with the information provided or with the KB Article: %kbArticle%\`r\`n\`r\`nWhen did it last work: %lastWork%\`r\`nWhat Happened/Changed: %whatChanged%\`r\`nInsert Sub-Template:\`n%subTemplate%\`r\`nTroubleshooting:\`n%troubleShooting%\`n\`nIdentify what you're currently doing with the site, or ask a specific question:\nreturn\n\n`;

textSheet3 += `CopyText2:\n`;
textSheet3 += `    Gui, Submit, NoHide ; Get the values fron Inputs\n`;
textSheet3 += `      if (lastWork = "") {\n`
textSheet3 += `          DidItEverWorked := "No"\n`
textSheet3 += `      } else {\n`
textSheet3 += `          DidItEverWorked := "Yes"\n`
textSheet3 += `      }\n`
textSheet3 += `      if (doneBySite = "") {\n`
textSheet3 += `          StepsTaken := "- no steps taken"\n`
textSheet3 += `      } else {\n`
textSheet3 += `          StepsTaken := doneBySite\n`
textSheet3 += `      }\n`
textSheet3 += outputText;
textSheet3 += `    if (workingWithList == "") {\n`;
textSheet3 += `        workedWith := ""\n`;
textSheet3 += `        WorkingWith :=\n\n`;
textSheet3 += `    } else {\n`;
textSheet3 += `        workedWith := "Worked With: "\n\n`;
textSheet3 += `    }\n\n`;
textSheet3 += `    GuiControl,, outputField, Issue Reported: %issueReported%\`nDid it ever work: %DidItEverWorked%\`r\`nActivity at Time of Occurrence:\`n- issue started %lastWork%\`n%whatChanged%\`r\`nSteps Taken Before Call: \`n%StepsTaken%\`r\`nKB/Resource(s) utilized: KB %kbArticle%\`r\`nAction/Capture Information: \`n%answersToProbing%\`n%subTemplate%\`r\`nContact Person: %contactPerson%\`r\`nCallback: %cbNumber%\`n\`r\`nTroubleshooting: \`n%troubleShooting%\`n[**]\`nResults: %tsResults% \`n\`n%workedWith%%WorkingWith% %workingWithList%\nreturn\n\n`;


textSheet3 += `CopyClipboard:\n`;
textSheet3 += `    Gui, Submit, NoHide ; Get the value from the second input field\n`;
textSheet3 += `    Clipboard := outputField ; Copy the content of the second input field to the clipboard\nreturn\n\n`;

textSheet3 += `ResetFields:\n`;
textSheet3 += `    ShowResetConfirmation()\nreturn\n\n`;

textSheet3 += `ShowResetConfirmation() {\n`;
textSheet3 += `    MsgBox, 4, Reset, There's no CTRL+Z on this one. Are you sure?\n`;
textSheet3 += `    IfMsgBox, Yes\n`;
textSheet3 += `    {\n`;
textSheet3 += `        WinGetPos, GuiX, GuiY, , , ahk_id %GuiHwnd%\n`;
textSheet3 += `        ToolTip, Customer ID, %GuiX%+480, %GuiY%+230\n`;
textSheet3 += `        SetTimer, RemoveToolTip, 3000\n`;
textSheet3 += `        Gui, Submit, NoHide\n`;
textSheet3 += `        ; Clear vars\n`;
textSheet3 += `        hardwareCount := ""\n`;
textSheet3 += `        registerTypeList := ""\n`;
textSheet3 += `        commanderTypeList := ""\n`;
textSheet3 += `        contractTypeList := ""\n`;
textSheet3 += `        applicationList := ""\n`;
textSheet3 += `        supportForTheDayList := ""\n`;
textSheet3 += `        workingWithList := ""\n`;
textSheet3 += `        commonTsList := ""\n`;
textSheet3 += `        resourceList := ""\n\n`;
textSheet3 += `        ; Clear vals\n`;
textSheet3 += `        GuiControl,, hardwareCount, %hardwareCount%\n`;
textSheet3 += `        GuiControl Choose, brandList, 0\n`;
textSheet3 += `        GuiControl Choose, registerTypeList, 0\n`;
textSheet3 += `        GuiControl Choose, commanderTypeList, 0\n`;
textSheet3 += `        GuiControl Choose, contractTypeList, 0\n`;
textSheet3 += `        GuiControl Choose, applicationList, 0\n`;
textSheet3 += `        GuiControl Choose, workingWithList, 0\n`;
textSheet3 += `        GuiControl Choose, commonTsList, 0\n`;
textSheet3 += `        GuiControl Choose, resourceList, 1\n\n`;
textSheet3 += `        ; Clear fields\n`;
textSheet3 += `        GuiControl,, contactPerson,\n`;
textSheet3 += `        GuiControl,, cbNumber,\n`;
textSheet3 += `        GuiControl,, serviceId,\n`;
textSheet3 += `        GuiControl,, storeName,\n`;
textSheet3 += `        GuiControl,, appVersion,\n`;
textSheet3 += `        GuiControl,, issueReported,\n`;
textSheet3 += `        GuiControl,, doneBySite,\n`;
textSheet3 += `        GuiControl,, kbArticle,\n`;
textSheet3 += `        GuiControl,, lastWork,\n`;
textSheet3 += `        GuiControl,, whatChanged,\n`;
textSheet3 += `        GuiControl,, subTemplate,\n`;
textSheet3 += `        GuiControl,, troubleShooting,\n`;
textSheet3 += `        GuiControl,, answersToProbing,\n`;
textSheet3 += `        GuiControl,, tsResults,\n`;
textSheet3 += `        GuiControl,, baseVersion,\n`;
textSheet3 += `        GuiControl,, extraNotes,\n`;
textSheet3 += `        GuiControl,, CustomerID,\n\n`;
textSheet3 += `        ; Clear output field\n`;
textSheet3 += `        GuiControl,, outputField,\n`;
textSheet3 += `    }\n`;
textSheet3 += `}\n\n`;

textSheet3 += `RemoveToolTip:\n`;
textSheet3 += `    ToolTip\n`;
textSheet3 += `    SetTimer, RemoveToolTip, Off\n`;
textSheet3 += `    return\n\n`;
textSheet3 += `GuiClose2:\n`;
textSheet3 += `    ExitApp\nreturn\n\n`;

textSheet3 += `    return\n\n`;

textSheet3 += `^1::SendInput, Cashier\n`;
textSheet3 += `^2::SendInput, Manager\n`;
textSheet3 += `^3::SendInput, Owner\n`;
textSheet3 += `^0::SendInput, df{Enter}uptime{Enter}suiteinfo{Enter}\n`;
textSheet3 += `^!+o::SendInput, Miscellaneous\n`;

Logger.log(textSheet3);