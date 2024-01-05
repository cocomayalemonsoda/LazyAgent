# Lazy Agent Guide - For Google Apps Script
Lazy Agent is a tool crafted for Technical Support Representatives, aiming to streamline the case documentation process. It is written in AutoHotkey syntax, generated through this Google Apps Script Project.

This guide will delve into the sections of the Google Sheet where users can effortlessly edit, add, modify, and create necessary updates. Designed to cater to users unfamiliar with programming languages like JavaScript, Bash, or AutoHotkey scripting, this project ensures accessibility for a broader audience.

For all of the sheets, we'll be using this guide:

**Color Codes**

*GREEN* – User can edit all the rows below except those with RED color.

*YELLOW* – User can edit items but may need to check JavaScript if error occurs.

*RED* – User MUST NOT edit items in cells, otherwise will cause errors to the .ahk file.

**Definition of Terms**

*Hotstrings* – Texts that will be generated upon typing the keywords

*Category* – Used for classification of the hotstrings

*Keyword* – Characters you need to type to generate the corresponding string

*String* – a.k.a Text

*Description* – short description of the specific hotkey/hotstrings that will be generated upon typing

<br>

## Sheet Name: Sheet1
![Lazy Agent Image](Lazy%20Agent%20Images/image005.png)


### How to add items on the list? 
1.	Select a row number then right click > “insert 1 row below” – No need to enter a category if it falls under the same category as the previous item. 
2.	Enter the keyword that you want – entering duplicate keywords will cause errors in the .ahk file.
3.	Enter the String/Text that you want to be generated upon typing the keyword you just created. – If it requires multiple lines, it’s best to just copy it from notepad and paste it in the cell. Add a description if needed.
4.	OPTIONAL: Add changes that you made in Sheet2

### How to generate update? 
1.	Click the Verifone logo on the upper left corner of the sheet.
2.	If it asks for permission, follow the prompt to “Advanced” and “continue to Lazy Agent” – You may need to run(Click the logo) it again after giving the permission
3.	After the version changes adds .1(patch) to the version number, wait for the user interface to show “Finished running script” then the new version should be directly uploaded to the google drive link: https://drive.google.com/drive/folders/13lkO41l5RAJA5-ZvcL1uzwOFM0OvIzK6

## Sheet Name: Sheet2 - Optional Changelogs Sheet
![Lazy Agent Image](Lazy%20Agent%20Images/image007.png)

-	This is where you can input the date, changes and descriptions that you’ve made in the updated versions. This is also to inform other people that have access to the App about the changes in the Lazy agent. NOTE: Please leave the version number, it will be automatically updated every time the code runs.

## Sheet Name: Sheet3
Sheet3 is divided into 2 sections: **Section 1** and **Section 2**.

### Section 1
![Lazy Agent Image](Lazy%20Agent%20Images/image011.png)

### How to add items on the list? 

The user must not add by inserting a new row in between the rows. Instead, they can only add new rows on the bottom of the list:
1.	Add a new name like so:
![Lazy Agent Image](Lazy%20Agent%20Images/image012.png)
 
2.	Highlight the 4 cells above. Click and drag the corner (lower right) to copy to the bottom row.
![Lazy Agent Image](Lazy%20Agent%20Images/image014.png)
![Lazy Agent Image](Lazy%20Agent%20Images/image016.png)

3.	Assign the name to which position they belong. Type “Yes” on the cell you want to assign them to. There can only be one “Yes” in a row so other cells will be grayed out (Automatically) but still be editable.
![Lazy Agent Image](Lazy%20Agent%20Images/image018.png)

### Section 1
![Lazy Agent Image](Lazy%20Agent%20Images/image020.png)

### How to add items on the list? 
-	You can add by simply typing down item that you want to add right below the last item on the list.
-	All the items in the RED column MUST NOT be edited. They are auto generated based on Section 1.
- The order of items that will show up in the dropdown list inside the Lazy Agent .ahk file will depend on the order of items in Sheet3.

 ## Sheet Name: CommonTS
 
 ![Lazy Agent Image](Lazy%20Agent%20Images/image022.png)
 
This is the sheet for which the items in the **Common TS** is located:
 ![Lazy Agent Image](Lazy%20Agent%20Images/image025.png)
 
There are two types of “Troubleshooting Steps” in this section: **Troubleshooting steps** and **Bash script**

### How to add items on the list? 
- *Troubleshooting steps*

1.	At the bottom of the list, add in the “TS Name” – This will appear on the dropdown menu. Keep it short so that the text won’t be cut out.
2.	Before you paste the Troubleshooting steps in the cell, it must be formatted to comply with the Autohotkey syntax by adding \`r`n to the ends of EACH LINE. Essentially making the entire troubleshooting steps to be one line only.

For example

```
Troubleshooting steps

Rebooted the commander via power cable
Confirmed all lights went off
Plugged back in after 5 minutes
Checked Commander status > A9 and System OK
```

```
Format:

-Rebooted the commander via power cable\`r`n-Confirmed all lights went off\`r`n-Plugged back in after 5 minutes\`r`n-Checked Commander status > A9 and System OK\`r`n
```

3.	Paste the formatted version in the troubleshooting steps cell.
4.	You may also write the KB article that was used for the troubleshooting steps.

<br>

- *Bash Scripts*

1.	Insert a new row at the bottom of the last bash script.
2.	Enter the TS name.
3.	Format the bash script to be a single line bash script or add the same `r`n to each line like the format for Troubleshooting steps.
4.	 Go to Extensions > Apps Script
 ![Lazy Agent Image](Lazy%20Agent%20Images/image026.png)

5.	Click writeToC3.gs and hit ctrl+F > search for “//Bash Script#3”
 ![Lazy Agent Image](Lazy%20Agent%20Images/image028.png)

It should lead you to the first commented code for the bash script. Remove the “//” to make it included in the JavaScript code.
NOTE: If you need a new one, just copy and name the variable accordingly. “3” beside the name is an arbitrary number.

6.	Click the Down arrow while still searching for “//Bash Script#3”. – The entire code is still commented out so you need to remove the /* and */ at the start and end of the JavaScript code.
Here’s the code we want to edit:

```
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
```

### Adding more bash scripts

NOTE: If you need to add more bash scripts, copy and paste the same code and name the variables accordingly (Highlighted in yellow). Particular [Row][Column] location in the array (Highlighted in green) must also be edited.

```
var tsNamex+1 = dataCommonTs[n+1][0];
var tsStepsx+1 = dataCommonTs[n+1][1];
var kbArticleUsedx+1 = dataCommonTs[n+1][2];
if (tsNamex+1.trim() !== "" && tsStepsx+1.trim() !== "") {
  bashScript_x+1 += `    else If (commonTsList = "${tsNamex+1}") {\n`;
  bashScript_x+1 += `    if (applicationList == "Buypass"){\n`
  bashScript_x+1 += `    GuiControl,, outputField, ${tsStepsx+1}\n`;
  bashScript_x+1 += `    GuiControl,, kbArticle, ${kbArticleUsedx+1}\n`;
  bashScript_x+1 += `     return\n`
  bashScript_x+1 += `    }\n`;
}

```

```
Where: 
x is the last number on the list of bash scripts
n is the row number of the last item on the list

NOTE: To make it simple, keep x = n
```

7.	Click the Down arrow again while still searching for “//Bash Script#3”. – The code is still commented out, so you need to remove the “//” at the start and end of the JavaScript code.
Here’s the code we want to edit:

```
textSheet3 += bashScript_1;
textSheet3 += bashScript_2;
//textSheet3 += bashScript_3;
textSheet3 += commonTS;
```

Take note that if you add another bash script (for example bashScript_4);, you need to add a new line. For the next ones, you just need to name them accordingly

 ```
textSheet3 += bashScript_1;
textSheet3 += bashScript_2;
textSheet3 += bashScript_3;
textSheet3 += bashScript_4;
//textSheet3 += bashScript_x+1;
textSheet3 += commonTS;
```

8.	Click the save icon and Refresh the Lazy Agent Sheet before generating a new update.
OPTIONAL: You can run the script by running the function inside Code.gs
	The updated file will also go to the google drive with an updated file name and version.

##  ## Sheet Name: Resources
 ![Lazy Agent Image](Lazy%20Agent%20Images/image030.png)
 
These are the ones that will appear here:
 ![Lazy Agent Image](Lazy%20Agent%20Images/image032.png)

All items can be edited but make sure of the links.
Keep the name short to prevent it from being cut out.
Items can be found in the this link:
https://docs.google.com/spreadsheets/d/1UAY4aXERKS2409RWxnbyNZjoIMe0vMVQPHOPcU8vvDU/edit#gid=1985928954

### How to add items on the list? 

1.	Upload the file in this link:
https://drive.google.com/drive/folders/1uxBcON0LfHhZojlJ6ATqQkT2KeVf3-5D

2.	After successfully uploading, right click > share > copy link 
 ![Lazy Agent Image](Lazy%20Agent%20Images/image034.png)

3.	At the bottom of the Resources sheet, Enter the name of the file you want to add.
4.	Paste the link on the cell right next to it. 
 ![Lazy Agent Image](Lazy%20Agent%20Images/image036.png)
