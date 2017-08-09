
# exportGSheet
Simple PDF export for individual Google Sheets + UI interface in Google Apps Script.

*Based on a fork from [Cliff Hazell](https://gist.github.com/ixhd/3660885)* 

**Features**
- Export a **single active spreadsheet** as PDF with one click. 
- Export either the full sheet or a highlighted range. 
- Exports formula-calculated cells as display values to avoid #REF errors.
- Creates new button on the UI menu row called 'Save to PDF', with options.


**Steps** 
1. Copy the entire script to the Google sheets file you want to use. 
    Menu bar > Tools > Script Editor...

2. Replace the values on lines 2 - 4 of the code, to set the name & destination of the exported file.
```javascript 
// change these values
var name = ""; // leave as empty string if you are getting value from a cell in ActiveSheet 
var nameRange = "<Cell range>"; // spell out the cell where you are reading the file name as, eg. "A1"
var folderId = "<Desination folder Id>"; // Google Id of destination folder, eg. "0B2_h6nTAN7gBU3ZLRFVkLmxVYkU"
```
3. Run within Script Editor or reload your Sheets graphical UI to see new button. Presto!
