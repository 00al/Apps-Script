function skuGenerator() {
  //take titles from column in sheet, output titles & skus in another sheet if they don't already exist
  ar1 = arrayFromSheet("NewTitles", 1, 4); //define the sheet name, start row and column indices from which to take new titles
  an = "ExistingTitlesAndSKUs"; //define the sheet name where existing titles are stored
  ar2 = arrayFromSheet(an, 1, 0); //define the start row and column indices of the sheet where existing titles are (to be) stored
  arN = []; //array to store new titles & skus
  for (var i = 0; i < ar1.length; i++) {
    //loop through new titles sheet, create new SKUs for any that don't already exist
    if (ar2.indexOf(ar1[i]) == -1) {
      arN.push([ar1[i], skuFormatter(GetMD5Hash(ar1[i]))]); //create new SKUs by hashing the title with GetMD5Hash function and formatting it via skuFormatter function
    }
  }
  var ui = SpreadsheetApp.getUi();
  // if there are new titles, append new titles and SKUs to sheet
  if (arN.length > 0) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(an);
    var lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, arN.length, arN[0].length).setValues(arN);
    ui.alert(i + " new SKUs added to the bottom of " + an);
    // else stop
  } else {
    ui.alert("There are no new SKUs to generate");
  }
}

function skuFormatter(hashedTitle) {
  //take a hashed title, format it and then return it
  var pre = "F-"; //define what to start all SKUs with
  var beg = hashedTitle.slice(0, 3); //define which part of the hashed title to use for the second part of the SKU
  var mid = hashedTitle.slice(14, 17); //define which part of the hashed title to use for the third part of the SKU
  var end = hashedTitle.slice(29, 32); //define which part of the hashed title to use for the fourth part of the SKU
  var sku = (pre + beg + "-" + mid + "-" + end).toUpperCase(); //add all parts together and make uppercase
  return sku;
}

function arrayFromSheet(sheetName, row, col) {
  //create an array from a sheet. define sheet name as string, start row & column as int
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //required to getSheetByName below
  var sh = ss.getSheetByName(sheetName); //define sheet
  var values = sh.getDataRange().getValues(); //get all values in an indexed 2D array [ROW][COL]
  var shtarr = []; //an (empty) array to store the values
  for (n = row; n < values.length; ++n) {
    //start at row
    var cell = values[n][col]; //[col] is the index of the column starting from 0
    shtarr.push(cell); // populate the array with cell values
  }
  return shtarr; //return the array
}

function GetMD5Hash(value) {
  //input a text value to hash and return the hashed value as text
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, value);
  var txtHash = "";
  for (j = 0; j < rawHash.length; j++) {
    var hashVal = rawHash[j];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += "0";
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function onOpen() {
  //upon opening spreadsheet create a sub menu and button to run skuGenerator function
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name: "Go!",
      functionName: "skuGenerator",
    },
  ];
  ss.addMenu("Generate SKUs", entries);
}
