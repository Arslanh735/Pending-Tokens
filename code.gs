function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle("Saylani ERP Elite")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pending_Storage");
  var stats = { "Chicken": 0, "Goat": 0, "Cow": 0, "Camel": 0 };

  if (!sheet) return stats;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return stats;

  for (var i = 1; i < data.length; i++) {
    var animal = data[i][0];
    var qty = Number(data[i][4]);
    if (stats.hasOwnProperty(animal) && !isNaN(qty)) {
      stats[animal] += qty;
    }
  }
  return stats;
}

function getFilteredReport(branchName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("IssuedRecords");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const filteredRows = data.slice(1).filter(row => row[1] === branchName);

  return { headers: headers, rows: filteredRows };
}

function saveTokens(dataArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName("Pending_Storage") || ss.insertSheet("Pending_Storage");
  var masterSheet = ss.getSheetByName("Master_Received_Log") || ss.insertSheet("Master_Received_Log");

  if (pendingSheet.getLastRow() === 0) {
    pendingSheet.appendRow(["Animal", "Rate", "DN", "Donor", "Qty", "Branch", "Date", "Category"]);
  }

  dataArray.forEach(function (row) {
    var rowData = [row.animal, row.rate, row.dn, row.donor, row.qty, row.branch, new Date(), row.animal_cat];
    pendingSheet.appendRow(rowData);
    masterSheet.appendRow(rowData);
  });
  return "Data Saved Successfully!";
}

function processBulkIssue(items, mainBranch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName("Pending_Storage");
  if (!pendingSheet) return { status: "Error", msg: "Pending Storage sheet not found!" };

  var usedSheet = ss.getSheetByName("Used_Storage") || ss.insertSheet("Used_Storage");

  if (usedSheet.getLastRow() === 0) {
    usedSheet.appendRow(["Category","Rate","DN_Number","Donor_Name","Quantity","Branch","Date","Voucher_No","Type","V_No_Master"]);
  }

  var lastRow = usedSheet.getLastRow();
  var lastVoucher = lastRow > 1 ? usedSheet.getRange(lastRow, 10).getValue() : "";

  var nextNumber = 1;
  if (lastVoucher) {
    var num = parseInt(lastVoucher.toString().replace("V-", ""));
    if (!isNaN(num)) nextNumber = num + 1;
  }

  var vNo = "V-" + nextNumber.toString().padStart(4, '0'); // V-0001 format

  var data = pendingSheet.getDataRange().getValues();
  var issuedItems = [];

  for (var j = 0; j < items.length; j++) {
    var req = items[j];
    var remainingToIssue = Number(req.qty);

    var reqAnimal = req.animal.toString().trim().toUpperCase();
    var reqRate = req.rate.toString().trim();
    var reqCat = req.category.toString().trim().toUpperCase();

    for (var i = 1; i < data.length; i++) {
      if (remainingToIssue <= 0) break;

      var sheetAnimal = data[i][0].toString().trim().toUpperCase();
      var sheetRate = data[i][1].toString().trim();
      var sheetCat = data[i][7] ? data[i][7].toString().trim().toUpperCase() : "";

      if (sheetAnimal === reqAnimal && sheetRate === reqRate && sheetCat === reqCat) {

        var sQty = Number(data[i][4]);
        var takeQty = Math.min(sQty, remainingToIssue);

        issuedItems.push({
          animal: data[i][0],
          rate: data[i][1],
          dn: data[i][2],
          donor: data[i][3],
          qty: takeQty,
          cat: data[i][7]
        });

        usedSheet.appendRow([
          data[i][0],   
          data[i][1],   
          data[i][2],   
          data[i][3],   
          takeQty,      
          mainBranch,   
          new Date(),   
          vNo,          
          data[i][7],   
          vNo           
        ]);

        if (sQty === takeQty) {
          pendingSheet.deleteRow(i + 1);
          data.splice(i, 1);
          i--;
        } else {
          pendingSheet.getRange(i + 1, 5).setValue(sQty - takeQty);
          data[i][4] = sQty - takeQty;
        }

        remainingToIssue -= takeQty;
      }
    }

    if (remainingToIssue > 0) {
      return {
        status: "Error",
        msg: "Stock shortage for " + req.animal + " (" + reqCat + "). Needed: " + remainingToIssue
      };
    }
  }

  return {
    status: "OK",
    vNo: vNo,
    branch: mainBranch,
    items: issuedItems,
    date: new Date().toLocaleDateString()
  };
}

function getAnimalDetails(animalName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pending_Storage");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var details = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == animalName) {
      details.push({ cat: data[i][7], rate: data[i][1], qty: data[i][4], donor: data[i][3], dn: data[i][2] });
    }
  }
  return details;
}

function generateBranchReport(branchName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("IssuedRecords");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const filteredData = data.filter((row, index) => {
    if (index === 0) return false;
    return row[1] === branchName;
  });

}

function getVoucherDetails(vNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var items = [];
  var branch = "";
  var date = "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][7] == vNo) {
      branch = data[i][5];
      date = data[i][6] ? Utilities.formatDate(new Date(data[i][6]), "GMT+5", "dd-MM-yyyy") : "";

      items.push({
        animal: data[i][0],
        rate: data[i][1],
        dn: data[i][2],
        donor: data[i][3],
        qty: data[i][4],
        cat: data[i][8]
      });
    }
  }

  if (items.length > 0) {
    return {
      vNo: vNo,
      branch: branch,
      date: date,
      items: items 
    };
  }
  return null;
}
function getReportsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");

  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var reports = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    try {
      var rawDate = row[6];
      var formattedDate = "";
      var displayDate = "";

      if (rawDate instanceof Date) {
        formattedDate = Utilities.formatDate(rawDate, "GMT+5", "yyyy-MM-dd");
        displayDate = Utilities.formatDate(rawDate, "GMT+5", "dd-MM-yyyy");
      }

      reports.push({
        animal: row[0] || "",
        rate: row[1] || 0,
        dn: row[2] || "",
        donor: row[3] || "",
        qty: row[4] || 0,
        branch: row[5] || "",
        date: formattedDate,
        displayDate: displayDate,
        vNo: row[7] || "",
        cat: row[8] || ""
      });
    } catch (e) {
      console.log("Error in row " + i + ": " + e.message);
    }
  }
  return reports;
}
function getTrendData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var startRow = Math.max(2, lastRow - 499);
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, 1, numRows, 9).getValues(); 

  var results = [];
  for (var i = 0; i < data.length; i++) {
    var dateVal = data[i][6];
    var qtyVal = Number(data[i][4]) || 0;
    
    if (dateVal instanceof Date) {
      results.push({
        date: dateVal.toISOString(), 
        qty: qtyVal
      });
    }
  }
  return results;
}
function getBranchDetails(branchName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];

  var searchBranch = branchName.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowBranch = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";

    if (rowBranch === searchBranch || rowBranch.includes(searchBranch)) {

      var rawDate = data[i][6];
      var displayDate = "";

      if (rawDate instanceof Date) {
        displayDate = Utilities.formatDate(rawDate, "GMT+5", "dd-MM-yyyy");
      }

      result.push({
        animal: data[i][0],
        rate: data[i][1],
        qty: data[i][4],
        branch: data[i][5],
        displayDate: displayDate,
        vNo: data[i][7]
      });
    }
  }

  return result;
}
function getBranchHierarchy(branchName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  var result = {};

  var searchBranch = branchName.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowBranch = data[i][5] ? data[i][5].toString().trim().toLowerCase() : "";

    if (rowBranch !== searchBranch) continue;

    var rawDate = data[i][6];
    if (!(rawDate instanceof Date)) continue;

    var month = Utilities.formatDate(rawDate, "GMT+5", "MMMM yyyy");
    var date = Utilities.formatDate(rawDate, "GMT+5", "dd-MM-yyyy");

    var rate = data[i][1];
    var qty = Number(data[i][4]) || 0;

    if (!result[month]) result[month] = {};
    if (!result[month][date]) result[month][date] = {};

    result[month][date][rate] = (result[month][date][rate] || 0) + qty;
  }

  return result;
}
function verifyDN(dn) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");

  var data = sheet.getDataRange().getValues();

  var results = [];

  for (var i = 1; i < data.length; i++) {

    var rowDN = data[i][2]; // Column C = DN

    if (rowDN == dn) {

      var qty = Number(data[i][4]); // Qty
      var verified = Number(data[i][10] || 0); // K column

      results.push({
        row: i + 1,
        animal: data[i][0],
        rate: data[i][1],
        qty: qty,
        verified: verified,
        remaining: qty - verified,
        branch: data[i][5],
        voucher: data[i][7],
        type: data[i][8]
      });
    }
  }

  return results;
}
function consumeDN(rowNo){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");

  var qty = Number(sheet.getRange(rowNo, 5).getValue());
  var verified = Number(sheet.getRange(rowNo, 11).getValue() || 0);

  if(verified >= qty){
    return "LIMIT_EXCEEDED";
  }

  sheet.getRange(rowNo, 11).setValue(verified + 1);

  return "OK";
}
function doGet(e) {

  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "index";

  if (page === "view") {
    return HtmlService.createHtmlOutputFromFile("view")
      .setTitle("Management View");
  }

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ERP System");
}
function searchVoucher(value) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Used_Storage");

  var data = sheet.getDataRange().getValues();

  var result = [];

  for (var i = 1; i < data.length; i++) {

    if (
      data[i][7] == value ||   // Voucher No
      data[i][2] == value      // DN
    ) {
      result.push({
        vNo: data[i][7],
        branch: data[i][5],
        animal: data[i][0],
        qty: data[i][4]
      });
    }
  }

  return result;
}
function getFullPendingStock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pending_Storage");
  
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var animal = row[0];
    var rate = row[1];
    var qty = Number(row[4]) || 0;

    if (animal && qty > 0) {
      var key = animal + "|" + rate;
      if (!summary[key]) {
        summary[key] = {
          animal: animal,
          rate: Number(rate),
          qty: 0
        };
      }
      summary[key].qty += qty;
    }
  }

  // Convert to array and sort
  var result = Object.values(summary);
  result.sort((a, b) => {
    if (a.animal === b.animal) return a.rate - b.rate;
    return a.animal.localeCompare(b.animal);
  });

  return result;
}
function generateFullStockPDF(data) {
  // Yeh client side pe call hoga
  // Hum JS mein PDF bana rahe hain
}