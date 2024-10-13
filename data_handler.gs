/** @OnlyCurrentDoc */

// Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var lock = LockService.getScriptLock();
var ui = SpreadsheetApp.getUi();

// Consolidate sheet references into an object
var sheets = {
  report: ss.getSheetByName("Report"),
  data1: ss.getSheetByName("data1"),
  data2: ss.getSheetByName("data2"),
  data3: ss.getSheetByName("data3"),
  data4: ss.getSheetByName("data4"),
  data5: ss.getSheetByName("data5"),
  data6: ss.getSheetByName("data6"),
  data7: ss.getSheetByName("data7"),
  data8: ss.getSheetByName("data8"),
  data9: ss.getSheetByName("data9"),
  data10: ss.getSheetByName("data10"),
  agent: ss.getSheetByName("Agents"),
  credit: ss.getSheetByName("Credits"),
  loan: ss.getSheetByName("Loans"),
  exchange: ss.getSheetByName("Exchanges"),
  mProfit: ss.getSheetByName("Monthly Profits"),
  records: ss.getSheetByName("Records"),
  note: ss.getSheetByName("Notes"),
  // Add other sheet references here...
};

// Function to get sheet by name
function getSheet(name) {
  return sheets[name];
}

// Function to get data from report sheet
function getReportData(range) {
  return sheets.report.getRange(range).getValue();
}

// Function to set data to a specific sheet and range
function setData(sheetName, data, range) {
  var sheet = getSheet(sheetName);
  if (sheet) {
    sheet.getRange(range).setValues([data]);
  }
}

/**
 * Function to clear ranges, range lists, or individual cells on a sheet.
 * @param {Sheet} sheet - The sheet from which the ranges will be cleared.
 * @param {string[] | string[][]} rangesOrRangeLists - An array of range strings or range lists to clear.
 */
function clearRanges(sheet, rangesOrRangeLists) {
  if (!Array.isArray(rangesOrRangeLists)) return;

  rangesOrRangeLists.forEach(function (rangeOrRangeList) {
    if (Array.isArray(rangeOrRangeList)) {
      // If it's a range list, clear each range in the list
      rangeOrRangeList.forEach(function (range) {
        sheet
          .getRange(range)
          .clear({ contentsOnly: true, skipFilteredRows: true });
      });
    } else {
      // If it's a single range, clear it
      sheet
        .getRange(rangeOrRangeList)
        .clear({ contentsOnly: true, skipFilteredRows: true });
    }
  });
}

// Function to save data for various sheets
function saveData(dataRanges, targetSheets) {
  dataRanges.forEach(function (dataRange, index) {
    var data = dataRange.map(function (range) {
      return getReportData(range);
    });
    // Get the target sheet
    var sheet = getSheet(targetSheets[index]);
    if (sheet) {
      // Get the last row of the target sheet
      var lastRow = sheet.getLastRow();

      // Populate data starting from the next row after the last row
      sheet.getRange(lastRow + 1, 1, 1, data.length).setValues([data]);
    }
  });
}

// Save report data function
function saveReportData() {
  var dataRanges = [
    ["A15", "B17", "B15", "K3", "K4", "K5", "K6", "K7", "K8", "B15"], //Agent
    ["B17", "A15", "B3", "B4", "B5", "B6", "B7", "B8"], //data1
    ["B17", "A15", "C3", "C4", "C5", "C6", "C7", "C8"], //data2
    ["B17", "A15", "D3", "D4", "D5", "D6", "D7", "D8"], //data3
    ["B17", "A15", "E3", "E4", "E5", "E6", "E7", "E8"], //data4
    ["B17", "A15", "F3", "F4", "F5", "F6", "F7", "F8"], //data5
    ["B17", "A15", "G3", "G4", "G5", "G6", "G7", "G8"], //data6
    ["B17", "A15", "H3", "H4", "H5", "H6", "H7", "H8"], //data7
    ["B17", "A15", "I3", "I4", "I5", "I6", "I7", "I8"], //data10
    ["B17", "A15", "J3", "J4", "J5", "J6", "J7", "J8"], //data8
    ["B17", "A15", "K3", "K4", "K5", "K6", "K7", "K8"], //data9
    // Add other data ranges here...
  ];
  var targetSheets = [
    "agent",
    "data1",
    "data2",
    "data3",
    "data4",
    "data5",
    "data6",
    "data7",
    "data10",
    "data8",
    "data9", // Add other target sheets here...
  ];
  saveData(dataRanges, targetSheets);
}

// Function to save  credit data
function saveCreditData() {
  var ranges = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];
  var dataRanges = [];

  // Iterate through each column range
  ranges.forEach(function (range) {
    var value = sheets.report.getRange(range + "10").getValue();
    // Check if the value is not 0
    if (value != 0) {
      var rowData = [
        sheets.report.getRange("B17").getValue(),
        sheets.report.getRange("A15").getValue(),
        sheets.report.getRange(range + "1").getValue(),
        value,
      ];
      dataRanges.push(rowData);
    }
  });

  // Set data to credit sheet
  dataRanges.forEach(function (data) {
    sheets.credit
      .getRange(sheets.credit.getLastRow() + 1, 1, 1, data.length)
      .setValues([data]);
  });
}

// Function to save loan record data
function saveLoanRecordData() {
  var loan_data = [];
  var loanRanges = ["D", "E", "F", "G"];

  // Iterate through each loan range
  loanRanges.forEach(function (range) {
    // Check if the range is not blank
    if (
      !sheets.report.getRange(range + "15").isBlank() &&
      !sheets.report.getRange(range + "14").isBlank()
    ) {
      var rowData = [
        sheets.report.getRange("B17").getValue(),
        sheets.report.getRange(range + "14").getValue(),
        sheets.report.getRange(range + "15").getValue(),
        sheets.report.getRange("A15").getValue(),
        sheets.report.getRange(range + "13").getValue(),
      ];
      loan_data.push(rowData);
    }
  });

  // Set data to loan sheet
  loan_data.forEach(function (data) {
    sheets.loan
      .getRange(sheets.loan.getLastRow() + 1, 1, 1, data.length)
      .setValues([data]);
  });
}

// Function to save monthly exchange history
function saveExchangeHistory() {
  var exchanged = [];

  // Check if the range is not blank
  if (!sheets.report.getRange("H15").isBlank()) {
    exchanged.push([
      sheets.report.getRange("B17").getValue(),
      sheets.report.getRange("H15").getValue(),
      sheets.report.getRange("I15").getValue(),
    ]);
  }

  // Set data to exchange sheet
  exchanged.forEach(function (data) {
    sheets.exchange
      .getRange(sheets.exchange.getLastRow() + 1, 1, 1, data.length)
      .setValues([data]);
  });
}

// Function to save notes
function saveNotes() {
  var notes_data = [];

  // Iterate through each note range
  for (var row = 21; row <= 25; row++) {
    if (!sheets.report.getRange("C" + row + ":H" + row).isBlank()) {
      notes_data.push([
        sheets.report.getRange("A15").getValue(),
        sheets.report.getRange("B17").getValue(),
        sheets.report.getRange("C" + row).getValue(),
        sheets.report
          .getRange("D" + row + ":H" + row)
          .getMergedRanges()[0]
          .getValue(),
      ]);
    }
  }

  // Set data to note sheet
  notes_data.forEach(function (data) {
    sheets.note
      .getRange(sheets.note.getLastRow() + 1, 1, 1, data.length)
      .setValues([data]);
  });

  // Clear note ranges
  if (notes_data.length > 0) {
    clearRanges(sheets.report, ["C21:H25"]);
  }
}

// Reset report data
function resetAll() {
  var reportSheet = sheets.report;

  // Set active sheet to Report
  ss.setActiveSheet(reportSheet, true);

  // Copy and paste values from B9:K9 to B2
  reportSheet
    .getRange("B9:K9")
    .copyTo(
      reportSheet.getRange("B2"),
      SpreadsheetApp.CopyPasteType.PASTE_VALUES,
      false
    );

  // Copy and paste values from L15 to C15
  reportSheet
    .getRange("L15")
    .copyTo(
      reportSheet.getRange("C15"),
      SpreadsheetApp.CopyPasteType.PASTE_VALUES,
      false
    );
}

// Clear report data
function clearReportData() {
  var reportSheet = sheets.report;

  // Set active sheet to Report
  ss.setActiveSheet(reportSheet, true);

  // Clear content from specified ranges
  reportSheet
    .getRangeList(["B3:K9", "D15:J15", "D14:G14"])
    .clear({ contentsOnly: true, skipFilteredRows: true });
}

// Clear monthly records
function clearRecords() {
  var sheetsToClear = ["Newperson", "Exchanges", "Credits", "Notes", "Agents"];
  var rangesToClear = {
    Newperson: ["B3:F50", "I3:M50", "P3:T50"],
    Exchanges: ["A3:C100"],
    Credits: ["A3:D200"],
    Notes: ["A2:D100"],
    Agents: ["A2:K300"],
  };

  sheetsToClear.forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var range = rangesToClear[sheetName];
      clearRanges(sheet, range);
    }
  });
}

// Clear monthly  records
function clearRecords() {
  var Sheets = [
    "data1",
    "data2",
    "data3",
    "data4",
    "data5",
    "data6",
    "data7",
    "data10",
    "data8",
    "data9",
  ];
  var rangeToClear = ["A2:H300"];

  Sheets.forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      clearRanges(sheet, rangeToClear);
    }
  });

  ss.setActiveSheet(ss.getSheetByName("Report"), true);
  ss.getRange("A17").activate();
}

// Function to split worked hours in the "J" column of the agent sheet
function splitWorkedHours() {
  var lastRow = sheets.agent.getLastRow();
  var shiftRange = sheets.agent.getRange("J" + lastRow);

  // Check if the range is not empty before splitting
  if (!shiftRange.isBlank()) {
    // Split text into columns based on the "-" delimiter
    shiftRange.splitTextToColumns("-");
  }
}

// Reset function
function resetReport() {
  var dataFunctions = [
    saveReportData,
    splitWorkedHours,
    saveCreditData,
    saveNotes,
  ];
  var resetFunctions = [resetAll, clearReportData];

  // Execute data saving functions
  dataFunctions.forEach(function (func) {
    func();
  });

  // Save loan record if the range is not blank
  if (!sheets.report.getRange("D15:G15").isBlank()) {
    saveLoanRecordData();
  }

  // Save exchange record if the range is not blank
  if (!sheets.report.getRange("H15:I15").isBlank()) {
    saveExchangeHistory();
  }

  // Execute data resetting functions
  resetFunctions.forEach(function (func) {
    func();
  });

  // Set shift formula
  var cell = sheets.report.getRange("B15");
  cell.setFormula(
    "=IF(A15=Records!A14, Records!B14, IF(A15=Records!A15, Records!B15, IF(A15=Records!A16, Records!B16, IF(A15=Records!A17, Records!B17, IF(A15=Records!A18, Records!B18, IF(A15=Records!A19, Records!B19, IF(A15=Records!A20, Records!B20, IF(A15=Records!A21, Records!B21, Records!B22))))))))"
  );
}

// Function to check if another instance of the script is already running
function isScriptAlreadyRunning() {
  return !lock.tryLock(0);
}

// Function to display a message
function showMessage(message, title) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title || "Info", message, ui.ButtonSet.OK);
}

// Function to display a  confirmation dialog if there is a shortover
function shortoverConfirmationDialog(shortOverValue) {
  var result = ui.alert(
    "⚠️ Warning",
    "The shortover is not zero (currently " +
      shortOverValue +
      "). Please check the following:\n\n1. Verify that the shortover is zero or not.\n2. If the shortover cell is appearing 'Red' even if the shortover is zero, then you can continue with the reset by clicking 'Yes'.\n3. If the shortover is supposed to be correct even if it is not zero, then as well you can continue with the reset by clicking 'Yes'.\n(🚨 Alert: Don't forget to write the 📝 note regarding the shortover in case of 3!!)\n4. Otherwise click on 'No' to cancel it.",
    ui.ButtonSet.YES_NO
  );

  return result === ui.Button.YES;
}

// Function to display a confirmation dialog if the agent name is repeating
function agentNameConfirmationDialog(agentName) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "⚠️ Warning",
    "The agent name " +
      agentName +
      " is repeating in the sheet. Either reset is executing multiple times or Please check the following:\n\n1. Verify that the agent name is correct.\n2. Check the last entry made by last reset in the Agents sheet.\n3. Rectify any other errors found in the report.\n4. Click on 'Yes' if the agent/agents name are expected to be repeated (rare scenario but possible)\n5. Otherwise click on 'No' to fix the issue first.\nIf the issue persists, please contact support for further assistance.",
    ui.ButtonSet.YES_NO
  );

  return result === ui.Button.YES;
}

// Reset button function
function reset() {
  // Check if another instance of the script is already running
  if (isScriptAlreadyRunning()) {
    showMessage("The script is already running. Please wait!", "⚠️ Warning!!");
    return; // Abort the reset function
  }

  // Get the value in cell H17
  var shortOverValue = sheets.report.getRange("H17").getValue();

  // Check if the shortover value is not zero
  if (shortOverValue !== 0) {
    // Show confirmation dialog explaining the reason
    var continueReset = shortoverConfirmationDialog(shortOverValue);

    // If the user clicks "No", abort the script
    if (!continueReset) {
      return; // Abort the reset function
    }
  }

  // Check if the value in cell A15 (Agent name) of Report sheet and last row value in the first column of Agent sheet match
  var matchAgentReport =
    sheets.report.getRange("A15").getValue() ===
    sheets.agent.getRange(sheets.agent.getLastRow(), 1).getValue();

  if (matchAgentReport) {
    // Get the agent name for showing in the confirmation dialog
    var agentName = sheets.report.getRange("A15").getValue();

    // Show confirmation dialog explaining the reason
    var continueReset = agentNameConfirmationDialog(agentName);

    if (!continueReset) {
      // User opted not to continue with the reset process
      return; // Abort the reset function
    }
  }

  /** Check if the report isn't missing any final credits */
  // Check if all specified ranges are not blank
  var rangeChecks = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];
  var allRangesNotBlank = rangeChecks.every(function (range) {
    return !sheets.report.getRange(range + "10").isBlank();
  });

  if (!allRangesNotBlank) {
    // Get the first empty column index
    var emptyColumnIndex =
      rangeChecks.findIndex(function (range) {
        return sheets.report.getRange(range + "10").isBlank();
      }) + 2; // Adding 2 to match the first row index

    // Get the corresponding column letter
    var emptyColumnLetter = String.fromCharCode(64 + emptyColumnIndex);

    // Get the value of the first empty cell in the first row
    var emptyColumnValue = sheets.report
      .getRange(1, emptyColumnIndex)
      .getValue();

    // Show warning message
    showMessage(
      "You are missing final credit in column " +
        emptyColumnLetter +
        " for " +
        emptyColumnValue +
        ". Please ensure that all required fields are filled in the report.",
      "⚠️ Incomplete Report"
    );
    return; // Abort the reset function
  }
  /** End */

  // Check if the loan record isn't missing cashtag
  if (!sheets.report.getRange("D15:G15").isBlank()) {
    var loanRanges = ["D", "E", "F", "G"];
    var isLoanRecordComplete = true;

    for (var i = 0; i < loanRanges.length; i++) {
      var range = loanRanges[i];
      if (
        !sheets.report.getRange(range + "15").isBlank() &&
        sheets.report.getRange(range + "14").isBlank()
      ) {
        var loanStatusForMissingCashtag = sheets.report
          .getRange(range + "13")
          .getValue();
        showMessage(
          "Cashtag is missing for " +
            loanStatusForMissingCashtag +
            " in cell " +
            range +
            "14 for loan record. Please ensure that all required fields are filled in the report.",
          "⚠️ Incomplete Report"
        );
        isLoanRecordComplete = false;
        break; // Exit the loop
      }
    }

    // Check if any loan record is incomplete and abort the reset function
    if (!isLoanRecordComplete) {
      return; // Abort the reset function
    }
  }

  // Lock the script
  lock.waitLock(20000); // Wait 20 seconds for others' use of the code section and lock to stop and then proceed

  try {
    // Call the reset report function
    resetReport();
    showMessage(
      "Reset successfull! Data has been cleared and updated.",
      "🔄 Reset Successfull"
    );
  } catch (e) {
    // Handle any errors that may occur during ResetReport() function
    showMessage("An error occurred: " + e.message, "🛑 Error!!");
  } finally {
    // Release the lock
    lock.releaseLock();
  }
}

// Reset Functionality for mobile
function onEdit(e) {
  var range = e.range;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var column = range.getColumn();
  var row = range.getRow();

  // Check if the edited cell is in the "Report" sheet and matches the desired cell
  if (sheetName === "Report" && column === 10 && row === 20) {
    reset();
  }
}

// Function to exchange the end wallet value of the month
function exchangewalletAmt() {
  var walletAmt = [
    [
      sheets.report.getRange("B17").getValue(),
      sheets.report.getRange("K15").getValue(),
    ],
  ];
  sheets.exchange
    .getRange(
      sheets.exchange.getLastRow() + 1,
      1,
      walletAmt.length,
      walletAmt[0].length
    )
    .setValues(walletAmt);

  // Clear monthly wallet cash amount
  var rangesToClear = ["B21:B25", "B42:B46", "C15"];
  clearRanges(sheets.report, rangesToClear);
}

// Function to exchange the loan net amount of the month
function exchangeLoanAmt() {
  var loanAmt = [
    [
      sheets.report.getRange("B17").getValue(),
      sheets.report.getRange("C17").getValue(),
    ],
  ];
  sheets.exchange
    .getRange(
      sheets.exchange.getLastRow() + 1,
      1,
      loanAmt.length,
      loanAmt[0].length
    )
    .setValues(loanAmt);

  // Clear monthly loan record data
  clearRanges(sheets.loan, ["A2:E500"]);
}

// Function to record monthly profit data
function saveMonthlyProfitData() {
  var profitData = [
    [
      sheets.report.getRange("B17").getValue(),
      sheets.report.getRange("A17").getValue(),
      sheets.report.getRange("E17").getValue(),
      sheets.report.getRange("G17").getValue(),
    ],
  ];
  sheets.mProfit
    .getRange(
      sheets.mProfit.getLastRow() + 1,
      1,
      profitData.length,
      profitData[0].length
    )
    .setValues(profitData);
}

// Monthly reset button function
function monthlyReset() {
  // Check if another instance of the script is already running
  if (isScriptAlreadyRunning()) {
    showMessage("The script is already running. Please wait!", "⚠️ Warning!!");
    return;
  }

  // Lock the script
  lock.waitLock(20000); // Wait 20 seconds for others' use of the code section and lock to stop and then proceed

  try {
    // Clear and exchange the wallet net amount of the month
    // exchangewalletAmt();

    // Clear and exchange the loan net amount of the month
    exchangeLoanAmt();

    // Record monthly profit data
    saveMonthlyProfitData();

    // Clear monthly records
    clearRecords();

    // Clear  records
    clearRecords();

    showMessage("Monthly reset completed successfully!", "🎉 Congratulation");
  } catch (e) {
    // Handle any errors that may occur
    showMessage("An error occurred: " + e.message, "🛑 Error!!");
  } finally {
    // Release the lock
    lock.releaseLock();
  }
}

// Set work hour formula for agents
function setWorkHourFormula() {
  var cell = sheets.agent.getRange("L2:L300");
  cell.setFormula(
    '=IF(J2="", , IF(IFERROR(K2-J2+(J2>K2),)=0, 1,IFERROR(K2-J2+(J2>K2))))'
  );
}

// Refactored Test function
function test() {
  // Implement as per your logic...
}

// Create custom menu
function createCustomMenu() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Custom Menu")
    .addItem("Monthly Reset", "monthlyReset")
    .addItem("Reset", "reset")
    .addItem("Clear Report Data", "clearReportData")
    .addItem("Save Loan Report Data", "saveLoanRecordData")
    .addItem("Save Exchange History", "saveExchangeHistory")
    .addItem("Exchange wallet Amount", "exchangewalletAmt")
    .addItem("Exchange Loan Amount", "exchangeLoanAmt")
    .addItem("Clear Monthly Records", "clearRecords")
    .addItem("Clear  Records", "clearRecords")
    .addItem("Save Notes", "saveNotes")
    .addItem("Set Work Hour Formula", "setWorkHourFormula")
    .addItem("Test", "test")
    .addToUi();
}

function onOpen() {
  createCustomMenu();
}
