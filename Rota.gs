function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a comprehensive menu combining all the functionalities
  ui.createMenu('Rota Management')
    .addItem('Trim White Spaces', 'removeLeadingTrailingSpaces')
    .addItem('Check for Clashes', 'newScript')
    .addItem('Generate Shift Report', 'generateShiftReport')
    .addItem('Generate Full Report', 'generateFullReport')
    .addItem('Generate Location Report', 'generateLocationReport')
    .addItem('Query by Total Hours', 'queryByTotalHours')
    .addItem('Full Report Alphabetical Order', 'generateFullReportAlphabetical')
    .addItem('Full Report - Least to Most Hours', 'generateFullReportLeastHours')
    .addItem('Full Report - Most to Least Hours', 'generateFullReportMostHours')
    .addItem('Name and Total Hours Report', 'generateNameAndTotalHoursReport')
    .addToUi();
}

// Include all other functions here as they are
function newScript(e) {
  clearHighlights();
  highlightMatchingCellsInRange();
}

function clearHighlights() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("D2:R63");
  dataRange.setBackground(null);
}

function highlightMatchingCellsInRange() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("D2:R63"); // Specify the desired range
  var data = dataRange.getValues(); // Get the values in the specified range

  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[0].length; col++) {
      var cellValue = data[row][col].toString().trim().toLowerCase(); // Convert to lowercase and remove leading/trailing spaces

      if (cellValue !== "") {
        // Check for clashes within the same row
        for (var i = 0; i < data[0].length; i++) {
          if (i !== col) {
            var adjacentValue = data[row][i].toString().trim().toLowerCase();

            if (cellValue === adjacentValue) {
              var rangeToHighlight = dataRange.getCell(row + 1, col + 1);
              rangeToHighlight.setBackground("red");
            }
          }
        }

        // Check for consecutive shift clashes across different sites
        if (row < data.length - 1) {
          for (var nextCol = 0; nextCol < data[0].length; nextCol++) {
            var nextCellValue = data[row + 1][nextCol].toString().trim().toLowerCase();

            if (cellValue === nextCellValue) {
              var currentRange = dataRange.getCell(row + 1, col + 1);
              var nextRange = dataRange.getCell(row + 2, nextCol + 1);
              currentRange.setBackground("red");
              nextRange.setBackground("red");
            }
          }
        }
      }
    }
  }
}


function removeLeadingTrailingSpaces() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1'); // Replace 'Sheet1' with your sheet name
  var range = sheet.getRange('D2:R63'); // Replace with the desired range

  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        var trimmedValue = values[i][j].trim(); // Remove leading and trailing spaces

        if (values[i][j] !== trimmedValue) {
          values[i][j] = trimmedValue; // Update the cell value to the trimmed value
        }
      }
    }
  }

  range.setValues(values); // Update the range with the corrected values
}

function onEdit(e) {
  // Define the threshold in hours
  var hoursThreshold = 270; // Set your hours threshold here

  // Get the active sheet and the edited range
  var sheet = e.source.getActiveSheet();
  var editedRange = e.range;

  // Check if the edit is within the specified range D2:R63
  if (editedRange.getColumn() >= 4 && editedRange.getColumn() <= 15 && editedRange.getRow() >= 2 && editedRange.getRow() <= 63) {
    var newValue = e.value;
    var shiftTypeRange = sheet.getRange("C2:C63").getValues();
    var locationHeaders = sheet.getRange("D1:O1").getValues()[0];
    var shiftDataRange = sheet.getRange("D2:R63").getValues();

    var totalHours = 0;

    // Calculate total hours for newValue
    for (var i = 0; i < shiftDataRange.length; i++) {
      for (var j = 0; j < shiftDataRange[0].length; j++) {
        if (shiftDataRange[i][j] === newValue) {
          var shiftType = shiftTypeRange[i][0];
          var location = locationHeaders[j];
          var hours = calculateHours(shiftType, location);
          totalHours += parseHours(hours);
        }
      }
    }

    // Check against the hours threshold and revert if necessary
    if (totalHours > hoursThreshold) {
      // Revert the cell to its previous value if the hours threshold is exceeded
      editedRange.setValue(e.oldValue);
    }
  }
}

function calculateHours(shiftType, location) {
  if (shiftType === 'Night' && location === 'Symons Rd (DW)') {
    return '8h';
  } else if (shiftType === 'Day' && location === '115a Northend Rd (AI)') {
    return '4h';
  } else if (shiftType === 'Day' && location === 'Gads Hill Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Kitchener Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Sydney Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Peareswood Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Symons Mid Shift') {
    return '8h';
  } else {
    return '12h';
  }
}

function parseHours(hoursString) {
  var parts = hoursString.split(' ');
  var hours = 0;
  parts.forEach(function(part) {
    if (part.includes('h')) {
      hours += parseInt(part);
    } else if (part.includes('m')) {
      hours += parseInt(part) / 60;
    }
  });
  return hours;
}


function generateShiftReport() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the name for the report:');
  var name = response.getResponseText().trim();

  if (response.getSelectedButton() == ui.Button.CLOSE) {
    return; // Exit if the user closes the prompt
  }

  createReport(ui, name);
}

function createReport(ui, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dateRange = sheet.getRange("A2:A63").getValues();
  var shiftTypeRange = sheet.getRange("C2:C63").getValues();
  var locationHeaders = sheet.getRange("D1:T1").getValues()[0];
  var shiftDataRange = sheet.getRange("D2:R61").getValues();

  var today = new Date();
  var totalHours = 0;
  var totalHoursAll = 0;
  var report = [['DATE', 'SHIFT TYPE', 'LOCATION', 'HOURS']]; // Capitalized headers
  var pastShifts = [];
  var futureShifts = [];

  for (var i = 0; i < shiftDataRange.length; i++) {
    var date = dateRange[i][0];
    if (!date && i > 0) {
      date = dateRange[i - 1][0]; // If the date cell is empty, use the date from the cell above
    }

    for (var j = 0; j < shiftDataRange[0].length; j++) {
      if (shiftDataRange[i][j] === name) {
        var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd/MM/yyyy");
        var shiftType = shiftTypeRange[i][0];
        var location = locationHeaders[j];
        var hours = calculateHours(shiftType, location);
        var shiftDetails = [formattedDate, shiftType, location, hours];
        var hoursValue = parseHours(hours);

        totalHoursAll += hoursValue;

        if (new Date(date) <= today) {
          pastShifts.push(shiftDetails);
          totalHours += hoursValue;
        } else {
          futureShifts.push(shiftDetails);
        }
      }
    }
  }

  // Combine past and future shifts with a demarcation line
  report = report.concat(pastShifts);
  if (pastShifts.length > 0 && futureShifts.length > 0) {
    report.push(['----', '----', '----', '----']); // Demarcation line
  }
  report = report.concat(futureShifts);

  if (report.length > 1) {
    var reportText = report.map(function(row) { return row.join('\t'); }).join('\n');
    var docUrl = createReportDocument(name, reportText, totalHours, totalHoursAll);

    var htmlContent = `
      <style>
        body { font-family: 'Arial', sans-serif; }
        pre { margin: 10px; }
        .button {
          margin: 5px;
          padding: 8px 15px;
          border: none;
          border-radius: 4px;
          background-color: #4CAF50;
          color: white;
          font-family: 'Arial', sans-serif;
        }
        .button:hover {
          background-color: #45a049;
        }
      </style>
      <pre>${reportText}</pre>
      <button class="button" onclick="window.open('${docUrl}', '_blank')">Print Report</button>
      <button class="button" onclick="google.script.host.close()">Close</button>
      <button class="button" onclick="google.script.run.withSuccessHandler(google.script.host.close).generateShiftReport()">Run Again</button>
    `;

    var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(500)
      .setHeight(350);
    ui.showModalDialog(htmlOutput, name + ' Shift Report (' + totalHours.toFixed(2) + ' of ' + totalHoursAll.toFixed(2) + ' Total Hours)');
  } else {
    ui.alert('No shifts found for ' + name);
  }
}

function createReportDocument(name, reportText, totalHours, totalHoursAll) {
  var doc = DocumentApp.create(name + ' Shift Report');
  var body = doc.getBody();

  // Add a header to the document
  var header = doc.addHeader();

  // Replace 'YOUR_DRIVE_FILE_ID' with the file ID of your logo image
  var fileId = '1Y-Cn8cTf9xNcT7ssovUQJPnQYP1lV81P';
  var imageBlob = DriveApp.getFileById(fileId).getBlob();
  
  // Insert the image and adjust its size
  var image = header.appendImage(imageBlob);
  image.setWidth(200); // Set the width in pixels
  image.setHeight(42.4); // Set the height in pixels

  // Optionally, set the alignment of the image
  // image.getParent().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var headerText = name + ' Shift Report (' + totalHours.toFixed(2) + ' of ' + totalHoursAll.toFixed(2) + ' Total Hours)';
  var paragraph = header.appendParagraph(headerText);
  paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.setText(reportText);
  doc.saveAndClose();
  return doc.getUrl();
}

function generateFullReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A2:R61").getValues();
  var names = extractUniqueNames(dataRange);

  var doc = DocumentApp.create('Bethal Manna Apr-May Full Shift Report');

  names.forEach(function(name, index) {
    // Use a visual separator for clarity between sections
    if (index > 0) {
      appendVisualSeparator(doc);
    }
    var { reportData, totalHours } = createReportData(name, dataRange);
    if (reportData.length > 1) {
      var reportText = reportData.map(row => row.join('\t')).join('\n');
      appendReportToDocument(doc, name, reportText, totalHours);
    } else {
      appendTextToDocument(doc, 'No shifts found for ' + name);
    }
  });

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('Full report created: ' + doc.getUrl());
}

function createReportData(name, dataRange) {
  var locationHeaders = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D1:U1").getValues()[0];
  var today = new Date();
  var report = [['DATE', 'SHIFT TYPE', 'LOCATION', 'HOURS']];
  var pastShifts = [];
  var futureShifts = [];

  dataRange.forEach(function(row, i) {
    var date = row[0] || (i > 0 ? dataRange[i - 1][0] : null);
    if (!date) return;

    var shiftType = row[2];
    row.slice(3).forEach(function(cell, j) {
      if (cell === name) {
        var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd/MM/yyyy");
        var location = locationHeaders[j];
        var hours = calculateHours(shiftType, location);
        var shiftDetails = [formattedDate, shiftType, location, hours];

        if (new Date(date) <= today) {
          pastShifts.push(shiftDetails);
        } else {
          futureShifts.push(shiftDetails);
        }
      }
    });
  });

  var totalHours = 0;
  pastShifts.concat(futureShifts).forEach(function(shift) {
    totalHours += parseHours(shift[3]);
  });

  return { reportData: [...report, ...pastShifts, ...(pastShifts.length > 0 && futureShifts.length > 0 ? [['----', '----', '----', '----']] : []), ...futureShifts], totalHours: totalHours };
}

function appendReportToDocument(doc, name, reportText, totalHours) {
  var body = doc.getBody();
  var headerText = name + ' Shift Report (Total Hours: ' + totalHours.toFixed(2) + ')';
  body.appendParagraph(headerText).setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Adjust table or text formatting here if needed
  var table = body.appendTable();
  var fontSize = 8; // Specify your desired font size here

  reportText.split('\n').forEach(function(row, rowIndex) {
    var tableRow = table.appendTableRow();
    row.split('\t').forEach(function(cell) {
      var tableCell = tableRow.appendTableCell(cell);
      tableCell.setFontSize(fontSize);
    });
  });
}



function appendTextToDocument(doc, text) {
  var body = doc.getBody();
  body.appendParagraph(text);
}

function extractUniqueNames(dataRange) {
  var names = new Set();
  dataRange.forEach(row => {
    row.slice(3).forEach(cell => {
      if (cell) names.add(cell);
    });
  });
  return Array.from(names);
}

function calculateHours(shiftType, location) {
  if (shiftType === 'Night' && location === 'Symons Rd (DW)') {
    return '8h';
  } else if (shiftType === 'Day' && location === '115a Northend Rd (AI)') {
    return '4h';
  } else if (shiftType === 'Day' && location === 'Gads Hill Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Kitchener Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Sydney Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Peareswood Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'Symons Mid Shift') {
    return '8h';
  } else if (shiftType === 'Day' && location === 'THOROLD Mid Shift') {
    return '8h';
  } else {
    return '12h';
  }
}

function parseHours(hoursString) {
  var parts = hoursString.split(' ');
  var hours = 0;
  parts.forEach(function(part) {
    if (part.includes('h')) {
      hours += parseInt(part);
    } else if (part.includes('m')) {
      hours += parseInt(part) / 60;
    }
  });
  return hours;
}

function appendVisualSeparator(doc) {
  var body = doc.getBody();
  // Add an empty paragraph as a spacer or create a visual separator with text
  body.appendParagraph("😊😊😊😊😊😊😊😊😊😊😊😊😊😊😊").setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}
function generateLocationReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A2:R61").getValues();
  var names = extractUniqueNames(dataRange);
  var doc = DocumentApp.create('Employee Multiple Location Report');
  
  names.forEach(function(name, index) {
    var locations = extractLocationsForName(name, dataRange);
    // Only process names with more than one unique location
    if (locations.length > 1) {
      if (index > 0) {
        appendVisualSeparator(doc);
      }
      appendLocationsToDocument(doc, name, locations);
    }
  });

  if (doc.getBody().getText().trim() === '') {
    doc.getBody().appendParagraph('No employees found working at multiple locations.');
  } else {
    SpreadsheetApp.getUi().alert('Multiple location report created: ' + doc.getUrl());
  }
}

function extractLocationsForName(name, dataRange) {
  var locationHeaders = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("D1:U1").getValues()[0];
  var locations = new Set();
  
  dataRange.forEach(row => {
    row.slice(3).forEach((cell, j) => {
      if (cell === name) {
        locations.add(locationHeaders[j]);
      }
    });
  });
  
  return Array.from(locations);
}

function appendLocationsToDocument(doc, name, locations) {
  var body = doc.getBody();
  body.appendParagraph(name + ' works at/appears at:').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  locations.forEach(location => {
    body.appendParagraph(location).setHeading(DocumentApp.ParagraphHeading.NORMAL);
  });
}

function queryByTotalHours() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A2:R61").getValues();
  var names = extractUniqueNames(dataRange);

  var doc = DocumentApp.create('Total Hours Query Report');

  names.forEach(function(name, index) {
    if (index > 0) {
      appendVisualSeparator(doc);
    }
    var { reportData, totalHours } = createReportData(name, dataRange);
    // Filter by total hours < 140 or > 156
    if (totalHours < 140 || totalHours > 156) {
      if (reportData.length > 1) {
        var reportText = reportData.map(row => row.join('\t')).join('\n');
        appendReportToDocument(doc, name, reportText, totalHours);
      } else {
        appendTextToDocument(doc, 'No shifts found for ' + name + ' within the specified hours range');
      }
    }
  });

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('Total hours query report created: ' + doc.getUrl());
}

function generateFullReportAlphabetical() {
  generateSortedFullReport((a, b) => a.name.localeCompare(b.name));
}

function generateFullReportLeastHours() {
  generateSortedFullReport((a, b) => a.totalHours - b.totalHours);
}

function generateFullReportMostHours() {
  generateSortedFullReport((a, b) => b.totalHours - a.totalHours);
}

function generateSortedFullReport(sortFunction) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A2:R61").getValues();
  var names = extractUniqueNames(dataRange);

  var reports = names.map(name => {
    var { reportData, totalHours } = createReportData(name, dataRange);
    return { name, reportData, totalHours };
  });

  // Sort reports based on the provided sort function
  reports.sort(sortFunction);

  var doc = DocumentApp.create('Sorted Full Shift Report');

  reports.forEach(function(report, index) {
    if (index > 0) {
      appendVisualSeparator(doc);
    }
    if (report.reportData.length > 1) {
      var reportText = report.reportData.map(row => row.join('\t')).join('\n');
      appendReportToDocument(doc, report.name, reportText, report.totalHours);
    } else {
      appendTextToDocument(doc, 'No shifts found for ' + report.name);
    }
  });

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('Sorted full report created: ' + doc.getUrl());
}
function generateNameAndTotalHoursReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange("A2:R61").getValues();
  var names = extractUniqueNames(dataRange);

  // Sort names in alphabetical order
  names.sort();

  var doc = DocumentApp.create('Name and Total Hours Report (Alphabetical Order)');
  var body = doc.getBody();

  names.forEach(function(name) {
    var { totalHours } = createReportData(name, dataRange);
    body.appendParagraph(name + ' Shift Report (Total Hours: ' + totalHours.toFixed(2) + ')').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    // Optionally, you can add a visual separator between names
    appendVisualSeparator(doc); 
  });

  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('Name and Total Hours Report (Alphabetical Order) created: ' + doc.getUrl());
}
