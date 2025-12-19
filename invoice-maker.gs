function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📜 Invoice')
    .addItem('Show Invoice', 'showInvoice')
    .addToUi();
}

function showInvoice() {
  // Month data template
  var month_data_template = [
    {month: 'January', period_start: '30 January', period_end: '29 February'},
    {month: 'February', period_start: '29 February', period_end: '30 March'},
    {month: 'March', period_start: '30 March', period_end: '30 April'},
    {month: 'April', period_start: '30 April', period_end: '30 May'},
    {month: 'May', period_start: '30 May', period_end: '30 June'},
    {month: 'June', period_start: '30 June', period_end: '30 July'},
    {month: 'July', period_start: '30 July', period_end: '30 August'},
    {month: 'August', period_start: '30 August', period_end: '30 September'},
    {month: 'September', period_start: '30 September', period_end: '30 October'},
    {month: 'October', period_start: '30 October', period_end: '30 November'},
    {month: 'November', period_start: '30 November', period_end: '30 December'},
    {month: 'December', period_start: '30 December', period_end: '30 January'}
  ];
  
  // Get month data from Payments sheet
  var paymentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
  var month = '';
  var periodStart = '';
  var periodEnd = '';
  
  if (paymentsSheet) {
    var paymentsData = paymentsSheet.getDataRange().getValues();
    var paymentsHeaders = paymentsData[0];
    var monthCol = paymentsHeaders.indexOf('Month');
    
    if (monthCol !== -1) {
      // Find last row with data
      var lastRow = -1;
      for (var i = paymentsData.length - 1; i > 0; i--) {
        if (paymentsData[i][monthCol] && paymentsData[i][monthCol] !== '') {
          lastRow = i;
          break;
        }
      }
      
      if (lastRow !== -1) {
        var lastMonth = paymentsData[lastRow][monthCol];
        
        // Find current month in template and get next month
        for (var j = 0; j < month_data_template.length; j++) {
          if (month_data_template[j].month === lastMonth) {
            var nextIndex = (j + 1) % month_data_template.length; // Wrap around for December -> January
            month = month_data_template[nextIndex].month;
            periodStart = month_data_template[nextIndex].period_start;
            periodEnd = month_data_template[nextIndex].period_end;
            break;
          }
        }
        
        // Write next month data to Payments sheet
        var periodCol = paymentsHeaders.indexOf('Period');
        var statusCol = paymentsHeaders.indexOf('Status');
        
        if (month !== '') {
          var nextRowIndex = lastRow + 1; // Next row (0-indexed)
          var nextRowNumber = nextRowIndex + 1; // Convert to 1-indexed for setValue
          
          // Set Month column
          paymentsSheet.getRange(nextRowNumber, monthCol + 1).setValue(month);
          
          // Set Period column if it exists
          if (periodCol !== -1) {
            paymentsSheet.getRange(nextRowNumber, periodCol + 1).setValue(periodStart + ' - ' + periodEnd);
          }
          
          // Set Status column if it exists
          if (statusCol !== -1) {
            paymentsSheet.getRange(nextRowNumber, statusCol + 1).setValue('🔵 Upcoming');
          }
        }
      }
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Report" not found!');
    return;
  }
  
  // Get all data
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indices
  var fullCol = headers.indexOf('Full');
  var devCol = headers.indexOf('Dev');
  var collabCol = headers.indexOf('Collab');
  var teamCol = headers.indexOf('Team');
  var nameCol = headers.indexOf('Name');
  var costCol = headers.indexOf('Cost');
  
  if (fullCol === -1 || devCol === -1 || collabCol === -1 || teamCol === -1 || nameCol === -1 || costCol === -1) {
    SpreadsheetApp.getUi().alert('Required columns not found!');
    return;
  }
  
  // Get prices from row 2 (index 1)
  var price_full = data[1][fullCol];
  var price_dev = data[1][devCol];
  var price_collab = data[1][collabCol];
  
  // Find "Subtotal" row
  var subtotalRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Subtotal') {
      subtotalRow = i;
      break;
    }
  }
  
  if (subtotalRow === -1) {
    SpreadsheetApp.getUi().alert('Row with "Subtotal" not found!');
    return;
  }
  
  // Calculate team_count (row number minus 3, but we need index-based calculation)
  // subtotalRow is 0-indexed, so if it's row 10, subtotalRow = 9
  // team_count = (subtotalRow + 1) - 3 = subtotalRow - 2
  var team_count = subtotalRow - 2;
  
  // Create arrays for teams (i=0 to team_count-1, rows i+3 which is index i+2)
  var team_name = [];
  var count_full = [];
  var count_dev = [];
  var count_collab = [];
  
  for (var i = 0; i < team_count; i++) {
    var rowIndex = i + 2; // row i+3 is index i+2 (since row 1 is index 0)
    if (rowIndex < data.length) {
      team_name[i] = data[rowIndex][nameCol];
      count_full[i] = data[rowIndex][fullCol];
      count_dev[i] = data[rowIndex][devCol];
      count_collab[i] = data[rowIndex][collabCol];
    }
  }
  
  // Get totals from Subtotal row
  var totalcount_full = data[subtotalRow][fullCol];
  var totalcount_dev = data[subtotalRow][devCol];
  var totalcount_collab = data[subtotalRow][collabCol];
  var cost_subtotal = data[subtotalRow][costCol];
  
  // Find "Prorated Costs" row
  var proratedRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Prorated Costs') {
      proratedRow = i;
      break;
    }
  }
  
  var cost_prorated = 0;
  if (proratedRow !== -1) {
    cost_prorated = data[proratedRow][costCol];
  }
  
  // Find "Total" row
  var totalRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][teamCol] === 'Total') {
      totalRow = i;
      break;
    }
  }
  
  var cost_total = 0;
  if (totalRow !== -1) {
    cost_total = data[totalRow][costCol];
  }
  
  // Build invoice text
  var invoiceText = '';
  
  // Add subscription details at the beginning if month data is available
  if (month !== '') {
    invoiceText += 'جزئیات اشتراک فیگما فاکتور ماه ' + month + ' (از ' + periodStart + ' تا ' + periodEnd + ') به شرح زیر است:\n';
  }
  
  invoiceText += 'قیمت هر سیت: فول (<span dir="ltr">$' + price_full + '</span>) — دولوپر (<span dir="ltr">$' + price_dev + '</span>) — کلب (<span dir="ltr">$' + price_collab + '</span>)\n\n';
  
  // Add team information (skip if all counts are zero, otherwise print only non-zero counts)
  for (var i = 0; i < team_name.length; i++) {
    // Skip if all three counts are zero
    if (count_full[i] == 0 && count_dev[i] == 0 && count_collab[i] == 0) {
      continue;
    }
    
    // Build team line with only non-zero counts
    var teamLine = team_name[i] + ': ';
    var parts = [];
    
    if (count_full[i] > 0) {
      parts.push(count_full[i] + ' فول');
    }
    if (count_dev[i] > 0) {
      parts.push(count_dev[i] + ' دولوپر');
    }
    if (count_collab[i] > 0) {
      parts.push(count_collab[i] + ' کلب');
    }
    
    teamLine += parts.join(' - ');
    invoiceText += teamLine + '\n';
  }
  
  invoiceText += '\nکل سیت‌ها: ' + totalcount_full + ' فول - ' + totalcount_dev + ' دولوپر - ' + totalcount_collab + ' کلب\n\n';
  
  // Conditional cost printing
  if (cost_prorated == 0) {
    invoiceText += 'مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  } else if (cost_prorated > 0) {
    invoiceText += 'مجموعا <span dir="ltr">$' + cost_subtotal + '</span> بعلاوه <span dir="ltr">$' + cost_prorated + '</span> هزینه سرشکن ماه قبل، مبلغ نهایی <span dir="ltr">$' + cost_total + '</span>';
  }
  
  // Show modal dialog with download button
  var htmlContent = '<div style="font-family: Vazirmatn, sans-serif; padding: 20px; white-space: pre-wrap; direction: rtl; text-align: right;">' + 
    invoiceText + 
    '</div>' +
    '<div style="padding: 20px; text-align: center;">' +
    '<button onclick="downloadReport(\'' + month + '\')" style="padding: 10px 20px; font-size: 14px; cursor: pointer; background-color: #4285f4; color: white; border: none; border-radius: 4px;">Download Report</button>' +
    '</div>' +
    '<script>' +
    'function downloadReport(month) {' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    if (result && result.data) {' +
    '      var byteCharacters = atob(result.data);' +
    '      var byteNumbers = new Array(byteCharacters.length);' +
    '      for (var i = 0; i < byteCharacters.length; i++) {' +
    '        byteNumbers[i] = byteCharacters.charCodeAt(i);' +
    '      }' +
    '      var byteArray = new Uint8Array(byteNumbers);' +
    '      var blob = new Blob([byteArray], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});' +
    '      var url = URL.createObjectURL(blob);' +
    '      var a = document.createElement("a");' +
    '      a.href = url;' +
    '      a.download = result.filename;' +
    '      document.body.appendChild(a);' +
    '      a.click();' +
    '      document.body.removeChild(a);' +
    '      URL.revokeObjectURL(url);' +
    '    }' +
    '  }).downloadReportFile(month);' +
    '}' +
    '</script>';
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(450),
    'Invoice'
  );
}

function downloadReportFile(month) {
  try {
    var paidUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paid Users');
    
    if (!paidUsersSheet) {
      SpreadsheetApp.getUi().alert('Sheet "Paid Users" not found!');
      return null;
    }
    
    var data = paidUsersSheet.getDataRange().getValues();
    if (data.length === 0) {
      SpreadsheetApp.getUi().alert('No data found in "Paid Users" sheet!');
      return null;
    }
    
    var headers = data[0];
    var columnsToInclude = ['Name', 'Email', 'DK Email', 'Team', 'Department', 'Seat'];
    var columnIndices = [];
    
    // Find column indices
    for (var i = 0; i < columnsToInclude.length; i++) {
      var colIndex = headers.indexOf(columnsToInclude[i]);
      if (colIndex === -1) {
        SpreadsheetApp.getUi().alert('Column "' + columnsToInclude[i] + '" not found!');
        return null;
      }
      columnIndices.push(colIndex);
    }
    
    // Create filtered data with only selected columns
    var filteredData = [];
    filteredData.push(columnsToInclude); // Header row
    
    for (var i = 1; i < data.length; i++) {
      var row = [];
      for (var j = 0; j < columnIndices.length; j++) {
        row.push(data[i][columnIndices[j]]);
      }
      filteredData.push(row);
    }
    
    // Create temporary spreadsheet
    var tempSpreadsheet = SpreadsheetApp.create('Temp_' + new Date().getTime());
    var tempSheet = tempSpreadsheet.getActiveSheet();
    
    // Write filtered data
    tempSheet.getRange(1, 1, filteredData.length, columnsToInclude.length).setValues(filteredData);
    
    // Export as Excel
    var fileName = (month !== '' ? month : 'Report') + '_supernova_invoice.xlsx';
    var blob = tempSpreadsheet.getBlob().setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    
    // Convert blob to base64
    var base64Data = Utilities.base64Encode(blob.getBytes());
    
    // Delete temporary spreadsheet
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    
    // Return base64 data and filename for client-side download
    return {
      data: base64Data,
      filename: fileName
    };
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating report: ' + error.toString());
    return null;
  }
}

