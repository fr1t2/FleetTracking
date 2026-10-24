/**
 * @fileoverview Export Engine for CSV and PDF Reports.
 */

/**
 * Escapes values for safe CSV encoding.
 */
function escapeCsvValue(val) {
  if (val === null || val === undefined) return '""';
  var str = String(val);
  return '"' + str.split('"').join('""') + '"';
}

/**
 * Generates CSV string for the selected shift's dispatch ledger.
 */
function exportShiftManifestCsv(dateStr, shift) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dSheet = ss.getSheetByName(DEPLOYMENTS_SHEET);
  if (!dSheet || dSheet.getLastRow() < 2) {
    return { success: false, message: 'No deployment data found.' };
  }

  var tz = Session.getScriptTimeZone();
  var targetDate = dateStr || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var targetShift = shift || 'First Shift';

  var rows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 14).getValues();
  var headers = ['Deployment ID', 'Date', 'Shift', 'Vehicle', 'Rental ID', 'Driver / Tech', 'Zone', 'Goals', 'Target', 'Status', 'Time Out'];
  var outputRows = [headers.join(',')];

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowDate = '';
    if (r[1] instanceof Date) {
      rowDate = Utilities.formatDate(r[1], tz, 'yyyy-MM-dd');
    } else if (r[1]) {
      var p = new Date(r[1]);
      rowDate = !isNaN(p) ? Utilities.formatDate(p, tz, 'yyyy-MM-dd') : String(r[1]);
    }

    var rowShift = String(r[2] || '').trim();

    if (rowDate === targetDate && rowShift.toLowerCase() === targetShift.toLowerCase()) {
      var timeOut = '';
      if (r[10] instanceof Date) {
        timeOut = Utilities.formatDate(r[10], tz, 'hh:mm a');
      } else if (r[10]) {
        timeOut = String(r[10]);
      }

      var line = [
        escapeCsvValue(r[0]),
        escapeCsvValue(rowDate),
        escapeCsvValue(rowShift),
        escapeCsvValue(r[3]),
        escapeCsvValue(r[4]),
        escapeCsvValue(r[5]),
        escapeCsvValue(r[6]),
        escapeCsvValue(r[7]),
        escapeCsvValue(r[8]),
        escapeCsvValue(r[9]),
        escapeCsvValue(timeOut)
      ];
      outputRows.push(line.join(','));
    }
  }

  if (outputRows.length === 1) {
    return { success: false, message: 'No dispatches recorded for ' + targetShift + ' on ' + targetDate + '.' };
  }

  logAudit('EXPORT_CSV', 'Deployments', targetDate + '_' + targetShift, 'Shift manifest exported to CSV');

  var cleanShiftName = targetShift.split(' ').join('_');
  return {
    success: true,
    filename: 'VEO_Dispatch_' + targetDate + '_' + cleanShiftName + '.csv',
    csvContent: outputRows.join('\r\n')
  };
}

/**
 * Generates CSV string of the full fleet registry.
 */
function exportFleetRosterCsv() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vSheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!vSheet || vSheet.getLastRow() < 1) {
    return { success: false, message: 'Vehicles sheet is empty.' };
  }

  var tz = Session.getScriptTimeZone();
  var data = vSheet.getRange(1, 1, vSheet.getLastRow(), 17).getValues();
  var outputRows = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var escapedRow = [];
    for (var j = 0; j < row.length; j++) {
      var cell = row[j];
      var str = '';
      if (cell instanceof Date) {
        str = Utilities.formatDate(cell, tz, 'yyyy-MM-dd HH:mm');
      } else {
        str = String(cell || '');
      }
      escapedRow.push(escapeCsvValue(str));
    }
    outputRows.push(escapedRow.join(','));
  }

  logAudit('EXPORT_CSV', 'Vehicles', 'FULL_ROSTER', 'Full fleet roster exported to CSV');

  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return {
    success: true,
    filename: 'VEO_Fleet_Master_' + todayStr + '.csv',
    csvContent: outputRows.join('\r\n')
  };
}

/**
 * Generates a clean printable HTML document for printing or saving directly to PDF.
 */
function getPrintableShiftManifestHtml(dateStr, shift) {
  var state = getDashboardState();
  var cfg = state.config;
  var marketName = (cfg.settings && cfg.settings['Market_Name']) ? cfg.settings['Market_Name'] : 'VEO Fleet Market';
  var targetDate = dateStr || state.todayDate;
  var targetShift = shift || 'First Shift';

  // Filter deployments for this date and shift
  var filtered = [];
  if (state.deployments && state.deployments.length) {
    for (var k = 0; k < state.deployments.length; k++) {
      var d = state.deployments[k];
      if (d.date === targetDate && d.shift.toLowerCase() === targetShift.toLowerCase()) {
        filtered.push(d);
      }
    }
  }

  var rowsHtml = '';
  if (!filtered.length) {
    rowsHtml = '<tr><td colspan="7" style="text-align:center; padding: 24px; color:#64748b;">No dispatches recorded for this shift.</td></tr>';
  } else {
    for (var m = 0; m < filtered.length; m++) {
      var item = filtered[m];
      rowsHtml += '<tr style="border-bottom: 1px solid #e2e8f0; font-size: 12px;">' +
        '<td style="padding: 8px 10px; font-weight:600;">' + (m + 1) + '. ' + item.vehicle + '</td>' +
        '<td style="padding: 8px 10px;">' + item.rentalId + '</td>' +
        '<td style="padding: 8px 10px; font-weight:600;">' + item.tech + '</td>' +
        '<td style="padding: 8px 10px;"><span style="background:#e0f2fe; color:#0369a1; padding:2px 6px; border-radius:4px; font-weight:600;">' + item.zone + '</span></td>' +
        '<td style="padding: 8px 10px;">' + (item.goals || '—') + '</td>' +
        '<td style="padding: 8px 10px;">' + (item.target || '—') + '</td>' +
        '<td style="padding: 8px 10px;">' + item.departureTime + '</td>' +
        '</tr>';
    }
  }

  logAudit('EXPORT_PDF', 'Deployments', targetDate + '_' + targetShift, 'Shift manifest printed/exported to PDF');

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<title>VEO Dispatch Manifest - ' + targetShift + '</title>' +
    '<style>' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 24px; color: #0f172a; margin: 0; }' +
    '.header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #0f172a; padding-bottom: 12px; margin-bottom: 16px; }' +
    'h1 { margin: 0; font-size: 20px; text-transform: uppercase; letter-spacing: 0.05em; }' +
    '.meta { font-size: 13px; color: #475569; margin-top: 4px; }' +
    '.badge { background: #0f172a; color: #fff; padding: 4px 10px; border-radius: 4px; font-size: 13px; font-weight: 700; }' +
    'table { width: 100%; border-collapse: collapse; text-align: left; }' +
    'th { background: #f8fafc; border-bottom: 1px solid #cbd5e1; padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569; }' +
    '.footer { margin-top: 24px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 8px; }' +
    '@media print { .no-print { display: none; } body { padding: 0; } }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="no-print" style="margin-bottom: 16px; display: flex; gap: 8px;">' +
    '<button onclick="window.print()" style="padding: 8px 16px; background:#2563eb; color:#fff; border:none; border-radius:6px; font-weight:600; cursor:pointer;">🖨️ Print / Save to PDF</button>' +
    '<button onclick="window.close()" style="padding: 8px 16px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:6px; cursor:pointer;">Close</button>' +
    '</div>' +
    '<div class="header">' +
    '<div>' +
    '<h1>🚚 ' + marketName + ' - Daily Dispatch Manifest</h1>' +
    '<div class="meta">Date: <strong>' + targetDate + '</strong> &bull; Generated: ' + nowStr + '</div>' +
    '</div>' +
    '<div><span class="badge">' + targetShift + '</span></div>' +
    '</div>' +
    '<table>' +
    '<thead>' +
    '<tr>' +
    '<th>Vehicle</th><th>Rental ID</th><th>Driver / Technician</th><th>Work Zone</th><th>Focus Directives & Goals</th><th>Battery Target</th><th>Departure</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' + rowsHtml + '</tbody>' +
    '</table>' +
    '<div class="footer">' +
    '<span>Operational continuity confirmed by Operations Supervisor / Shift Lead.</span>' +
    '<span>VEO Fleet Tracking Portal</span>' +
    '</div>' +
    '<script>' +
    'window.onload = function() { setTimeout(function() { window.print(); }, 400); };' +
    '</script>' +
    '</body>' +
    '</html>';
}