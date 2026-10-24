/**
 * @fileoverview CSV and PDF Report Generation.
 */

function exportShiftManifestCsv(dateStr, shift) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName(DEPLOYMENTS_SHEET);
  if (!dSheet || dSheet.getLastRow() < 2) return { success: false, message: 'No deployment data found.' };

  const targetDate = dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const targetShift = String(shift || 'First Shift').trim();

  const rows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 14).getValues();
  const headers = ['Deployment ID', 'Date', 'Shift', 'Vehicle', 'Rental ID', 'Driver / Tech', 'Zone', 'Goals', 'Target', 'Status', 'Time Out'];
  const outputRows = [headers.join(',')];

  rows.forEach(r => {
    const rowDate = r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    const rowShift = String(r[2] || '').trim();

    if (rowDate === targetDate && rowShift.toLowerCase() === targetShift.toLowerCase()) {
      const line = [
        r[0],
        rowDate,
        rowShift,
        `"${r[3]}"`, `"${r[4]}"`,
        `"${r[5]}"`, `"${r[6]}"`,
        `"${(r[7] \vert{}\vert{} '').toString().replace(/"/g, '""')}"`, `"${(r[8] || '').toString().replace(/"/g, '""')}"`,
        r[9],
        r[10] ? Utilities.formatDate(new Date(r[10]), Session.getScriptTimeZone(), 'hh:mm a') : ''
      ];
      outputRows.push(line.join(','));
    }
  });

  if (outputRows.length === 1) {
    return { success: false, message: `No dispatches recorded for ${targetShift} on${targetDate}.` };
  }

  logAudit('EXPORT_CSV', 'Deployments', `${targetDate}_${targetShift}`, 'Shift manifest exported to CSV');

  return {
    success: true,
    filename: `VEO_Dispatch_${targetDate}_${targetShift.replace(/\s+/g, '_')}.csv`,
    csvContent: outputRows.join('\r\n')
  };
}

function getPrintableShiftManifestHtml(dateStr, shift) {
  const state = getDashboardState();
  const targetDate = dateStr || state.todayDate;
  const targetShift = String(shift || 'First Shift').trim();

  const matchingDeployments = state.deployments.filter(d => {
    return d.date === targetDate && d.shift.toLowerCase() === targetShift.toLowerCase();
  });

  let rowsHtml = '';
  if (!matchingDeployments.length) {
    rowsHtml = `<tr><td colspan="7" style="text-align:center; padding: 24px; color:#64748b;">No dispatches recorded for ${targetShift} on${targetDate}.</td></tr>`;
  } else {
    rowsHtml = matchingDeployments.map((d, idx) => `
      <tr style="border-bottom: 1px solid #e2e8f0; font-size: 12px;">
        <td style="padding: 8px 10px; font-weight:600;">${idx + 1}.${d.vehicle}</td>
        <td style="padding: 8px 10px;">${d.rentalId}</td>
        <td style="padding: 8px 10px; font-weight:600;">${d.tech}</td>
        <td style="padding: 8px 10px;"><span style="background:#e0f2fe; color:#0369a1; padding:2px 6px; border-radius:4px; font-weight:600;">${d.zone}</span></td>
        <td style="padding: 8px 10px;">${d.goals || '—'}</td>
        <td style="padding: 8px 10px;">${d.target || '—'}</td>
        <td style="padding: 8px 10px;">${d.departureTime}</td>
      </tr>
    `).join('');
  }

  logAudit('EXPORT_PDF', 'Deployments', `${targetDate}_${targetShift}`, 'Shift manifest printed/exported to PDF');

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <title>VEO Dispatch Manifest - ${targetShift}</title>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 24px; color: #0f172a; margin: 0; }
      .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #0f172a; padding-bottom: 12px; margin-bottom: 16px; }
      h1 { margin: 0; font-size: 20px; text-transform: uppercase; letter-spacing: 0.05em; }
      .meta { font-size: 13px; color: #475569; margin-top: 4px; }
      .badge { background: #0f172a; color: #fff; padding: 4px 10px; border-radius: 4px; font-size: 13px; font-weight: 700; }
      table { width: 100%; border-collapse: collapse; text-align: left; }
      th { background: #f8fafc; border-bottom: 1px solid #cbd5e1; padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569; }
      .footer { margin-top: 24px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 8px; }
      @media print {
        .no-print { display: none; }
        body { padding: 0; }
      }
    </style>
  </head>
  <body>
    <div class="no-print" style="margin-bottom: 16px; display: flex; gap: 8px;">
      <button onclick="window.print()" style="padding: 8px 16px; background:#2563eb; color:#fff; border:none; border-radius:6px; font-weight:600; cursor:pointer;">🖨️ Print / Save to PDF</button>
      <button onclick="window.close()" style="padding: 8px 16px; background:#f1f5f9; border:1px solid #cbd5e1; border-radius:6px; cursor:pointer;">Close</button>
    </div>

    <div class="header">
      <div>
        <h1>🚚 VEO Fleet Operations - Daily Dispatch Manifest</h1>
        <div class="meta">Date: <strong>${targetDate}</strong></div>
      </div>
      <div>
        <span class="badge">${targetShift}</span>
      </div>
    </div>

    <table>
      <thead>
        <tr>
          <th>Vehicle</th>
          <th>Rental ID</th>
          <th>Driver / Technician</th>
          <th>Work Zone</th>
          <th>Focus Directives & Goals</th>
          <th>Battery Target</th>
          <th>Departure</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>

    <div class="footer">
      <span>Operational continuity confirmed by Operations Lead.</span>
      <span>VEO Fleet Tracking Portal</span>
    </div>

    <script>
      window.onload = function() {
        setTimeout(function() { window.print(); }, 400);
      };
    </script>
  </body>
  </html>
  `;
}