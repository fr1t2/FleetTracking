function getDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const cfg = getSystemConfig(); 

  // 1. Vehicles
  const vSheet = ss.getSheetByName('Vehicles');
  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 18).getValues();
    vRows.forEach(r => {
      const cleanName = String(r[0] || '').trim();
      const rentalId = String(r[1] || '').trim();
      if (!cleanName && !rentalId) return;
      vehicles.push({
        cleanName: cleanName, rentalId: rentalId, agency: String(r[2] || ''), makeModel: String(r[3] || ''), 
        vin: String(r[4] || ''), plate: String(r[5] || ''), 
        plateExpDate: (r[6] instanceof Date) ? Utilities.formatDate(r[6], tz, 'yyyy-MM-dd') : String(r[6] || '—'),
        samsaraSerial: String(r[7] || ''), status: String(r[8] || 'Warehouse (Ready)').trim(),
        pickupDate: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9] || '—'),
        zone: String(r[11] || '—'), driver: String(r[12] || '—'),
        lastService: (r[13] instanceof Date) ? Utilities.formatDate(r[13], tz, 'yyyy-MM-dd') : '—',
        notes: String(r[14] || ''), addons: String(r[17] || '')
      });
    });
  }

  // 2. Deployments
  const dSheet = ss.getSheetByName('Deployments');
  const deployments = [];
  if (dSheet && dSheet.getLastRow() > 1) {
    const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 16).getValues();
    dRows.forEach(r => {
      let depDate = '';
      if (r[1] instanceof Date) depDate = Utilities.formatDate(r[1], tz, 'yyyy-MM-dd');
      else if (r[1]) { const parsed = new Date(r[1]); depDate = !isNaN(parsed) ? Utilities.formatDate(parsed, tz, 'yyyy-MM-dd') : String(r[1]).split(' ')[0]; }

      let timeOut = '—', timeIn = '';
      if (r[10] instanceof Date) timeOut = Utilities.formatDate(r[10], tz, 'hh:mm a');
      else if (r[10]) { const po = new Date(r[10]); timeOut = !isNaN(po) ? Utilities.formatDate(po, tz, 'hh:mm a') : String(r[10]); }
      
      if (r[11] instanceof Date) timeIn = Utilities.formatDate(r[11], tz, 'hh:mm a');
      else if (r[11]) { const pi = new Date(r[11]); timeIn = !isNaN(pi) ? Utilities.formatDate(pi, tz, 'hh:mm a') : String(r[11]); }

      deployments.push({
        id: String(r[0] || ''), date: depDate, shift: String(r[2] || '').trim(), vehicle: String(r[3] || ''), 
        rentalId: String(r[4] || ''), tech: String(r[5] || ''), zone: String(r[6] || ''), goals: String(r[7] || ''),
        target: String(r[8] || ''), status: String(r[9] || 'Dispatched'), departureTime: timeOut, returnTime: timeIn,
        coDriver: String(r[14] || ''), isTraining: String(r[15] || '') === 'Yes'
      });
    });
  }

  // 3. Incidents (Now pulls Driver Involved from Col 14 if it exists)
  const iSheet = ss.getSheetByName('Incidents_Service');
  const incidents = [];
  if (iSheet && iSheet.getLastRow() > 1) {
    const maxCol = iSheet.getLastColumn();
    const iRows = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, maxCol < 14 ? 14 : maxCol).getValues();
    iRows.forEach(r => {
      incidents.push({
        id: String(r[0] || ''), dateReported: (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        vehicle: String(r[2] || ''), rentalId: String(r[3] || ''), category: String(r[4] || ''), details: String(r[5] || ''), 
        impact: String(r[7] || ''), resolution: String(r[8] || ''), dateResolved: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9]),
        driverInvolved: String(r[13] || '') // Column N
      });
    });
  }

  // 4. Tech Profiles (NEW)
  const tSheet = ss.getSheetByName('Tech_Profiles');
  const techs = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    const tRows = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 11).getValues();
    tRows.forEach(r => {
      if(r[0]) {
        techs.push({
          name: String(r[0]).trim(), phone: String(r[1] || ''), shift: String(r[2] || 'First Shift'), 
          startTime: String(r[3] || ''), endTime: String(r[4] || ''), workDays: String(r[5] || 'Mon-Fri'), 
          passedTest: String(r[6] || 'No') === 'Yes', goFast: String(r[7] || 'No') === 'Yes', 
          restricted: String(r[8] || 'No') === 'Yes', status: String(r[9] || 'Available'), notes: String(r[10] || '')
        });
      }
    });
  }

  return { config: cfg, vehicles: vehicles, deployments: deployments, incidents: incidents, techs: techs };
}