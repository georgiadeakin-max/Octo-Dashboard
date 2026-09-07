const SHEET_ID = '1l0RPipVIIHJBhyRyadqhR2XrqzPRin4vivfVLvuJqZI';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Octopus Energy Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🐙 Delivery Tools')
    .addItem('Show Dashboard & Map', 'showMap')
    .addToUi();
}

function showMap() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, "Octopus Energy Dashboard");
}

// 🧠 Smart Math Parser for Pick Progress
function parsePickProgress(rawVal) {
  if (rawVal === null || rawVal === undefined || rawVal === "") return 0;
  
  if (typeof rawVal === 'number') {
    if (rawVal > 0 && rawVal <= 1) return rawVal * 100;
    return rawVal;
  }
  
  let strVal = String(rawVal).toLowerCase().trim();
  if (strVal === 'done' || strVal === 'complete' || strVal === 'yes' || strVal === 'true') return 100;
  
  let cleaned = strVal.replace(/[^0-9.]/g, '');
  if (cleaned === '') return 0;
  
  let num = parseFloat(cleaned);
  if (isNaN(num)) return 0;
  
  return num;
}

function getAppData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const sheet = ss.getSheetByName('Master');
  const data = sheet ? sheet.getDataRange().getValues() : [];

  const cancelSheet = ss.getSheetByName('CANCELLED');
  const cancelData = cancelSheet ? cancelSheet.getDataRange().getValues() : [];

  if (data.length <= 1) return null;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  
  // 🔍 Fuzzy Finders to prevent column naming/spacing issues
  let colDateIdx = headers.findIndex(h => h.includes('collection date') || h.includes('col date'));
  let pickProgIdx = headers.findIndex(h => h.includes('pick progress'));
  if (pickProgIdx === -1) pickProgIdx = headers.findIndex(h => h.includes('progress'));
  let dispatchIdx = headers.findIndex(h => h.includes('dispatch'));
  let originSiteIdx = headers.findIndex(h => h.includes('origin site') || h.includes('site')); 
  let vehicleIdx = headers.findIndex(h => h.includes('collection vehicle') || h.includes('vehicle'));
  let deliveryDateIdx = headers.findIndex(h => h.includes('delivery date') || h.includes('del date'));
  let postcodeIdx = headers.findIndex(h => h.includes('postal code') || h.includes('postcode'));
  let gxoOctoIdx = headers.findIndex(h => h.replace(/\s/g, '').includes('gxo/octo'));
  let pipelineIdx = headers.findIndex(h => h.includes('pipeline'));
  let preDelIdx = headers.findIndex(h => h.replace(/[-\s_]/g, '').includes('predelivery'));

  let cDateIdx = -1, cLocIdx = -1, cPipeIdx = -1;
  if (cancelData.length > 0) {
    const cHeaders = cancelData[0].map(h => String(h).trim().toLowerCase());
    cDateIdx = cHeaders.findIndex(h => h.includes('collection date'));
    if (cDateIdx === -1) cDateIdx = cHeaders.findIndex(h => h.includes('cancellation date'));
    if (cDateIdx === -1) cDateIdx = cHeaders.findIndex(h => h.includes('date'));
    cLocIdx = cHeaders.findIndex(h => h.includes('location of stock'));
    cPipeIdx = cHeaders.findIndex(h => h.includes('pipeline'));
  }

  // Define Dates
  const today = new Date();
  today.setHours(0,0,0,0);
  
  let daysToMonday = (1 - today.getDay() + 7) % 7; 
  if (today.getDay() === 0) daysToMonday = -6;
  else if (today.getDay() > 1) daysToMonday = 1 - today.getDay();

  if (today.getDay() === 5) daysToMonday = 3;       
  else if (today.getDay() === 6) daysToMonday = 2;  

  const targetMonday = new Date(today);
  targetMonday.setDate(today.getDate() + daysToMonday);

  const pad = (n) => n < 10 ? '0' + n : n;
  
  const mapTargetDate = new Date(today);
  let mapDaysToAdd = 1; 
  if (today.getDay() === 5) mapDaysToAdd = 3; 
  else if (today.getDay() === 6) mapDaysToAdd = 2; 
  mapTargetDate.setDate(mapTargetDate.getDate() + mapDaysToAdd);
  mapTargetDate.setHours(0, 0, 0, 0);

  function processPeriod(startDate, daysCount) {
    let endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + daysCount);
    
    let titleString = `${pad(startDate.getDate())}/${pad(startDate.getMonth()+1)}/${startDate.getFullYear()}`;

    const initStats = () => ({ totalOrders: 0, notStarted: 0, inProgress: 0, complete: 0, dispatched: 0, pipeline: {} });
    const initDelStats = () => ({ total: 0, octopus: 0, gxo: 0, other: 0 });
    
    const initOrdersAgg = () => [
      { day: 'Mon', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Tue', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Wed', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Thu', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Fri', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Sat', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 },
      { day: 'Sun', notStarted: 0, inProgress: 0, complete: 0, dispatched: 0 }
    ];

    const initPreDelStats = () => {
      let pds = { totalPreDel: 0, totalStandard: 0, daily: [] };
      for (let i = 0; i < daysCount; i++) {
        let d = new Date(startDate);
        d.setDate(startDate.getDate() + i);
        let dayStr = d.toLocaleDateString('en-GB', {weekday:'short'});
        pds.daily.push({ day: dayStr, solarTotal: 0, solarPreDel: 0, hpTotal: 0, hpPreDel: 0 });
      }
      return pds;
    };

    let stats = initStats(), hinckleyStats = initStats(), sheffieldStats = initStats();
    let delCombined = initDelStats(), delHinckley = initDelStats(), delSheffield = initDelStats(), delSlough = initDelStats();
    let vehicleBreakdown = {};
    let cancelStats = { total: 0, daily: [], reasons: {} };
    
    let preDelStats = initPreDelStats();
    let preDelHinckley = initPreDelStats();
    let preDelSheffield = initPreDelStats();

    let dailyStats = [], hinckleyDaily = [], sheffieldDaily = [];
    let delDaily = [], delHDaily = [], delSDaily = [], delSloughDaily = [];
    
    let ordersAggDaily = initOrdersAgg();
    let ordersAggHinckley = initOrdersAgg();
    let ordersAggSheffield = initOrdersAgg();

    for (let i = 0; i < daysCount; i++) {
      let d = new Date(startDate);
      d.setDate(startDate.getDate() + i);
      let dayStr = d.toLocaleDateString('en-GB', {weekday:'short'});
      
      let baseDaily = () => ({ day: dayStr, notStarted: 0, inProgress: 0, complete: 0, dispatched: 0, pipeline: {} });
      let baseDelDaily = () => ({ day: dayStr, octopus: 0, gxo: 0, other: 0 });

      dailyStats.push(baseDaily()); hinckleyDaily.push(baseDaily()); sheffieldDaily.push(baseDaily());
      delDaily.push(baseDelDaily()); delHDaily.push(baseDelDaily()); delSDaily.push(baseDelDaily()); delSloughDaily.push(baseDelDaily());
      cancelStats.daily.push(0);
    }

    for (let i = 1; i < data.length; i++) {
      let pipelineVal = "Unassigned / Blank";
      if (pipelineIdx !== -1 && data[i][pipelineIdx] !== undefined && data[i][pipelineIdx] !== "") {
        pipelineVal = String(data[i][pipelineIdx]).trim();
      }

      // --- EXCLUSION: Skip FSO Request completely ---
      let pipeUp = pipelineVal.toUpperCase();
      if (pipeUp.includes('FSO REQUEST')) continue;

      let rawVeh = data[i][vehicleIdx] ? String(data[i][vehicleIdx]).trim() : "";
      let rawUp = rawVeh.toUpperCase();
      let isSlough = rawUp.includes('TR-SLO') || rawUp.includes('SLOU');
      let site = String(data[i][originSiteIdx]).trim().toLowerCase();
      let isHinckley = !isSlough && site.includes('hinckley');
      let isSheffield = !isSlough && site.includes('sheffield');
      
      let isSolar = pipeUp.startsWith('SOLAR') || pipeUp.includes('SOLAR-') || pipeUp.includes('SOLAR -');
      let isHP = pipeUp.startsWith('HP') || pipeUp.includes('HP-') || pipeUp.includes('HP -');

      // --- 1. Process Orders & Collection based on Collection Date ---
      let colDateVal = data[i][colDateIdx];
      if (colDateVal) {
        let colDate = new Date(colDateVal);
        colDate.setHours(0,0,0,0);
        if (colDate >= startDate && colDate < endDate) {
          let dayIdx = Math.round((colDate.getTime() - startDate.getTime()) / 86400000);
          let weekdayIdx = (colDate.getDay() + 6) % 7; 
          
          let progressVal = pickProgIdx !== -1 ? parsePickProgress(data[i][pickProgIdx]) : 0;
          let disp = dispatchIdx !== -1 ? String(data[i][dispatchIdx]).trim().toLowerCase() : "";
          let isDispatched = (disp === 'true' || disp === 'yes' || disp === 'done');

          let status = 'notStarted';
          if (isDispatched) {
            status = 'dispatched';
          } else if (progressVal >= 100) {
            status = 'complete';
          } else if (progressVal > 0 && progressVal < 100) {
            status = 'inProgress';
          } else {
            status = 'notStarted';
          }

          stats.totalOrders++; stats[status]++; dailyStats[dayIdx][status]++;
          stats.pipeline[pipelineVal] = (stats.pipeline[pipelineVal] || 0) + 1;
          ordersAggDaily[weekdayIdx][status]++;
          
          if (isSlough || isHinckley) { 
            hinckleyStats.totalOrders++; hinckleyStats[status]++; hinckleyDaily[dayIdx][status]++; 
            hinckleyStats.pipeline[pipelineVal] = (hinckleyStats.pipeline[pipelineVal] || 0) + 1;
            ordersAggHinckley[weekdayIdx][status]++;
          }
          else if (isSheffield) { 
            sheffieldStats.totalOrders++; sheffieldStats[status]++; sheffieldDaily[dayIdx][status]++; 
            sheffieldStats.pipeline[pipelineVal] = (sheffieldStats.pipeline[pipelineVal] || 0) + 1;
            ordersAggSheffield[weekdayIdx][status]++;
          }
        }
      }

      // --- 2. Process Delivery Splits AND Pre-Delivery based on Delivery Date ---
      let delDateVal = data[i][deliveryDateIdx];
      if (delDateVal) {
        let delDate = new Date(delDateVal);
        delDate.setHours(0,0,0,0);
        if (delDate >= startDate && delDate < endDate) {
          let dayIdx = Math.round((delDate.getTime() - startDate.getTime()) / 86400000);

          let carrier = 'other';
          let specificType = rawVeh === "" ? "Unassigned / Blank" : rawVeh;

          if (isSlough) {
            let goVal = (gxoOctoIdx !== -1 && data[i][gxoOctoIdx]) ? String(data[i][gxoOctoIdx]).trim().toUpperCase() : "";
            if (goVal.includes('OCTO')) carrier = 'octopus'; else if (goVal.includes('GXO')) carrier = 'gxo';
            specificType = rawVeh;
          } else if (rawUp.startsWith('OCTO')) {
            carrier = 'octopus'; specificType = 'Octopus';
          } else if (rawUp.startsWith('TR-') || rawUp.startsWith('GXO')) {
            carrier = 'gxo';
            if (rawUp.includes('LUTON')) specificType = 'GXO Luton'; else if (rawUp.startsWith('TR-')) specificType = rawVeh; else specificType = 'GXO Other';
          }

          delCombined.total++; delCombined[carrier]++; delDaily[dayIdx][carrier]++;
          if (isSlough) { delSlough.total++; delSlough[carrier]++; delSloughDaily[dayIdx][carrier]++; } 
          else if (isHinckley) { delHinckley.total++; delHinckley[carrier]++; delHDaily[dayIdx][carrier]++; } 
          else if (isSheffield) { delSheffield.total++; delSheffield[carrier]++; delSDaily[dayIdx][carrier]++; }

          vehicleBreakdown[specificType] = (vehicleBreakdown[specificType] || 0) + 1;

          // Pre-Delivery Tracking on Delivery Date
          let isPreDelivery = false;
          if (preDelIdx !== -1) {
             let pdVal = String(data[i][preDelIdx]).trim().toLowerCase();
             isPreDelivery = (pdVal === 'true' || pdVal === 'yes');
          }
          
          if (isPreDelivery) preDelStats.totalPreDel++; else preDelStats.totalStandard++;
          if (isSolar) { preDelStats.daily[dayIdx].solarTotal++; if (isPreDelivery) preDelStats.daily[dayIdx].solarPreDel++; }
          if (isHP) { preDelStats.daily[dayIdx].hpTotal++; if (isPreDelivery) preDelStats.daily[dayIdx].hpPreDel++; }

          if (isSlough || isHinckley) { 
            if (isPreDelivery) preDelHinckley.totalPreDel++; else preDelHinckley.totalStandard++;
            if (isSolar) { preDelHinckley.daily[dayIdx].solarTotal++; if (isPreDelivery) preDelHinckley.daily[dayIdx].solarPreDel++; }
            if (isHP) { preDelHinckley.daily[dayIdx].hpTotal++; if (isPreDelivery) preDelHinckley.daily[dayIdx].hpPreDel++; }
          }
          else if (isSheffield) { 
            if (isPreDelivery) preDelSheffield.totalPreDel++; else preDelSheffield.totalStandard++;
            if (isSolar) { preDelSheffield.daily[dayIdx].solarTotal++; if (isPreDelivery) preDelSheffield.daily[dayIdx].solarPreDel++; }
            if (isHP) { preDelSheffield.daily[dayIdx].hpTotal++; if (isPreDelivery) preDelSheffield.daily[dayIdx].hpPreDel++; }
          }
        }
      }
    }

    // Process Cancellations (Collection Date)
    if (cDateIdx !== -1) {
      for (let i = 1; i < cancelData.length; i++) {
        // --- EXCLUSION: Skip FSO Request on Cancellations ---
        if (cPipeIdx !== -1 && cancelData[i][cPipeIdx] !== undefined) {
          let cPipeVal = String(cancelData[i][cPipeIdx]).trim().toUpperCase();
          if (cPipeVal.includes('FSO REQUEST')) continue;
        }

        let cDateVal = cancelData[i][cDateIdx];
        if (cDateVal) {
          let cDate = new Date(cDateVal);
          cDate.setHours(0,0,0,0);
          if (cDate >= startDate && cDate < endDate) {
            let dayIdx = Math.round((cDate.getTime() - startDate.getTime()) / 86400000);
            cancelStats.total++;
            cancelStats.daily[dayIdx]++;
            let reason = "Unspecified / Blank";
            if (cLocIdx !== -1 && cancelData[i][cLocIdx] !== undefined && cancelData[i][cLocIdx] !== "") reason = String(cancelData[i][cLocIdx]).trim();
            cancelStats.reasons[reason] = (cancelStats.reasons[reason] || 0) + 1;
          }
        }
      }
    }

    const sortedBreakdown = Object.keys(vehicleBreakdown).sort().reduce((obj, key) => { obj[key] = vehicleBreakdown[key]; return obj; }, {});

    return { 
      wcString: titleString, stats, dailyStats, hinckleyStats, hinckleyDaily, sheffieldStats, sheffieldDaily,
      ordersAggDaily, ordersAggHinckley, ordersAggSheffield,
      delCombined, delDaily, delHinckley, delHDaily, delSheffield, delSDaily, delSlough, delSloughDaily,
      vehicleBreakdown: sortedBreakdown, cancelStats: cancelStats, 
      preDelStats: preDelStats, preDelHinckley: preDelHinckley, preDelSheffield: preDelSheffield
    };
  }

  // Delivery Map Generation (Strict Date Match)
  const deliveries = [];
  for (let i = 1; i < data.length; i++) {
    // --- EXCLUSION: Skip FSO Request on Delivery Map ---
    let pipelineVal = "Unassigned / Blank";
    if (pipelineIdx !== -1 && data[i][pipelineIdx] !== undefined && data[i][pipelineIdx] !== "") {
      pipelineVal = String(data[i][pipelineIdx]).trim();
    }
    if (pipelineVal.toUpperCase().includes('FSO REQUEST')) continue;

    let delDateVal = data[i][deliveryDateIdx];
    if (delDateVal) {
      let rowDate = new Date(delDateVal);
      if (rowDate.getFullYear() === mapTargetDate.getFullYear() &&
          rowDate.getMonth() === mapTargetDate.getMonth() &&
          rowDate.getDate() === mapTargetDate.getDate()) {
          
        let pc = data[i][postcodeIdx];
        let rawVeh = data[i][vehicleIdx] ? String(data[i][vehicleIdx]).trim().toUpperCase() : "";
        let isSlough = rawVeh.includes('TR-SLO') || rawVeh.includes('SLOU');
        let carrierName = "Unknown";
        if (isSlough) {
          let goVal = (gxoOctoIdx !== -1 && data[i][gxoOctoIdx]) ? String(data[i][gxoOctoIdx]).trim().toUpperCase() : "";
          if (goVal.includes('OCTO')) carrierName = 'Octopus'; else if (goVal.includes('GXO')) carrierName = 'GXO';
        } else if (rawVeh.startsWith('OCTO') || rawVeh.startsWith('TR-SLO') || rawVeh.includes('SLOU')) {
          carrierName = 'Octopus';
        } else if (rawVeh.startsWith('TR-') || rawVeh.startsWith('GXO')) {
          carrierName = 'GXO';
        }
        if (pc) deliveries.push({ postcode: pc.toString().trim(), vehicle: carrierName });
      }
    }
  }

  // ONLY return the 1 week dashboard data now
  return { 
    dashboard1W: processPeriod(targetMonday, 7),
    map: { deliveries: deliveries, dateString: mapTargetDate.toDateString() }
  };
}
