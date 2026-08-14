/***********************************************************************
 * SHREE JI APARTMENT — MAINTENANCE MANAGEMENT APP
 * Google Apps Script Backend (connects the web app to this Google Sheet)
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Delete any starter code, paste this entire file in.
 * 3. Click the "Run" button once on the function "setupSheets" (select it
 *    from the dropdown next to Run first). Approve the permissions asked.
 *    This creates all 5 sheets, seeds your 11 flats, sets up ₹1000/month,
 *    generates this month's entries, and installs the auto-monthly trigger.
 * 4. Click Deploy -> New deployment -> select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize again if asked.
 * 5. Copy the "Web app URL" it gives you (ends in /exec).
 * 6. Paste that URL into the app's Settings tab -> "Backend URL" field.
 * That's it — the web app now reads/writes this Sheet directly.
 ***********************************************************************/

const SHEET = {
  FLATS: 'Flats_Master',
  PAYMENTS: 'Maintenance_Payments',
  EXPENSES: 'Expenses',
  FUND: 'Fund_Statement',
  DASHBOARD: 'Dashboard_Data'
};

const MONTH_NAMES = ["January","February","March","April","May","June","July",
  "August","September","October","November","December"];

function getSS(){ return SpreadsheetApp.getActiveSpreadsheet(); }

/* ================= ONE-TIME SETUP ================= */
function setupSheets(){
  const ss = getSS();

  // Flats_Master
  let sh = ss.getSheetByName(SHEET.FLATS) || ss.insertSheet(SHEET.FLATS);
  sh.clear();
  sh.appendRow(['Flat No','Owner Name','Status','Monthly Maintenance','Joining Date','Mobile Number','Remarks']);
  sh.setFrozenRows(1);
  const flats = [
    ['Flat 1','Khursheed Ahmad','Active',1000,'','',''],
    ['Flat 2','Gaytri Devi','Active',1000,'','',''],
    ['Flat 3','Rakhi','Active',1000,'','',''],
    ['Flat 4','Manoj Raghav','Active',1000,'','',''],
    ['Flat 5','Nitin','Active',1000,'','',''],
    ['Flat 6','Naveen','Active',1000,'','',''],
    ['Flat 7','Punit','Active',1000,'','',''],
    ['Flat 8','Kapil','Active',1000,'','',''],
    ['Flat 9','Manoj Kumar','Active',1000,'','',''],
    ['Flat 10','Renu Yadav','Active',1000,'','',''],
    ['Flat 11','','Vacant',1000,'','','']
  ];
  flats.forEach(r=>sh.appendRow(r));

  // Maintenance_Payments
  sh = ss.getSheetByName(SHEET.PAYMENTS) || ss.insertSheet(SHEET.PAYMENTS);
  sh.clear();
  sh.appendRow(['Month','Year','Flat No','Owner Name','Amount','Payment Status','Payment Date','Payment Mode','Remarks','Record ID']);
  sh.setFrozenRows(1);

  // Expenses
  sh = ss.getSheetByName(SHEET.EXPENSES) || ss.insertSheet(SHEET.EXPENSES);
  sh.clear();
  sh.appendRow(['Expense Date','Category','Description','Amount','Added By','Record ID']);
  sh.setFrozenRows(1);

  // Fund_Statement
  sh = ss.getSheetByName(SHEET.FUND) || ss.insertSheet(SHEET.FUND);
  sh.clear();
  sh.appendRow(['Total Collection Till Date','Total Expenses Till Date','Current Available Fund']);
  sh.appendRow([0,0,0]);

  // Dashboard_Data
  sh = ss.getSheetByName(SHEET.DASHBOARD) || ss.insertSheet(SHEET.DASHBOARD);
  sh.clear();
  sh.appendRow(['Total Fund','Available Balance','Current Month Collection','Current Month Expense','Last Month Expense','Previous Month Expense','Pending Payments']);
  sh.appendRow([0,0,0,0,0,0,0]);

  // default config
  const props = PropertiesService.getScriptProperties();
  if(!props.getProperty('buildingName')) props.setProperty('buildingName','Shree Ji Apartment');
  if(!props.getProperty('dueDay')) props.setProperty('dueDay','10');
  if(!props.getProperty('defaultAmount')) props.setProperty('defaultAmount','1000');

  generateMonthlyEntries();
  refreshComputedSheets();
  createMonthlyTrigger();

  SpreadsheetApp.getUi().alert('Setup complete! Sheets created, 11 flats loaded, this month\'s entries generated, and the auto-monthly trigger is installed.');
}

/* Installs a trigger that runs generateMonthlyEntries automatically on the 1st of every month */
function createMonthlyTrigger(){
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'generateMonthlyEntries');
  if(!exists){
    ScriptApp.newTrigger('generateMonthlyEntries')
      .timeBased()
      .onMonthDay(1)
      .atHour(1)
      .create();
  }
}

/* ================= AUTO MONTHLY ENTRY GENERATION ================= */
function generateMonthlyEntries(){
  const ss = getSS();
  const flatsSh = ss.getSheetByName(SHEET.FLATS);
  const paySh = ss.getSheetByName(SHEET.PAYMENTS);
  const flats = sheetToObjects(flatsSh);
  const payments = sheetToObjects(paySh);

  const now = new Date();
  const month = MONTH_NAMES[now.getMonth()];
  const year = now.getFullYear();

  flats.filter(f => f['Status'] === 'Active').forEach(f => {
    const already = payments.some(p => p['Month']===month && String(p['Year'])===String(year) && p['Flat No']===f['Flat No']);
    if(!already){
      paySh.appendRow([month, year, f['Flat No'], f['Owner Name'], f['Monthly Maintenance'], 'Pending', '', '', '', Utilities.getUuid()]);
    }
  });
  refreshComputedSheets();
}

/* ================= HELPERS ================= */
function sheetToObjects(sh){
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row=>{
    const obj = {};
    headers.forEach((h,i)=> obj[h] = row[i]);
    return obj;
  });
}
function findRowIndexById(sh, idColName, id){
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(idColName);
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol]) === String(id)) return i+1; // 1-based sheet row
  }
  return -1;
}
function colIndex(sh, name){
  const headers = sh.getDataRange().getValues()[0];
  return headers.indexOf(name) + 1; // 1-based
}

/* ================= COMPUTED SHEETS ================= */
function refreshComputedSheets(){
  const ss = getSS();
  const payments = sheetToObjects(ss.getSheetByName(SHEET.PAYMENTS));
  const expenses = sheetToObjects(ss.getSheetByName(SHEET.EXPENSES));

  const totalCollection = payments.filter(p=>p['Payment Status']==='Paid')
    .reduce((s,p)=>s+Number(p['Amount']||0),0);
  const totalExpenses = expenses.reduce((s,e)=>s+Number(e['Amount']||0),0);
  const availableFund = totalCollection - totalExpenses;

  const fundSh = ss.getSheetByName(SHEET.FUND);
  fundSh.getRange(2,1,1,3).setValues([[totalCollection, totalExpenses, availableFund]]);

  const now = new Date();
  const curMonth = MONTH_NAMES[now.getMonth()], curYear = now.getFullYear();
  const last = new Date(now.getFullYear(), now.getMonth()-1, 1);
  const lastMonth = MONTH_NAMES[last.getMonth()], lastYear = last.getFullYear();
  const prev = new Date(now.getFullYear(), now.getMonth()-2, 1);
  const prevMonth = MONTH_NAMES[prev.getMonth()], prevYear = prev.getFullYear();

  const curCollection = payments.filter(p=>p['Payment Status']==='Paid' && p['Month']===curMonth && String(p['Year'])===String(curYear))
    .reduce((s,p)=>s+Number(p['Amount']||0),0);
  const pendingCount = payments.filter(p=>p['Month']===curMonth && String(p['Year'])===String(curYear) && p['Payment Status']!=='Paid').length;

  function expenseSum(m,y){
    return expenses.filter(e=>{
      const d = new Date(e['Expense Date']);
      return MONTH_NAMES[d.getMonth()]===m && d.getFullYear()===y;
    }).reduce((s,e)=>s+Number(e['Amount']||0),0);
  }
  const curExpense = expenseSum(curMonth, curYear);
  const lastExpense = expenseSum(lastMonth, lastYear);
  const prevExpense = expenseSum(prevMonth, prevYear);

  const dashSh = ss.getSheetByName(SHEET.DASHBOARD);
  dashSh.getRange(2,1,1,7).setValues([[totalCollection, availableFund, curCollection, curExpense, lastExpense, prevExpense, pendingCount]]);
}

/* ================= WEB APP ENTRY POINTS ================= */
function doGet(e){
  const action = (e.parameter && e.parameter.action) || 'getData';
  let result;
  if(action === 'getData') result = getAllData();
  else result = {error:'Unknown action'};
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  let payload = {};
  try{ payload = JSON.parse(e.postData.contents); }catch(err){ payload = {}; }
  const action = payload.action;
  let result = {ok:false, error:'Unknown action'};
  try{
    if(action === 'addPayment') result = addPayment(payload);
    else if(action === 'deletePayment') result = deletePayment(payload);
    else if(action === 'addFlat') result = addFlat(payload);
    else if(action === 'updateFlat') result = updateFlat(payload);
    else if(action === 'deleteFlat') result = deleteFlat(payload);
    else if(action === 'addExpense') result = addExpense(payload);
    else if(action === 'deleteExpense') result = deleteExpense(payload);
    else if(action === 'updateConfig') result = updateConfig(payload);
    refreshComputedSheets();
  }catch(err){
    result = {ok:false, error: String(err)};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= DATA READ ================= */
function getAllData(){
  const ss = getSS();
  const flats = sheetToObjects(ss.getSheetByName(SHEET.FLATS));
  const payments = sheetToObjects(ss.getSheetByName(SHEET.PAYMENTS));
  const expenses = sheetToObjects(ss.getSheetByName(SHEET.EXPENSES));
  const props = PropertiesService.getScriptProperties();
  const config = {
    buildingName: props.getProperty('buildingName') || 'Shree Ji Apartment',
    dueDay: Number(props.getProperty('dueDay') || 10),
    defaultAmount: Number(props.getProperty('defaultAmount') || 1000)
  };
  return {flats, payments, expenses, config};
}

/* ================= MUTATIONS ================= */
function addPayment(p){
  // p: {month, year, flatNo, ownerName, amount, status, date, mode, remarks}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.PAYMENTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  let rowIdx = -1;
  for(let i=1;i<data.length;i++){
    if(data[i][headers.indexOf('Month')]===p.month && String(data[i][headers.indexOf('Year')])===String(p.year) && data[i][headers.indexOf('Flat No')]===p.flatNo){
      rowIdx = i+1; break;
    }
  }
  const rowValues = [p.month, p.year, p.flatNo, p.ownerName, Number(p.amount), 'Paid', p.date, p.mode, p.remarks || ''];
  if(rowIdx > -1){
    sh.getRange(rowIdx,1,1,9).setValues([rowValues]);
  } else {
    sh.appendRow([...rowValues, Utilities.getUuid()]);
  }
  return {ok:true};
}
function deletePayment(p){
  // p: {month, year, flatNo}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.PAYMENTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  for(let i=1;i<data.length;i++){
    if(data[i][headers.indexOf('Month')]===p.month && String(data[i][headers.indexOf('Year')])===String(p.year) && data[i][headers.indexOf('Flat No')]===p.flatNo){
      sh.deleteRow(i+1);
      break;
    }
  }
  return {ok:true};
}
function addFlat(p){
  // p: {flatNo, ownerName, status, monthlyMaintenance, joiningDate, mobile, remarks}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.FLATS);
  sh.appendRow([p.flatNo, p.ownerName||'', p.status||'Active', Number(p.monthlyMaintenance)||1000, p.joiningDate||'', p.mobile||'', p.remarks||'']);
  return {ok:true};
}
function updateFlat(p){
  // p: {flatNo, ownerName, status, monthlyMaintenance, joiningDate, mobile, remarks}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.FLATS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  for(let i=1;i<data.length;i++){
    if(data[i][headers.indexOf('Flat No')] === p.flatNo){
      sh.getRange(i+1,2,1,6).setValues([[p.ownerName||'', p.status||'Active', Number(p.monthlyMaintenance)||1000, p.joiningDate||'', p.mobile||'', p.remarks||'']]);
      break;
    }
  }
  return {ok:true};
}
function deleteFlat(p){
  // p: {flatNo}
  const ss = getSS();
  const flatsSh = ss.getSheetByName(SHEET.FLATS);
  const data = flatsSh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0] === p.flatNo){ flatsSh.deleteRow(i+1); break; }
  }
  // cascade delete payment history for that flat
  const paySh = ss.getSheetByName(SHEET.PAYMENTS);
  const pdata = paySh.getDataRange().getValues();
  for(let i=pdata.length-1;i>=1;i--){
    if(pdata[i][2] === p.flatNo) paySh.deleteRow(i+1);
  }
  return {ok:true};
}
function addExpense(p){
  // p: {date, category, description, amount, addedBy}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.EXPENSES);
  sh.appendRow([p.date, p.category, p.description, Number(p.amount), p.addedBy || 'Admin', Utilities.getUuid()]);
  return {ok:true};
}
function deleteExpense(p){
  // p: {id}
  const ss = getSS();
  const sh = ss.getSheetByName(SHEET.EXPENSES);
  const rowIdx = findRowIndexById(sh, 'Record ID', p.id);
  if(rowIdx > -1) sh.deleteRow(rowIdx);
  return {ok:true};
}
const SETTINGS_PASSWORD = '2626';

function updateConfig(p){
  if(p.password !== SETTINGS_PASSWORD){
    return {ok:false, error:'Incorrect password'};
  }
  const props = PropertiesService.getScriptProperties();
  if(p.buildingName) props.setProperty('buildingName', p.buildingName);
  if(p.dueDay) props.setProperty('dueDay', String(p.dueDay));
  if(p.defaultAmount) props.setProperty('defaultAmount', String(p.defaultAmount));
  return {ok:true};
}
