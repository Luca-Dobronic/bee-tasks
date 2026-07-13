/**
 * CH CONSULT B-BBEE BOARD — Apps Script backend (Code.gs)
 * ------------------------------------------------------
 * SET-UP (once):
 * 1. Create a Google Sheet (any name). Copy its ID from the URL.
 * 2. Extensions → Apps Script → paste this file, set SHEET_ID below.
 * 3. Deploy → New deployment → Web app:
 *      Execute as: Me   |   Who has access: Anyone with the link
 * 4. Copy the /exec URL into the app's Connect box (you and Luca use the same URL).
 *
 * DATA MODEL (two tabs, auto-created):
 *   Elements: client | pillar | status | who | due
 *   Comments: id | client | pillar | parentId | who | when | text | flag
 *
 * SYNC MODEL: the app POSTs full state (last write wins) and GETs full state.
 * Simple and robust for a two-person team.
 */

const SHEET_ID = "1b9hZz63iwuuGmpC-mLvdNPgBgwvMMbUA1z-uopLp35Q";

const EL_HEAD = ["client","pillar","status","who","due"];
const CM_HEAD = ["id","client","pillar","parentId","who","when","text","flag","deleted"];
const LG_HEAD = ["when","who","action","detail"];
const CI_HEAD = ["client","period","fye","type","sector","drive"];
const AR_HEAD = ["client","period","closedOn","pillar","status","who","due","comments"];

function ss(){ return SpreadsheetApp.openById(SHEET_ID); }
function tab(name, head){
  let s = ss().getSheetByName(name);
  if(!s){ s = ss().insertSheet(name); s.appendRow(head); }
  return s;
}
function readTab(name, head){
  const s = tab(name, head);
  const vals = s.getDataRange().getValues();
  if(vals.length < 2) return [];
  return vals.slice(1).filter(r=>r[0]!=="").map(r=>{
    const o={}; head.forEach((h,i)=>o[h]= r[i]!==undefined? String(r[i]) : "");
    return o;
  });
}
function writeTab(name, head, rows){
  const s = tab(name, head);
  s.clearContents();
  const out=[head].concat(rows.map(o=>head.map(h=>o[h]!==undefined?o[h]:"")));
  s.getRange(1,1,out.length,head.length).setValues(out);
}

function doGet(){
  const data = { elements: readTab("Elements", EL_HEAD),
                 comments: readTab("Comments", CM_HEAD),
                 log:      readTab("Log", LG_HEAD),
                 clientinfo: readTab("ClientInfo", CI_HEAD),
                 archive:  readTab("Archive", AR_HEAD) };
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);                       // prevent Luca+Edrich writing at the same instant
  try{
    const data = JSON.parse(e.postData.contents);
    if(data.elements) writeTab("Elements", EL_HEAD, data.elements);
    if(data.comments) writeTab("Comments", CM_HEAD, data.comments);
    if(data.log)      writeTab("Log", LG_HEAD, data.log.slice(0,500));
    if(data.clientinfo) writeTab("ClientInfo", CI_HEAD, data.clientinfo);
    if(data.archive)  writeTab("Archive", AR_HEAD, data.archive);
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }finally{
    lock.releaseLock();
  }
}

/**
 * PHASE-2 (optional, off by default): daily due-date email.
 * In the Apps Script editor: Triggers → Add trigger → dailyDueReminder → time-driven → 7am.
 */
function dailyDueReminder(){
  const els = readTab("Elements", EL_HEAD);
  const today = new Date(new Date().toDateString());
  const soon = els.filter(r=>{
    if(!r.due || r.status==="Approved") return false;
    const d = Math.round((new Date(r.due)-today)/86400000);
    return d <= 2;                            // overdue or due within 2 days
  });
  if(!soon.length) return;
  const lines = soon.map(r=>{
    const d = Math.round((new Date(r.due)-today)/86400000);
    const state = d<0 ? Math.abs(d)+"d OVERDUE" : d===0 ? "due TODAY" : "due in "+d+"d";
    return "• "+r.client+" — "+r.pillar+" ("+r.status+", "+(r.who||"Both")+"): "+state;
  }).join("\n");
  MailApp.sendEmail({
    to: "edrich@chconsult.co.za, luca@chconsult.co.za",
    subject: "B-BBEE Board: "+soon.length+" element(s) due/overdue",
    body: "Good morning,\n\nThe following elements need attention:\n\n"+lines+
          "\n\nOpen the board: https://luca-dobronic.github.io/bee-tasks/\n"
  });
}
