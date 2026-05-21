// ====================================================================
// AM and Above Focused Partners — Code.gs v7
// FIXED: Column mapping for updated sheet (FTD added at col 24)
// FIXED: Hierarchy ZH→RH→SH→RM→AM using territory detection
// NEW:   FTD (For The Day) metric across all roles/zones
// ====================================================================

const SHEET_ID    = '1AgPaAik0vjh_9fcxX4NdV33-hWd4S5qNzwuFXb3xOis';
const USERS_SHEET = 'Users';
const LOG_SHEET   = 'LoginLog';
const DAILY_SHEET = 'DailyMTD';

// Zone sheet tab names to try in priority order
const ZONE_SHEETS = [
  {zone:'North', names:['Zone_North','North']},
  {zone:'South', names:['South','Zone_South']},
  {zone:'East',  names:['East & Central','East','Zone_East']},
  {zone:'West',  names:['West','Zone_West']},
  {zone:'RON',   names:['RON','Zone_RON']}
];
const MAIN_SHEETS = ['ALL PARTNERS','All_Partner List','Sheet1'];

// Column map — MAIN tab (ALL PARTNERS, 36 cols)
// KEY FIX: Zone added at col 5, pushed everything right by 1-2
const CM = {
  GID:1, NAME:2, CITY:3, STATE:4, ZONE:5,
  EMP_ID:6,
  OWNER_ROLE:7, OWNER_NAME:8, OWNER_EMP:9,
  MAX_POTENTIAL:10,
  MONTH_START:11, MONTH_END:23,
  OVERALL_POTENTIAL:24,
  TARGET:25,
  FTD_COL:null,
  MTD:26, LMTD:27,
  ACTIVE:28, GROWTH:29,
  CALLS:30, VISITS:31,
  REMARK:34
};

// Column map — ZONE sheets (Zone_North etc., 32 cols, FTD at 24)
const CZ = {
  GID:1, NAME:2, CITY:3, STATE:4, ZONE:5,
  OWNER_ROLE:6, OWNER_NAME:7,
  MAX_POTENTIAL:8,
  MONTH_START:9, MONTH_END:21,
  OVERALL_POTENTIAL:22,
  TARGET:23,
  FTD_COL:24,
  MTD:25, LMTD:26,
  ACTIVE:27, GROWTH:28,
  CALLS:29, VISITS:30,
  REMARK:31
};

const ROLE_LEVEL  = {ZH:5, RH:4, SH:3, RM:2, AM:1, MASTER:99};
const MASTER_GIDS = ['MASTER','IDK-MASTER','CENTRAL'];
const TELE_KW     = ['TELE-RM','TELE RM','TELERM'];

// ========================= ENTRY ====================================
function doGet(e) {
  var p = e.parameter || {};
  var result;
  try {
    switch (p.action) {
      case 'login':            result = handleLogin(p.gid, p.password); break;
      case 'checkPassword':    result = checkPasswordStatus(p.gid); break;
      case 'setPassword':      result = setPassword(p.gid, p.oldPassword, p.newPassword); break;
      case 'getDashboard':     result = getDashboard(p.gid); break;
      case 'getMaster':        result = getMasterDashboard(p.gid); break;
      case 'saveRemark':       result = saveRemark(p.gid, p.partnerGid, p.remark); break;
      case 'getLoginStats':    result = getLoginStats(p.gid); break;
      case 'getDailyTracking': result = getDailyTracking(p.gid); break;
      default: result = {success:false, message:'Unknown action.'};
    }
  } catch(err) {
    result = {success:false, message:'Server error: '+err.message+' | '+String(err.stack||'')};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ========================= HELPERS ==================================
function parseRole(raw) {
  return String(raw||'').replace(/[:.\s]/g,'').toUpperCase().slice(0,4);
}
function parseNumber(val) {
  if (val===null||val===undefined||val==='') return 0;
  if (typeof val==='number') return isFinite(val)?val:0;
  var n = parseFloat(String(val).replace(/[^0-9.\-]/g,''));
  return isNaN(n)?0:n;
}
function uniqueSorted(arr) {
  var seen={},out=[];
  arr.forEach(function(v){var s=String(v||'').trim(); if(s&&!seen[s]){seen[s]=true;out.push(s);}});
  return out.sort();
}
function isTeleRow(row, col) {
  var state=String(row[col.STATE]||'').toLowerCase();
  var role=String(row[col.OWNER_ROLE]||'').toUpperCase();
  if (state.indexOf('tele')!==-1) return true;
  for (var i=0;i<TELE_KW.length;i++){if(role.indexOf(TELE_KW[i])!==-1)return true;}
  return false;
}
function isMasterUser(user) {
  if (!user) return false;
  if (user.role==='MASTER') return true;
  return MASTER_GIDS.some(function(id){return user.gid.toUpperCase().indexOf(id)!==-1;});
}
function getSheet(ss, names) {
  for (var i=0;i<names.length;i++){var sh=ss.getSheetByName(names[i]);if(sh)return sh;}
  return null;
}

// ========================= PARTNER OBJECT ===========================
function buildPartner(row, col) {
  var gid = String(row[col.GID]||'').trim();
  if (!gid||gid===''||gid==='GID/GCD') return null;
  if (isTeleRow(row,col)) return null;

  var ownerRole = parseRole(String(row[col.OWNER_ROLE]||''));
  var ownerName = String(row[col.OWNER_NAME]||'').trim();
  if (!ownerRole) return null;

  var zone  = String(row[col.ZONE]||'').trim();
  var state = String(row[col.STATE]||'').trim();

  var monthlyData = [];
  for (var m=col.MONTH_START;m<=col.MONTH_END;m++) monthlyData.push(parseNumber(row[m]));

  var maxPot = parseNumber(row[col.MAX_POTENTIAL]);
  if (!maxPot && monthlyData.length) maxPot=Math.max.apply(null,monthlyData);

  var ftd  = col.FTD_COL!==null ? parseNumber(row[col.FTD_COL]) : 0;
  var mtd  = parseNumber(row[col.MTD]);
  var lmtd = parseNumber(row[col.LMTD]);

  var activeRaw = String(row[col.ACTIVE]||'').trim().toLowerCase();
  var isActive  = activeRaw==='1'||activeRaw==='active'||activeRaw==='yes'||mtd>0;

  var growthRaw = String(row[col.GROWTH]||'').trim();
  var growthNum = parseFloat(String(growthRaw).replace(/[^\-\d.]/g,''));
  var isGrowth;
  if (growthRaw.toLowerCase().indexOf('degrowth')!==-1) isGrowth=false;
  else if (growthRaw.toLowerCase().indexOf('growth')!==-1) isGrowth=true;
  else if (!isNaN(growthNum)) isGrowth=growthNum>=0;
  else isGrowth=mtd>=lmtd;

  var calls=parseNumber(row[col.CALLS]),visits=parseNumber(row[col.VISITS]);

  return {
    gid:gid, name:String(row[col.NAME]||'').trim(),
    city:String(row[col.CITY]||'').trim(), state:state, zone:zone,
    ownerRole:ownerRole, ownerName:ownerName,
    maxPotential:maxPot, overallPotential:parseNumber(row[col.OVERALL_POTENTIAL]),
    target:parseNumber(row[col.TARGET]),
    ftd:ftd, currentMonth:mtd, prevMonth:lmtd, monthlyData:monthlyData,
    isActive:isActive, isGrowth:isGrowth,
    calls:calls, visits:visits, connected:calls>0||visits>0,
    remark:String(row[col.REMARK]||'').trim()
  };
}

// ========================= LOAD ALL PARTNERS ========================
function loadAllPartners(ss) {
  var all={};

  // 1. Zone sheets (have FTD)
  ZONE_SHEETS.forEach(function(zs){
    var sh=getSheet(ss,zs.names);
    if (!sh) return;
    var rows=sh.getDataRange().getValues();
    for (var r=2;r<rows.length;r++){
      var p=buildPartner(rows[r],CZ);
      if (p&&!all[p.gid]) all[p.gid]=p;
    }
  });

  // 2. Main sheet fallback (no FTD)
  var mainSh=getSheet(ss,MAIN_SHEETS);
  if (mainSh){
    var rows=mainSh.getDataRange().getValues();
    for (var r=2;r<rows.length;r++){
      var p=buildPartner(rows[r],CM);
      if (p&&!all[p.gid]) all[p.gid]=p;
    }
  }

  return Object.values(all);
}

// ========================= TERRITORY & ACCESS =======================
function getUserTerritory(user, allPartners) {
  var myRole=user.role, myName=user.name.toLowerCase();
  var states={}, zones={};
  for (var i=0;i<allPartners.length;i++){
    var p=allPartners[i];
    if (p.ownerRole===myRole&&p.ownerName.toLowerCase()===myName){
      if(p.state)states[p.state]=true;
      if(p.zone) zones[p.zone]=true;
    }
  }
  return {states:Object.keys(states), zones:Object.keys(zones)};
}

function canSee(user, territory, partner) {
  if (isMasterUser(user)) return true;
  var pRole=partner.ownerRole, pName=partner.ownerName.toLowerCase(), myName=user.name.toLowerCase();

  if (user.role==='AM') return pRole==='AM'&&pName===myName;

  if (user.role==='RM'){
    if (pRole==='RM'&&pName===myName) return true;
    if (pRole==='AM'&&territory.states.indexOf(partner.state)!==-1) return true;
    return false;
  }

  if (user.role==='ZH'){
    if (!territory.zones.length) return false;
    return territory.zones.indexOf(partner.zone)!==-1||territory.states.indexOf(partner.state)!==-1;
  }

  // SH, RH: state-level territory
  if (!territory.states.length) return false;
  if (territory.states.indexOf(partner.state)===-1) return false;
  var myLevel=ROLE_LEVEL[user.role]||0, pLevel=ROLE_LEVEL[pRole]||0;
  if (pRole===user.role&&pName===myName) return true;
  if (pLevel>0&&pLevel<myLevel) return true;
  return false;
}

// ========================= AGGREGATIONS =============================
function buildSummary(partners) {
  var total=partners.length,curr=0,prev=0,ftd=0,maxPot=0,overPot=0,
      tgt=0,active=0,growth=0,connected=0,calls=0,visits=0;
  for (var i=0;i<total;i++){
    var p=partners[i];
    curr+=p.currentMonth; prev+=p.prevMonth; ftd+=p.ftd;
    maxPot+=p.maxPotential; overPot+=p.overallPotential; tgt+=p.target;
    if(p.isActive)active++; if(p.isGrowth)growth++; if(p.connected)connected++;
    calls+=p.calls; visits+=p.visits;
  }
  return {
    totalPartners:total, totalMaxPotential:maxPot, totalOverallPotential:overPot,
    totalTarget:tgt, totalFTD:ftd, currentMonthPremium:curr, prevMonthPremium:prev,
    activeCount:active, inactiveCount:total-active,
    growthCount:growth, degrowthCount:total-growth,
    connectedCount:connected, notConnectedCount:total-connected,
    totalCalls:calls, totalVisits:visits,
    achievementPct:tgt>0?Math.round(curr/tgt*100):0,
    momPct:prev>0?Math.round((curr-prev)/prev*100):0,
    maxPotAchPct:maxPot>0?Math.round(curr/maxPot*100):0,
    engagementPct:total>0?Math.round(connected/total*100):0,
    activePct:total>0?Math.round(active/total*100):0,
    growthPct:total>0?Math.round(growth/total*100):0
  };
}
function buildOverallProject(partners){
  var s=buildSummary(partners);
  return {totalPartners:s.totalPartners,activePartners:s.activeCount,inactivePartners:s.inactiveCount,
    connectedPartners:s.connectedCount,nonConnectedPartners:s.notConnectedCount,
    calls:s.totalCalls,visits:s.totalVisits,ftd:s.totalFTD,
    businessGenerated:s.currentMonthPremium,lmtd:s.prevMonthPremium,
    maxPotential:s.totalMaxPotential,overallPotential:s.totalOverallPotential,target:s.totalTarget,
    achievementPct:s.achievementPct,maxPotAchPct:s.maxPotAchPct,momPct:s.momPct,
    growthCount:s.growthCount,degrowthCount:s.degrowthCount,
    engagementPct:s.engagementPct,activePct:s.activePct,growthPct:s.growthPct};
}
function buildTeamBreakdown(partners,user){
  var myLevel=ROLE_LEVEL[user.role]||0,teamMap={};
  partners.forEach(function(p){
    var pLevel=ROLE_LEVEL[p.ownerRole]||0; if(pLevel>=myLevel)return;
    var key=p.ownerRole+'|'+p.ownerName;
    if(!teamMap[key])teamMap[key]={role:p.ownerRole,name:p.ownerName,zone:p.zone,partners:[]};
    teamMap[key].partners.push(p);
  });
  return Object.values(teamMap).map(function(m){
    return {role:m.role,name:m.name,zone:m.zone,
            summary:buildSummary(m.partners),overallProject:buildOverallProject(m.partners),partners:m.partners};
  }).sort(function(a,b){return b.summary.currentMonthPremium-a.summary.currentMonthPremium;});
}
function buildAmPerformance(partners){
  var amMap={};
  partners.forEach(function(p){
    if(p.ownerRole!=='AM')return;
    if(!amMap[p.ownerName])amMap[p.ownerName]={name:p.ownerName,zone:p.zone,states:{},partners:[]};
    amMap[p.ownerName].states[p.state]=true;
    amMap[p.ownerName].partners.push(p);
  });
  return Object.values(amMap).map(function(m){
    return {name:m.name,role:'AM',zone:m.zone,states:Object.keys(m.states),
            summary:buildSummary(m.partners),overallProject:buildOverallProject(m.partners),partners:m.partners};
  }).sort(function(a,b){return b.summary.currentMonthPremium-a.summary.currentMonthPremium;});
}
function buildRolePerformance(partners,role){
  var map={};
  partners.filter(function(p){return p.ownerRole===role;}).forEach(function(p){
    if(!map[p.ownerName])map[p.ownerName]={name:p.ownerName,zone:p.zone,partners:[]};
    map[p.ownerName].partners.push(p);
  });
  return Object.values(map).map(function(m){
    return {name:m.name,role:role,zone:m.zone,
            summary:buildSummary(m.partners),overallProject:buildOverallProject(m.partners),partnerCount:m.partners.length};
  }).sort(function(a,b){return b.summary.currentMonthPremium-a.summary.currentMonthPremium;});
}

// ========================= DAILY TRACKING ===========================
function ensureDailySheet(ss){
  var sh=ss.getSheetByName(DAILY_SHEET);
  if(!sh){
    sh=ss.insertSheet(DAILY_SHEET);
    sh.appendRow(['Date','TotalMTD','TotalFTD','NorthMTD','SouthMTD','EastMTD','WestMTD','RONMTD','Partners']);
    sh.getRange(1,1,1,9).setFontWeight('bold');
  }
  return sh;
}
function recordDailySnapshot(ss,allPartners){
  try{
    var today=Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd');
    var sh=ensureDailySheet(ss),data=sh.getDataRange().getValues();
    for(var i=1;i<data.length;i++){if(String(data[i][0]).slice(0,10)===today)return;}
    var zMap={},ftdTotal=0;
    allPartners.forEach(function(p){var z=p.zone||'Other';if(!zMap[z])zMap[z]=0;zMap[z]+=p.currentMonth;ftdTotal+=p.ftd;});
    var total=allPartners.reduce(function(s,p){return s+p.currentMonth;},0);
    sh.appendRow([today,total,ftdTotal,zMap['North']||0,zMap['South']||0,zMap['East']||0,zMap['West']||0,zMap['RON']||0,allPartners.length]);
  }catch(e){}
}
function getDailyTracking(gid){
  var user=getUser(gid);
  if(!user)return{success:false,message:'User not found.'};
  var ss=SpreadsheetApp.openById(SHEET_ID),sh=ss.getSheetByName(DAILY_SHEET);
  if(!sh)return{success:true,history:[],runRate:0,todayMTD:0,yesterdayMTD:0,todayFTD:0};
  var data=sh.getDataRange().getValues(),history=[];
  for(var i=1;i<data.length;i++){
    if(!data[i][0])continue;
    history.push({date:String(data[i][0]).slice(0,10),totalMTD:Number(data[i][1]||0),totalFTD:Number(data[i][2]||0),
      northMTD:Number(data[i][3]||0),southMTD:Number(data[i][4]||0),eastMTD:Number(data[i][5]||0),
      westMTD:Number(data[i][6]||0),ronMTD:Number(data[i][7]||0),partners:Number(data[i][8]||0)});
  }
  history.sort(function(a,b){return a.date.localeCompare(b.date);});
  var todayMTD=history.length?history[history.length-1].totalMTD:0;
  var yesterdayMTD=history.length>1?history[history.length-2].totalMTD:0;
  var todayFTD=history.length?history[history.length-1].totalFTD:0;
  return{success:true,history:history,runRate:todayMTD-yesterdayMTD,todayMTD:todayMTD,yesterdayMTD:yesterdayMTD,todayFTD:todayFTD};
}
function calcDailyRunRate(ss,currentMTD){
  try{
    var sh=ss.getSheetByName(DAILY_SHEET);
    if(!sh)return{runRate:0,yesterdayMTD:0};
    var data=sh.getDataRange().getValues();
    if(data.length<2)return{runRate:0,yesterdayMTD:0};
    var last=data[data.length-1],yesterdayMTD=Number(last[1]||0);
    return{runRate:currentMTD-yesterdayMTD,yesterdayMTD:yesterdayMTD};
  }catch(e){return{runRate:0,yesterdayMTD:0};}
}

// ========================= AUTH =====================================
function getUser(gid){
  if(!gid)return null;
  var ss=SpreadsheetApp.openById(SHEET_ID),sh=ss.getSheetByName(USERS_SHEET);
  if(!sh)return null;
  var data=sh.getDataRange().getValues(),target=String(gid).trim().toUpperCase();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0]).trim().toUpperCase()===target)
      return{rowIndex:i+1,gid:String(data[i][0]).trim().toUpperCase(),
             name:String(data[i][1]||'').trim(),role:String(data[i][2]||'').trim().toUpperCase(),
             zone:String(data[i][3]||'').trim(),password:String(data[i][4]||'').trim().toLowerCase()};
  }
  return null;
}
function checkPasswordStatus(gid){
  if(!gid)return{success:false,message:'Enter your User ID.'};
  var user=getUser(gid);if(!user)return{success:false,message:'User ID not found.'};
  return{success:true,hasPassword:!!(user.password&&user.password!==''&&user.password!=='null')};
}
function handleLogin(gid,hash){
  if(!gid)return{success:false,message:'Enter your User ID.'};
  var user=getUser(gid);if(!user)return{success:false,message:'User ID not found.'};
  if(!user.password||user.password===''||user.password==='null')return{success:false,needsPasswordSet:true,message:'No password set.'};
  if(user.password!==String(hash||'').toLowerCase())return{success:false,message:'Incorrect password.'};
  logLogin(user);return{success:true,user:{gid:user.gid,name:user.name,role:user.role,zone:user.zone}};
}
function setPassword(gid,oldHash,newHash){
  if(!gid||!newHash)return{success:false,message:'Missing parameters.'};
  var user=getUser(gid);if(!user)return{success:false,message:'User not found.'};
  var has=!!(user.password&&user.password!==''&&user.password!=='null');
  if(has&&user.password!==String(oldHash||'').toLowerCase())return{success:false,message:'Current password incorrect.'};
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET).getRange(user.rowIndex,5).setValue(String(newHash).toLowerCase());
  return{success:true};
}
function logLogin(user){
  try{var ss=SpreadsheetApp.openById(SHEET_ID),ls=ss.getSheetByName(LOG_SHEET);
    if(!ls){ls=ss.insertSheet(LOG_SHEET);ls.appendRow(['Timestamp','GID','Name','Role','Zone']);}
    ls.appendRow([new Date().toISOString(),user.gid,user.name,user.role,user.zone]);}catch(e){}
}
function getLoginStats(gid){
  var user=getUser(gid);if(!user)return{success:false,message:'Access denied.'};
  var ss=SpreadsheetApp.openById(SHEET_ID),ls=ss.getSheetByName(LOG_SHEET);
  if(!ls)return{success:true,stats:[],totalLogins:0};
  var data=ls.getDataRange().getValues(),stats={};
  for(var i=1;i<data.length;i++){var k=data[i][1];if(!stats[k])stats[k]={gid:data[i][1],name:data[i][2],role:data[i][3],zone:data[i][4],count:0,lastLogin:''};stats[k].count++;stats[k].lastLogin=data[i][0];}
  return{success:true,stats:Object.values(stats),totalLogins:data.length-1};
}

// ========================= REMARKS ==================================
function saveRemark(userGid,partnerGid,remark){
  if(!userGid||!partnerGid)return{success:false,message:'Missing params.'};
  var ss=SpreadsheetApp.openById(SHEET_ID),tgt=String(partnerGid).trim().toUpperCase();
  // Try zone sheets
  for(var z=0;z<ZONE_SHEETS.length;z++){
    var sh=getSheet(ss,ZONE_SHEETS[z].names);
    if(!sh)continue;
    var data=sh.getDataRange().getValues();
    for(var r=2;r<data.length;r++){
      if(String(data[r][CZ.GID]).trim().toUpperCase()===tgt){
        sh.getRange(r+1,CZ.REMARK+1).setValue(remark||'');return{success:true};
      }
    }
  }
  // Try main sheet
  var mainSh=getSheet(ss,MAIN_SHEETS);
  if(mainSh){
    var data=mainSh.getDataRange().getValues();
    for(var r=2;r<data.length;r++){
      if(String(data[r][CM.GID]).trim().toUpperCase()===tgt){
        mainSh.getRange(r+1,CM.REMARK+1).setValue(remark||'');return{success:true};
      }
    }
  }
  return{success:false,message:'Partner not found.'};
}

// ========================= DASHBOARD ================================
function getDashboard(gid){
  if(!gid)return{success:false,message:'GID required.'};
  var user=getUser(gid);if(!user)return{success:false,message:'User not found.'};
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var allPartners=loadAllPartners(ss);

  var partners,myTerritory,displayZones;
  if(isMasterUser(user)){
    partners=allPartners;
    myTerritory={states:[],zones:['North','South','East','West','RON']};
    displayZones=['North','South','East','West','RON'];
  } else {
    myTerritory=getUserTerritory(user,allPartners);
    displayZones=myTerritory.zones;
    partners=allPartners.filter(function(p){return canSee(user,myTerritory,p);});
  }

  var summary=buildSummary(partners),overallProject=buildOverallProject(partners);
  var rr=calcDailyRunRate(ss,summary.currentMonthPremium);
  if(isMasterUser(user))recordDailySnapshot(ss,allPartners);

  var myPartners=(user.role!=='AM'&&!isMasterUser(user))
    ?partners.filter(function(p){return p.ownerRole===user.role&&p.ownerName.toLowerCase()===user.name.toLowerCase();})
    :[];

  return{
    success:true,
    user:{gid:user.gid,name:user.name,role:user.role,zone:user.zone},
    summary:summary,overallProject:overallProject,partners:partners,
    teamBreakdown:(user.role!=='AM')?buildTeamBreakdown(partners,user):null,
    amPerformance:(user.role!=='AM')?buildAmPerformance(partners):null,
    myPartners:myPartners,
    filterOptions:{
      states:uniqueSorted(partners.map(function(p){return p.state;})),
      cities:uniqueSorted(partners.map(function(p){return p.city;})),
      owners:uniqueSorted(partners.map(function(p){return p.ownerName;}))
    },
    myZones:displayZones,dailyRunRate:rr.runRate,yesterdayMTD:rr.yesterdayMTD
  };
}

// ========================= MASTER DASHBOARD =========================
function getMasterDashboard(gid){
  var user=getUser(gid);if(!user)return{success:false,message:'User not found.'};
  var ss=SpreadsheetApp.openById(SHEET_ID);
  var allPartners=loadAllPartners(ss);

  var zoneMap={};
  allPartners.forEach(function(p){var z=p.zone||'Other';if(!zoneMap[z])zoneMap[z]=[];zoneMap[z].push(p);});

  var zoneSummaries=Object.keys(zoneMap)
    .filter(function(z){return z!=='Other'&&z!==''&&zoneMap[z].length>0;})
    .map(function(z){var zp=zoneMap[z],s=buildSummary(zp);
      return{zone:z,partnerCount:zp.length,summary:s,overallProject:buildOverallProject(zp)};})
    .sort(function(a,b){return b.summary.currentMonthPremium-a.summary.currentMonthPremium;});

  var stateMap={};
  allPartners.forEach(function(p){var st=p.state||'Unknown';if(!stateMap[st])stateMap[st]=[];stateMap[st].push(p);});
  var stateSummaries=Object.keys(stateMap)
    .map(function(st){return{state:st,zone:stateMap[st][0]?stateMap[st][0].zone:'',
      partnerCount:stateMap[st].length,summary:buildSummary(stateMap[st])};})
    .sort(function(a,b){return b.summary.currentMonthPremium-a.summary.currentMonthPremium;});

  var overallSummary=buildSummary(allPartners);
  var rr=calcDailyRunRate(ss,overallSummary.currentMonthPremium);
  recordDailySnapshot(ss,allPartners);

  return{
    success:true,totalPartners:allPartners.length,
    overallSummary:overallSummary,overallProject:buildOverallProject(allPartners),
    zoneSummaries:zoneSummaries,stateSummaries:stateSummaries,
    zhPerf:buildRolePerformance(allPartners,'ZH'),
    rhPerf:buildRolePerformance(allPartners,'RH'),
    shPerf:buildRolePerformance(allPartners,'SH'),
    rmPerf:buildRolePerformance(allPartners,'RM'),
    amPerf:buildRolePerformance(allPartners,'AM'),
    dailyRunRate:rr.runRate,yesterdayMTD:rr.yesterdayMTD
  };
}
