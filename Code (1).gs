// ====================================================================
// Partner Engage Portal — Code.gs v11
// Column maps VERIFIED from live sheet read 2026-05-21
// ====================================================================

const SHEET_ID    = '1AgPaAik0vjh_9fcxX4NdV33-hWd4S5qNzwuFXb3xOis';
const MAIN_SHEET  = 'All_Partner List';
const TELE_SHEET  = 'TELE_RM';
const USERS_SHEET = 'Users';
const LOG_SHEET   = 'LoginLog';
const DAILY_SHEET = 'DailyMTD';

const CM = {
  SN:           0,
  GID:          1,
  NAME:         2,
  CITY:         3,
  STATE:        4,
  ZONE:         5,
  OWNER_ROLE:   6,
  OWNER_NAME:   7,
  OWNER_EMP:    8,
  OWNER_EMAIL:  9,
  MAX_POT:      10,
  APR25:        11,
  MAY25:        12,
  JUN25:        13,
  JUL25:        14,
  AUG25:        15,
  SEP25:        16,
  OCT25:        17,
  NOV25:        18,
  DEC25:        19,
  JAN26:        20,
  FEB26:        21,
  MAR26:        22,
  APR26:        23,
  NET_COMBINED: 24,
  MONTHS_ACTIVE:25,
  AVG_MONTHLY:  26,
  OVERALL_POT:  27,
  TARGET:       28,
  FTD:          29,
  MTD:          30,
  LMTD:         31,
  ACTIVE:       32,
  GROWTH:       33,
  CALLS:        34,
  VISITS:       35,
  UNIQ_CALLS:   36,
  UNIQ_VISITS:  37,
  REMARK_SHEET: 38,
  REMARK_PART:  39,
};

const CT = {
  SN:           0,
  GID:          1,
  NAME:         2,
  CITY:         3,
  STATE:        4,
  OWNER_EMP:    5,
  OWNER_ROLE:   6,
  OWNER_NAME:   7,
  MAX_POT:      8,
  APR25:        9,
  MAY25:        10,
  JUN25:        11,
  JUL25:        12,
  AUG25:        13,
  SEP25:        14,
  OCT25:        15,
  NOV25:        16,
  DEC25:        17,
  JAN26:        18,
  FEB26:        19,
  MAR26:        20,
  APR26:        21,
  NET_COMBINED: 22,
  MONTHS_ACTIVE:23,
  AVG_MONTHLY:  24,
  OVERALL_POT:  25,
  TARGET:       26,
  FTD:          27,
  MTD:          28,
  GROWTH:       29,
  CALLS:        30,
};

const CU = {
  GID:  0,
  NAME: 1,
  ROLE: 2,
  ZONE: 3,
  PASS: 4,
};

var MONTH_LABELS = [
  "Apr'25","May'25","Jun'25","Jul'25","Aug'25","Sep'25",
  "Oct'25","Nov'25","Dec'25","Jan'26","Feb'26","Mar'26","Apr'26"
];

// ── Helpers ────────────────────────────────────────────────────────────
function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v).trim();
  if (s.match(/^#(N\/A|REF!|VALUE!|DIV\/0!|ERROR!|NAME\?)/)) return 0;
  s = s.replace(/[₹,\s]/g, '').replace(/^\\-/, '-');
  if (s.endsWith('%')) return parseFloat(s) || 0;
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function parseStr(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  return s.match(/^#(N\/A|REF!|VALUE!|DIV\/0!|ERROR!|NAME\?)/) ? '' : s;
}

function cleanRole(r) {
  return parseStr(r).replace(/:$/, '').trim().toUpperCase();
}

function hashPass(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    p, Utilities.Charset.UTF_8
  ).map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function getSheetByName(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase() === name.toLowerCase()) return all[i];
  }
  return null;
}

// ── Build partner from All_Partner List ────────────────────────────────
function buildMain(row) {
  var monthly = [];
  for (var c = CM.APR25; c <= CM.APR26; c++) monthly.push(parseNum(row[c]));
  var computed_active = monthly.filter(function(v){ return v > 0; }).length;
  var sheet_active    = parseNum(row[CM.MONTHS_ACTIVE]);
  var mtd             = parseNum(row[CM.MTD]);
  var lmtd            = parseNum(row[CM.LMTD]);
  var growthRaw       = parseStr(row[CM.GROWTH]);
  var growth          = 0;
  if (growthRaw !== '') {
    var gn = parseFloat(growthRaw.replace(/[%\\-]/g, function(m){ return m === '-' ? '-' : ''; }));
    if (isFinite(gn)) growth = growthRaw.includes('%') ? gn : gn * 100;
  } else {
    growth = lmtd > 0 ? ((mtd - lmtd) / lmtd * 100) : (mtd > 0 ? 100 : 0);
  }
  return {
    sn:           parseNum(row[CM.SN]),
    gid:          parseStr(row[CM.GID]),
    name:         parseStr(row[CM.NAME]),
    city:         parseStr(row[CM.CITY]),
    state:        parseStr(row[CM.STATE]),
    zone:         parseStr(row[CM.ZONE]),
    ownerRole:    cleanRole(row[CM.OWNER_ROLE]),
    ownerName:    parseStr(row[CM.OWNER_NAME]),
    ownerEmp:     parseStr(row[CM.OWNER_EMP]),
    ownerEmail:   parseStr(row[CM.OWNER_EMAIL]),
    maxPot:       parseNum(row[CM.MAX_POT]),
    monthly:      monthly,
    netCombined:  parseNum(row[CM.NET_COMBINED]),
    monthsActive: sheet_active > 0 ? sheet_active : computed_active,
    avgMonthly:   parseNum(row[CM.AVG_MONTHLY]),
    overallPot:   parseNum(row[CM.OVERALL_POT]),
    target:       parseNum(row[CM.TARGET]),
    ftd:          parseNum(row[CM.FTD]),
    mtd:          mtd,
    lmtd:         lmtd,
    isActive:     parseNum(row[CM.ACTIVE]) === 1,
    growth:       growth,
    calls:        parseNum(row[CM.CALLS]),
    visits:       parseNum(row[CM.VISITS]),
    uniqCalls:    parseNum(row[CM.UNIQ_CALLS]),
    uniqVisits:   parseNum(row[CM.UNIQ_VISITS]),
    remarkSheet:  parseStr(row[CM.REMARK_SHEET]),
    remarkPart:   parseStr(row[CM.REMARK_PART]),
    source:       'main',
  };
}

// ── Build partner from TELE_RM sheet ───────────────────────────────────
function buildTele(row) {
  var monthly = [];
  for (var c = CT.APR25; c <= CT.APR26; c++) monthly.push(parseNum(row[c]));
  var computed_active = monthly.filter(function(v){ return v > 0; }).length;
  var sheet_active    = parseNum(row[CT.MONTHS_ACTIVE]);
  var mtd             = parseNum(row[CT.MTD]);
  return {
    sn:           parseNum(row[CT.SN]),
    gid:          parseStr(row[CT.GID]),
    name:         parseStr(row[CT.NAME]),
    city:         parseStr(row[CT.CITY]),
    state:        'Tele-RM',
    zone:         'Tele-RM',
    ownerRole:    cleanRole(row[CT.OWNER_ROLE]),
    ownerName:    parseStr(row[CT.OWNER_NAME]),
    ownerEmp:     parseStr(row[CT.OWNER_EMP]),
    ownerEmail:   '',
    maxPot:       parseNum(row[CT.MAX_POT]),
    monthly:      monthly,
    netCombined:  parseNum(row[CT.NET_COMBINED]),
    monthsActive: sheet_active > 0 ? sheet_active : computed_active,
    avgMonthly:   parseNum(row[CT.AVG_MONTHLY]),
    overallPot:   parseNum(row[CT.OVERALL_POT]),
    target:       parseNum(row[CT.TARGET]),
    ftd:          parseNum(row[CT.FTD]),
    mtd:          mtd,
    lmtd:         0,
    isActive:     mtd > 0 || monthly[monthly.length - 1] > 0,
    growth:       0,
    calls:        parseNum(row[CT.CALLS]),
    visits:       0,
    uniqCalls:    0,
    uniqVisits:   0,
    remarkSheet:  '',
    remarkPart:   '',
    source:       'tele',
  };
}

// ── Load all partners ──────────────────────────────────────────────────
function loadAllPartners() {
  var partners = [];

  // 1. All_Partner List
  var sh = getSheetByName(MAIN_SHEET);
  if (sh) {
    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var gid = parseStr(data[r][CM.GID]);
      if (!gid.match(/^G[IC]D\d+/i)) continue;
      if (!parseStr(data[r][CM.NAME])) continue;
      try { partners.push(buildMain(data[r])); } catch(e) {}
    }
  }

  // 2. TELE_RM sheet
  var tsh = getSheetByName(TELE_SHEET);
  if (!tsh) {
    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var all2 = ss2.getSheets();
    for (var s = 0; s < all2.length; s++) {
      if (all2[s].getName().toLowerCase().indexOf('tele') !== -1) { tsh = all2[s]; break; }
    }
  }
  if (tsh) {
    var tdata = tsh.getDataRange().getValues();
    for (var tr = 1; tr < tdata.length; tr++) {
      var tgid = parseStr(tdata[tr][CT.GID]);
      if (!tgid.match(/^G[IC]D\d+/i)) continue;
      if (!parseStr(tdata[tr][CT.NAME])) continue;
      try { partners.push(buildTele(tdata[tr])); } catch(e) {}
    }
  }

  return partners;
}

// ── Load users ─────────────────────────────────────────────────────────
function loadUsers() {
  var sh = getSheetByName(USERS_SHEET);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var gid = parseStr(data[i][CU.GID]);
    if (!gid || !gid.startsWith('IDK-')) continue;
    var zonesRaw = parseStr(data[i][CU.ZONE]);
    users.push({
      gid:   gid.trim(),
      name:  parseStr(data[i][CU.NAME]),
      role:  parseStr(data[i][CU.ROLE]).toUpperCase(),
      zones: zonesRaw ? zonesRaw.split(',').map(function(z){ return z.trim().toLowerCase(); }) : [],
      pass:  parseStr(data[i][CU.PASS]),
    });
  }
  return users;
}

function findUser(gid) {
  var users = loadUsers();
  var g = gid.trim().toUpperCase();
  for (var i = 0; i < users.length; i++) {
    if (users[i].gid.toUpperCase() === g) return users[i];
  }
  return null;
}

// ── Zone matching ──────────────────────────────────────────────────────
function partnerZoneNorm(zone) {
  var z = zone.trim().toLowerCase();
  if (z === 'tele-rm' || z === 'tele_rm' || z === 'tele rm') return 'tele-rm';
  if (z.indexOf('east') !== -1) return 'east';
  return z;
}

function userCanAccessZone(userZones, partnerZone) {
  var pz = partnerZoneNorm(partnerZone);
  for (var i = 0; i < userZones.length; i++) {
    var uz = userZones[i].trim().toLowerCase();
    if (uz === 'all') return true;
    if (uz === pz) return true;
    if (uz.indexOf('tele') !== -1 && pz === 'tele-rm') return true;
    if (uz.indexOf('east') !== -1 && pz === 'east') return true;
    if (uz === 'north' && pz === 'north') return true;
    if (uz === 'south' && pz === 'south') return true;
    if (uz === 'west'  && pz === 'west')  return true;
    if (uz === 'ron'   && pz === 'ron')   return true;
  }
  return false;
}

function canUserSeePartner(user, partner) {
  // MASTER sees everything
  if (user.role === 'MASTER') return true;

  // Zone check
  if (!userCanAccessZone(user.zones, partner.zone)) return false;

  // ZH sees all in their zone
  if (user.role === 'ZH') return true;

  var uName    = user.name.trim().toLowerCase();
  var pOwner   = partner.ownerName.trim().toLowerCase();
  var pRole    = partner.ownerRole;

  if (user.role === 'RH') {
    if (pOwner === uName) return true;
    return pRole === 'SH' || pRole === 'RM' || pRole === 'AM';
  }
  if (user.role === 'SH') {
    if (pOwner === uName) return true;
    return pRole === 'RM' || pRole === 'AM';
  }
  if (user.role === 'RM') return pOwner === uName;
  if (user.role === 'AM') return pOwner === uName;

  return false;
}

// ── Summarise ──────────────────────────────────────────────────────────
function summarise(partners) {
  var s = {
    total: partners.length, connected: 0, notConnected: 0,
    ftd: 0, mtd: 0, lmtd: 0, growth: 0,
    overallPot: 0, target: 0, maxPot: 0, netCombined: 0,
    calls: 0, visits: 0,
    monthly: new Array(13).fill(0),
    achievement: 0,
    monthLabels: MONTH_LABELS,
  };
  partners.forEach(function(p) {
    if (p.isActive || p.mtd > 0 || p.calls > 0 || p.visits > 0) s.connected++;
    s.ftd         += p.ftd;
    s.mtd         += p.mtd;
    s.lmtd        += p.lmtd;
    s.overallPot  += p.overallPot;
    s.target      += p.target;
    s.maxPot      += p.maxPot;
    s.netCombined += p.netCombined;
    s.calls       += p.calls;
    s.visits      += p.visits;
    for (var m = 0; m < 13; m++) s.monthly[m] += (p.monthly[m] || 0);
  });
  s.notConnected = s.total - s.connected;
  s.growth       = s.lmtd > 0 ? ((s.mtd - s.lmtd) / s.lmtd * 100) : (s.mtd > 0 ? 100 : 0);
  s.achievement  = s.overallPot > 0 ? (s.mtd / s.overallPot * 100) : 0;
  return s;
}

// ── Group by owner ─────────────────────────────────────────────────────
function groupByOwner(partners) {
  var map = {};
  partners.forEach(function(p) {
    var key = p.ownerRole + '||' + p.ownerName.trim().toLowerCase();
    if (!map[key]) map[key] = { role: p.ownerRole, name: p.ownerName, emp: p.ownerEmp, list: [] };
    map[key].list.push(p);
  });
  return Object.values(map).map(function(o) {
    return Object.assign({ role: o.role, name: o.name, emp: o.emp }, summarise(o.list), { partners: o.list });
  }).sort(function(a, b){ return b.mtd - a.mtd; });
}

// ── LOGIN ──────────────────────────────────────────────────────────────
function doLogin(gid, pass) {
  var g = gid.trim().toUpperCase();

  // IDK-MASTER hardcoded (row in sheet has wrong column layout)
  if (g === 'IDK-MASTER') {
    var masterHash = '8d969eef6ecad3c29a3a629280e686cf0c3f5d5a86aff3ca12020c923adc6c92'; // = "123456"
    if (pass && pass !== '' && hashPass(pass) !== masterHash) {
      return { ok: false, error: 'Incorrect password' };
    }
    return { ok: true, user: { gid: 'IDK-MASTER', name: 'Master Admin', role: 'MASTER', zones: ['all'] } };
  }

  var user = findUser(gid);
  if (!user) return { ok: false, error: 'User not found' };

  // If password is set, validate it. If blank, allow any password.
  if (user.pass && user.pass.length > 10) {
    if (hashPass(pass) !== user.pass) return { ok: false, error: 'Incorrect password' };
  }

  try {
    var ls = getSheetByName(LOG_SHEET) || SpreadsheetApp.openById(SHEET_ID).insertSheet(LOG_SHEET);
    ls.appendRow([new Date().toISOString(), user.gid, user.name, user.role, user.zones.join(', ')]);
  } catch(e) {}

  return { ok: true, user: { gid: user.gid, name: user.name, role: user.role, zones: user.zones } };
}

// ── GET DATA (user dashboard) ──────────────────────────────────────────
function doGetData(userGid) {
  var user = findUser(userGid);
  if (!user) {
    if (userGid.toUpperCase() === 'IDK-MASTER') {
      user = { gid: 'IDK-MASTER', name: 'Master Admin', role: 'MASTER', zones: ['all'] };
    } else {
      return { ok: false, error: 'User not found: ' + userGid };
    }
  }

  var all  = loadAllPartners();
  var mine = all.filter(function(p){ return canUserSeePartner(user, p); });

  var mainMine = mine.filter(function(p){ return p.source === 'main'; });
  var teleMine = mine.filter(function(p){ return p.source === 'tele'; });

  return {
    ok:             true,
    user:           { gid: user.gid, name: user.name, role: user.role, zones: user.zones },
    summary:        summarise(mine),
    mainSummary:    summarise(mainMine),
    teleSummary:    summarise(teleMine),
    partners:       mine,
    ownerBreakdown: groupByOwner(mine),
    monthLabels:    MONTH_LABELS,
  };
}

// ── GET MASTER (master dashboard) ─────────────────────────────────────
function doGetMaster(userGid) {
  var user = findUser(userGid);
  if (!user) {
    if (userGid.toUpperCase() === 'IDK-MASTER') {
      user = { gid: 'IDK-MASTER', name: 'Master Admin', role: 'MASTER', zones: ['all'] };
    } else {
      return { ok: false, error: 'User not found' };
    }
  }

  var all = loadAllPartners();

  var zoneMap = {};
  all.forEach(function(p) {
    var z = p.zone || 'Unknown';
    if (!zoneMap[z]) zoneMap[z] = [];
    zoneMap[z].push(p);
  });
  var zoneSummaries = Object.keys(zoneMap).sort().map(function(z) {
    return Object.assign({ zone: z }, summarise(zoneMap[z]));
  });

  var cityMap = {}, stateMap = {};
  all.forEach(function(p) {
    var c = p.city  || 'Unknown';
    var s = p.state || 'Unknown';
    if (!cityMap[c])  cityMap[c]  = [];
    if (!stateMap[s]) stateMap[s] = [];
    cityMap[c].push(p);
    stateMap[s].push(p);
  });

  var cities = Object.keys(cityMap).map(function(c){ return Object.assign({ city: c },  summarise(cityMap[c]));  });
  var states = Object.keys(stateMap).map(function(s){ return Object.assign({ state: s }, summarise(stateMap[s])); });
  cities.sort(function(a, b){ return b.mtd - a.mtd; });
  states.sort(function(a, b){ return b.mtd - a.mtd; });

  return {
    ok:             true,
    user:           { gid: user.gid, name: user.name, role: user.role },
    overall:        summarise(all),
    zoneSummaries:  zoneSummaries,
    ownerBreakdown: groupByOwner(all),
    topCities:      cities.slice(0, 10),
    bottomCities:   cities.slice(-10).reverse(),
    topStates:      states.slice(0, 10),
    allStates:      states,
    totalPartners:  all.length,
    monthLabels:    MONTH_LABELS,
  };
}

// ── GET PARTNER DETAIL ─────────────────────────────────────────────────
function doGetPartner(gid) {
  var all = loadAllPartners();
  var g   = gid.trim().toUpperCase();
  for (var i = 0; i < all.length; i++) {
    if (all[i].gid.toUpperCase() === g) return { ok: true, partner: all[i] };
  }
  return { ok: false, error: 'Partner not found: ' + gid };
}

// ── DAILY MTD ──────────────────────────────────────────────────────────
function doSaveDailyMTD(ownerEmp, mtdVal) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = getSheetByName(DAILY_SHEET) || ss.insertSheet(DAILY_SHEET);
  sh.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'),
    ownerEmp, mtdVal, new Date().toISOString()
  ]);
  return { ok: true };
}

function doGetDailyStats(ownerEmp, days) {
  var sh = getSheetByName(DAILY_SHEET);
  if (!sh) return { ok: true, data: [] };
  var data   = sh.getDataRange().getValues();
  var result = [];
  for (var i = data.length - 1; i >= 1 && result.length < (days || 30); i--) {
    if (String(data[i][1]).trim() === String(ownerEmp).trim()) {
      result.unshift({ date: data[i][0], mtd: parseNum(data[i][2]) });
    }
  }
  return { ok: true, data: result };
}

// ── SHEET NAMES (debug) ────────────────────────────────────────────────
function doGetSheetNames() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return { ok: true, sheets: ss.getSheets().map(function(s){ return s.getName(); }) };
}

// ── ENTRY POINT ────────────────────────────────────────────────────────
function doGet(e) {
  var p      = (e && e.parameter) ? e.parameter : {};
  var cb     = p.callback || '';
  var result;

  try {
    switch (p.action) {
      case 'login':         result = doLogin(p.gid, p.pass);                         break;
      case 'getData':       result = doGetData(p.userGid);                           break;
      case 'getMaster':     result = doGetMaster(p.userGid);                         break;
      case 'getPartner':    result = doGetPartner(p.gid);                            break;
      case 'saveDailyMTD':  result = doSaveDailyMTD(p.ownerEmp, p.mtd);             break;
      case 'getDailyStats': result = doGetDailyStats(p.ownerEmp, parseInt(p.days)||30); break;
      case 'getSheetNames': result = doGetSheetNames();                              break;
      default:              result = { ok: false, error: 'Unknown action: ' + (p.action || 'none') };
    }
  } catch(err) {
    result = { ok: false, error: err.message, stack: err.stack };
  }

  var json = JSON.stringify(result);
  var out  = cb ? cb + '(' + json + ')' : json;
  return ContentService.createTextOutput(out)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
