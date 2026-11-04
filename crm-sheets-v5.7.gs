// ============================================================
//  CRM المعتمرين والرحلات - Google Apps Script (النسخة المحدّثة)
//  الأوراق (Sheets) المستخدمة:
//   Users, Clients, OmraPrograms, TurkeyPrograms, AdminNotes   (موجودة)
//   Programs, TripRequests, Settings                           (تُنشأ تلقائياً)
//  أعمدة Clients: A..W كما هي (23 عمود) + 5 أعمدة جديدة في آخرها:
//   X ProgramId | Y TripDetails | Z Feedback3 | AA FollowupTime | AB ReviewedBy
// ============================================================

var CLIENT_COLS = 30;

function doGet(e) {
  return handle_((e && e.parameter) ? e.parameter : {});
}

function doPost(e) {
  var p = {};
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    p = (e && e.parameter) ? e.parameter : {};
  }
  return handle_(p);
}

function handle_(p) {
  try {
    var action = p.action;
    var writeActions = {
      saveClient: 1, deleteClient: 1, toggleFollowup: 1, addProgram: 1, deleteProgram: 1,
      saveUser: 1, deleteUser: 1, sendAdminNote: 1, markNoteRead: 1,
      saveProgram: 1, deleteProgramById: 1, submitTripRequest: 1, decideTripRequest: 1, saveSetting: 1,
      transferClients: 1,
      saveTrip: 1, deleteTrip: 1, addTripMember: 1, removeTripMember: 1, saveRooming: 1, deleteRooming: 1,
      addWARecipient: 1, deleteWARecipient: 1, addWATemplate: 1, deleteWATemplate: 1, markWASent: 1
    };
    var lock = null;
    if (writeActions[action]) {
      lock = LockService.getScriptLock();
      lock.waitLock(25000);
    }
    try {
      return dispatch_(action, p);
    } finally {
      if (lock) lock.releaseLock();
    }
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function dispatch_(action, p) {
  if (action === "ping") return responseJSON({ status: "success", pong: true, version: 5.7 });
  if (action === "login") return loginUser(p.username, p.password);
  if (action === "getData") return getData(p);
  if (action === "saveClient") return saveClient(p);
  if (action === "deleteClient") return deleteClient(p.id);
  if (action === "toggleFollowup") return toggleFollowup(p.id, p.status);
  if (action === "addProgram") return addProgram(p);
  if (action === "deleteProgram") return deleteProgram(p);

  if (action === "getUsers") return getUsers(p.role);
  if (action === "saveUser") return saveUser(p);
  if (action === "deleteUser") return deleteUser(p);
  if (action === "sendAdminNote") return sendAdminNote(p);
  if (action === "markNoteRead") return markNoteRead(p.noteId);

  // الجديد
  if (action === "saveProgram") return saveProgram(p);
  if (action === "deleteProgramById") return deleteProgramById(p);
  if (action === "submitTripRequest") return submitTripRequest(p);
  if (action === "decideTripRequest") return decideTripRequest(p);
  if (action === "saveSetting") return saveSetting(p);
  if (action === "transferClients") return transferClients(p);
  if (action === "getArchive") return getArchive(p);
  if (action === "getFollowupLog") return getFollowupLog(p);

  // v5.5: قائمة الشركة والتسكين (للمديرين)
  if (action === "getTrips") return getTrips(p);
  if (action === "saveTrip") return saveTrip(p);
  if (action === "deleteTrip") return deleteTrip(p);
  if (action === "addTripMember") return addTripMember(p);
  if (action === "removeTripMember") return removeTripMember(p);
  if (action === "saveRooming") return saveRooming(p);
  if (action === "deleteRooming") return deleteRooming(p);

  // v5.6: إدارة الإرسال (واتساب)
  if (action === "getWA") return getWA(p);
  if (action === "addWARecipient") return addWARecipient(p);
  if (action === "deleteWARecipient") return deleteWARecipient(p);
  if (action === "addWATemplate") return addWATemplate(p);
  if (action === "deleteWATemplate") return deleteWATemplate(p);
  if (action === "markWASent") return markWASent(p);

  return responseJSON({ status: "error", message: "Action not found: " + action });
}

// ------------------------- أدوات مساعدة -------------------------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }

function isManager_(role) {
  role = String(role || "").trim().toLowerCase();
  return role === "admin" || role === "tech_manager" || role === "followup_manager";
}

function fmtDate_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, tz_(), "yyyy-MM-dd");
  return String(val);
}

function fmtTime_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, tz_(), "HH:mm");
  return String(val);
}

function fmtDateTime_(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, tz_(), "yyyy-MM-dd HH:mm");
  return String(val);
}

function nowStr_() {
  return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
}

function getOrCreateSheet_(name, headers) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function ensureClientColumns_(sheet) {
  var max = sheet.getMaxColumns();
  if (max < CLIENT_COLS) sheet.insertColumnsAfter(max, CLIENT_COLS - max);
  var hdr = ["ProgramId", "TripDetails", "Feedback3", "FollowupTime", "ReviewedBy", "SupervisorNote", "Feedback4"];
  var rng = sheet.getRange(1, 24, 1, 7);
  var cur = rng.getValues()[0];
  var changed = false;
  for (var i = 0; i < 7; i++) {
    if (!cur[i]) { cur[i] = hdr[i]; changed = true; }
  }
  if (changed) rng.setValues([cur]);
}

function safeParse_(s) {
  if (!s) return null;
  try { return JSON.parse(s); } catch (e) { return null; }
}

// ------------------------- الأرشيف: الحالات المفقودة (v5.4) -------------------------
var LOST_STATUSES_ = { "مش هيحجز": 1, "قفل السكه في وشي": 1, "مش بيرد": 1 };

function archiveSheet_() {
  var sh = ss_().getSheetByName("ClientsArchive");
  if (!sh) {
    var src = ss_().getSheetByName("Clients");
    sh = ss_().insertSheet("ClientsArchive");
    if (src && src.getLastRow() > 0) {
      var n = Math.max(src.getLastColumn(), CLIENT_COLS);
      sh.getRange(1, 1, 1, n).setValues(src.getRange(1, 1, 1, n).getValues());
    }
  }
  ensureClientColumns_(sh);
  return sh;
}

function findClientRow_(id) {
  var names = ["Clients", "ClientsArchive"];
  for (var s = 0; s < names.length; s++) {
    var sh = ss_().getSheetByName(names[s]);
    if (!sh) continue;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        return { sheet: sh, row: i + 1, data: data[i], isArchive: s === 1 };
      }
    }
  }
  return null;
}

function moveRowBetweenSheets_(fromSh, rowIndex, toSh) {
  var n = CLIENT_COLS;
  var row = fromSh.getRange(rowIndex, 1, 1, Math.max(fromSh.getLastColumn(), n)).getValues()[0];
  var padded = [];
  for (var c = 0; c < n; c++) padded.push(c < row.length ? row[c] : "");
  toSh.appendRow(padded);
  fromSh.deleteRow(rowIndex);
}

function mapClientRow_(row) {
  return {
    id: String(row[0]),
    agentCode: String(row[1] || ""),
    agentName: String(row[2] || ""),
    name: String(row[3] || ""),
    phone: String(row[4] || ""),
    address: String(row[5] || ""),
    members: row[6] || 1,
    programCategory: String(row[7] || ""),
    programName: String(row[8] || ""),
    turkeySubName: String(row[9] || ""),
    price: row[10] || 0,
    status: String(row[11] || ""),
    postponeDate: String(row[12] || ""),
    isFollowup: (String(row[13]).trim().toLowerCase() === "نعم" || row[13] === true),
    followupDate: fmtDate_(row[14]),
    followupDay: String(row[15] || ""),
    createdDate: fmtDate_(row[16]),
    fb1: String(row[17] || ""),
    fb2: String(row[18] || ""),
    paidAmount: row[19] || 0,
    paymentNotes: String(row[20] || ""),
    reviewedByFollowup: (String(row[21] || "").trim().toLowerCase() === "نعم" || row[21] === true),
    lastModifiedByRole: String(row[22] || ""),
    programId: String(row[23] || ""),
    tripDetails: safeParse_(row[24]),
    fb3: String(row[25] || ""),
    followupTime: fmtTime_(row[26]),
    reviewedBy: String(row[27] || ""),
    supervisorNote: String(row[28] || ""),
    fb4: String(row[29] || "")
  };
}

function appendNote_(from, to, clientId, clientName, text) {
  var sheet = getOrCreateSheet_("AdminNotes", ["ID", "FromUser", "ToUser", "ClientId", "ClientName", "NoteText", "Status", "CreatedAt"]);
  sheet.appendRow([
    new Date().getTime().toString() + Math.floor(Math.random() * 100),
    from || "النظام",
    to || "",
    clientId || "",
    clientName || "",
    text || "",
    "unread",
    nowStr_()
  ]);
}

// ------------------------- الدخول -------------------------
function loginUser(username, password) {
  var sheet = ss_().getSheetByName("Users");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Users غير موجودة!" });

  var data = sheet.getDataRange().getValues();
  var inputUser = String(username || "").trim().toLowerCase();
  var inputPass = String(password || "").trim();

  for (var i = 1; i < data.length; i++) {
    var dbUser = String(data[i][0] || "").trim().toLowerCase();
    var dbPass = String(data[i][1] || "").trim();
    if (dbUser === inputUser && dbPass === inputPass) {
      return responseJSON({
        status: "success",
        user: {
          username: data[i][0],
          code: String(data[i][2] || "").trim(),
          role: String(data[i][3] || "mowazaf").trim().toLowerCase(),
          name: data[i][4] || data[i][0]
        }
      });
    }
  }
  return responseJSON({ status: "error", message: "اسم المستخدم أو كلمة المرور غير صحيحة" });
}

// ------------------------- جلب البيانات -------------------------
function getData(p) {
  p = p || {};
  var role = String(p.role || "").trim().toLowerCase();
  var agentCode = String(p.code || "").trim().toLowerCase();
  var username = String(p.username || "").trim().toLowerCase();
  var dispName = String(p.name || "").trim().toLowerCase();
  var isAdminOrManager = isManager_(role);

  var ss = ss_();
  var clientSheet = ss.getSheetByName("Clients");
  if (!clientSheet) return responseJSON({ status: "error", message: "ورقة Clients غير موجودة" });
  var clientData = clientSheet.getDataRange().getValues();
  var clients = [];

  for (var i = 1; i < clientData.length; i++) {
    var row = clientData[i];
    if (!row[0] && row[0] !== 0) continue;

    var rowAgentCode = String(row[1] || "").trim().toLowerCase();
    var rowAgentName = String(row[2] || "").trim().toLowerCase();
    var isOwner = (agentCode !== "" && rowAgentCode === agentCode) ||
                  (rowAgentName !== "" && (rowAgentName === username || rowAgentName === dispName));
    if (!(isAdminOrManager || isOwner)) continue;

    clients.push({
      id: String(row[0]),
      agentCode: String(row[1] || ""),
      agentName: String(row[2] || ""),
      name: String(row[3] || ""),
      phone: String(row[4] || ""),
      address: String(row[5] || ""),
      members: row[6] || 1,
      programCategory: String(row[7] || ""),
      programName: String(row[8] || ""),
      turkeySubName: String(row[9] || ""),
      price: row[10] || 0,
      status: String(row[11] || ""),
      postponeDate: String(row[12] || ""),
      isFollowup: (String(row[13]).trim().toLowerCase() === "نعم" || row[13] === true),
      followupDate: fmtDate_(row[14]),
      followupDay: String(row[15] || ""),
      createdDate: fmtDate_(row[16]),
      fb1: String(row[17] || ""),
      fb2: String(row[18] || ""),
      paidAmount: row[19] || 0,
      paymentNotes: String(row[20] || ""),
      reviewedByFollowup: (String(row[21] || "").trim().toLowerCase() === "نعم" || row[21] === true),
      lastModifiedByRole: String(row[22] || ""),
      programId: String(row[23] || ""),
      tripDetails: safeParse_(row[24]),
      fb3: String(row[25] || ""),
      followupTime: fmtTime_(row[26]),
      reviewedBy: String(row[27] || ""),
      supervisorNote: String(row[28] || ""),
    fb4: String(row[29] || "")
    });
  }

  var omraPrograms = [];
  var omraSheet = ss.getSheetByName("OmraPrograms");
  if (omraSheet) {
    var oData = omraSheet.getDataRange().getValues();
    for (var j = 1; j < oData.length; j++) if (oData[j][0]) omraPrograms.push(String(oData[j][0]));
  }

  var turkeyPrograms = [];
  var turkeySheet = ss.getSheetByName("TurkeyPrograms");
  if (turkeySheet) {
    var tData = turkeySheet.getDataRange().getValues();
    for (var k = 1; k < tData.length; k++) if (tData[k][0]) turkeyPrograms.push(String(tData[k][0]));
  }

  // البرامج المفصلة
  var programs = [];
  var progSheet = ss.getSheetByName("Programs");
  if (progSheet) {
    var pData = progSheet.getDataRange().getValues();
    for (var q = 1; q < pData.length; q++) {
      if (!pData[q][0]) continue;
      if (String(pData[q][4] || "نعم") === "لا") continue;
      programs.push({
        id: String(pData[q][0]),
        category: String(pData[q][1] || ""),
        name: String(pData[q][2] || ""),
        details: safeParse_(pData[q][3]) || {},
        createdBy: String(pData[q][5] || ""),
        createdAt: fmtDateTime_(pData[q][6])
      });
    }
  }

  // التنبيهات الإدارية
  var adminNotes = [];
  var notesSheet = ss.getSheetByName("AdminNotes");
  if (notesSheet) {
    var nData = notesSheet.getDataRange().getValues();
    for (var m = 1; m < nData.length; m++) {
      var nRow = nData[m];
      if (!nRow[0]) continue;
      var toUser = String(nRow[2] || "").trim().toLowerCase();
      var forMe = (toUser === "all") || ((toUser !== "") && (toUser === username || toUser === agentCode || toUser === dispName));
      if (isAdminOrManager || forMe) {
        adminNotes.push({
          id: String(nRow[0]),
          fromUser: String(nRow[1] || ""),
          toUser: String(nRow[2] || ""),
          clientId: String(nRow[3] || ""),
          clientName: String(nRow[4] || ""),
          noteText: String(nRow[5] || ""),
          status: String(nRow[6] || "unread"),
          createdAt: fmtDateTime_(nRow[7])
        });
      }
    }
  }

  // طلبات تعديل الرحلات
  var tripRequests = [];
  var reqSheet = ss.getSheetByName("TripRequests");
  if (reqSheet) {
    var rData = reqSheet.getDataRange().getValues();
    for (var r = 1; r < rData.length; r++) {
      var rr = rData[r];
      if (!rr[0]) continue;
      var by = String(rr[3] || "").trim().toLowerCase();
      var byCode = String(rr[4] || "").trim().toLowerCase();
      var mine = (by !== "" && (by === username || by === dispName)) || (byCode !== "" && byCode === agentCode);
      if (isAdminOrManager || mine) {
        tripRequests.push({
          id: String(rr[0]),
          clientId: String(rr[1] || ""),
          clientName: String(rr[2] || ""),
          requestedBy: String(rr[3] || ""),
          requestedByCode: String(rr[4] || ""),
          proposed: safeParse_(rr[5]) || {},
          reason: String(rr[6] || ""),
          status: String(rr[7] || "pending"),
          decidedBy: String(rr[8] || ""),
          decidedAt: fmtDateTime_(rr[9]),
          createdAt: fmtDateTime_(rr[10])
        });
      }
    }
  }

  // الإعدادات (الأهداف)
  var settings = {};
  var setSheet = ss.getSheetByName("Settings");
  if (setSheet) {
    var sData = setSheet.getDataRange().getValues();
    for (var s = 1; s < sData.length; s++) if (sData[s][0]) settings[String(sData[s][0])] = String(sData[s][1] || "");
  }

  // v5.6: مزامنة تلقائية — نقل أي صف حالته من الحالات المفقودة للأرشيف فوراً
  try {
    var lostRows = [];
    for (var li = 1; li < clientData.length; li++) {
      if (clientData[li][0] && LOST_STATUSES_[String(clientData[li][11])]) lostRows.push(li + 1);
    }
    if (lostRows.length) {
      var archSh = archiveSheet_();
      for (var lr = lostRows.length - 1; lr >= 0; lr--) {
        moveRowBetweenSheets_(clientSheet, lostRows[lr], archSh);
      }
    }
  } catch (syncErr) {}

  return responseJSON({
    status: "success",
    clients: clients,
    omraPrograms: omraPrograms,
    turkeyPrograms: turkeyPrograms,
    programs: programs,
    adminNotes: adminNotes,
    tripRequests: tripRequests,
    settings: settings
  });
}

// ------------------------- العملاء -------------------------
function saveClient(p) {
  var sheet = ss_().getSheetByName("Clients");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Clients غير موجودة" });
  ensureClientColumns_(sheet);

  var id = p.id;
  var isEdit = !!(id && String(id) !== "0");
  var userRole = String(p.editorRole || p.userRole || "").trim().toLowerCase();
  var isFollowupMgr = (userRole === "followup_manager");

  var targetSheet = sheet, rowIndex = -1, old = null, inArchive = false;
  if (isEdit) {
    var found = findClientRow_(id);
    if (!found) return responseJSON({ status: "error", message: "العميل غير موجود" });
    targetSheet = found.sheet;
    rowIndex = found.row;
    old = found.data;
    inArchive = found.isArchive;
  }

  var newId = isEdit ? String(id) : new Date().getTime().toString();

  // الموظف الكاتب لا يتغير عند التعديل (حتى لو عدّل المدير)
  var agentCode = p.agentCode || "";
  var agentName = p.agentName || "";
  if (old && (old[1] || old[2])) {
    agentCode = String(old[1] || "");
    agentName = String(old[2] || "");
  }

  var reviewed = isFollowupMgr ? "نعم" : ((old && old[21]) ? old[21] : (p.reviewedByFollowup || "لا"));
  var reviewedBy = isFollowupMgr ? (p.editorName || "مسؤول المتابعة") : ((old && old[27]) ? old[27] : "");
  var createdDate = p.createdDate || ((old && old[16]) ? fmtDate_(old[16]) : Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd"));

  // تفاصيل الرحلة: الموظف لا يعدلها إلا عبر طلب موافقة
  var tripDetails = (p.tripDetails == null) ? "" : (typeof p.tripDetails === "string" ? p.tripDetails : JSON.stringify(p.tripDetails));
  var programId = String(p.programId || "");
  if (old) {
    var oldTrip = String(old[24] || "");
    var oldPid = String(old[23] || "");
    if (!isManager_(userRole) && programId === oldPid && oldTrip) tripDetails = oldTrip;
    else if (!tripDetails) tripDetails = oldTrip;
  }

  var supNote = isManager_(userRole) ? String(p.supervisorNote || "") : ((old && old[28] !== undefined) ? String(old[28]) : "");
  if (isFollowupMgr && p.supervisorNote) supNote = String(p.supervisorNote);

  var rowValues = [
    newId, agentCode, agentName, p.name || "", p.phone || "", p.address || "", p.members || 1,
    p.programCategory || "", p.programName || "", p.turkeySubName || "", p.price || 0, p.status || "",
    p.postponeDate || "", p.isFollowup || "لا", p.followupDate || "", p.followupDay || "", createdDate,
    p.fb1 || "", p.fb2 || "", p.paidAmount || 0, p.paymentNotes || "", reviewed, userRole, programId,
    tripDetails, p.fb3 || "", p.followupTime || "", reviewedBy, supNote, p.fb4 || ""
  ];

  var targetRow = isEdit ? rowIndex : sheet.getLastRow() + 1;
  targetSheet.getRange(targetRow, 5).setNumberFormat("@");
  targetSheet.getRange(targetRow, 15).setNumberFormat("@");
  targetSheet.getRange(targetRow, 17).setNumberFormat("@");
  targetSheet.getRange(targetRow, 27).setNumberFormat("@");
  targetSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);

  // نقل تلقائي بين الشيتين حسب حالة الحجز (أرشيف للحالات المفقودة)
  var finalLost = !!LOST_STATUSES_[String(p.status)];
  if (finalLost && !inArchive) {
    moveRowBetweenSheets_(targetSheet, targetRow, archiveSheet_());
  } else if (!finalLost && inArchive) {
    moveRowBetweenSheets_(targetSheet, targetRow, sheet);
  }

  // إشعارات تلقائية
  var prevStatus = old ? String(old[11] || "") : "";
  if ((p.status === "هيحجز" || p.status === "Near Closed") && prevStatus !== p.status) {
    appendNote_(agentName || "موظف", "المشرف", newId, p.name, "العميل أصبح (" + p.status + ") - قيد الحجز، يحتاج استكمال من المشرف.");
  }
  if (old && isFollowupMgr && agentName) {
    appendNote_(p.editorName || "مسؤول المتابعة", agentName, newId, p.name, "تم تعديل بيانات هذا العميل بواسطة مسؤول المتابعة.");
  }

  // v5.7: تسجيل "تمت المتابعة" في سجل المتابعات اليومي (للتقرير المطبوع)
  if (String(p.followupDone) === "نعم") {
    try { logFollowup_(newId, p); } catch (logErr) {}
  }

  // v5.5: أي عميل "تم الحجز والتعاقد" يُضاف تلقائياً لقائمة الشركة (رحلة باسم البرنامج أو "عام")
  if (String(p.status) === "تم الحجز والتعاقد") {
    var tripName = String(p.programName || "").trim() || "عام";
    addToCompanyTrip_(tripName, {
      clientId: newId,
      name: p.name || "",
      phone: p.phone || "",
      members: Number(p.members || 1),
      programName: p.programName || "",
      price: Number(p.price || 0),
      agentName: agentName || ""
    });
  }

  return responseJSON({ status: "success", id: newId });
}

function toggleFollowup(id, status) {
  var found = findClientRow_(id);
  if (!found) return responseJSON({ status: "error", message: "العميل غير موجود" });
  found.sheet.getRange(found.row, 14).setValue(status);
  return responseJSON({ status: "success" });
}

function deleteClient(id) {
  var found = findClientRow_(id);
  if (!found) return responseJSON({ status: "error", message: "العميل غير موجود" });
  found.sheet.deleteRow(found.row);
  return responseJSON({ status: "success" });
}

// ------------------------- البرامج (الأسماء السريعة القديمة) -------------------------
function addProgram(p) {
  if (p.role && !isManager_(p.role)) return responseJSON({ status: "error", message: "غير مصرح لك بإضافة البرامج" });
  var targetSheet = (p.type === "تركيا") ? "TurkeyPrograms" : "OmraPrograms";
  var sheet = ss_().getSheetByName(targetSheet);
  if (sheet) {
    sheet.appendRow([p.programName]);
    return responseJSON({ status: "success" });
  }
  return responseJSON({ status: "error", message: "ورقة البرامج غير موجودة" });
}

function deleteProgram(p) {
  if (p.role && !isManager_(p.role)) return responseJSON({ status: "error", message: "غير مصرح لك بحذف البرامج" });
  var targetSheet = (p.type === "تركيا") ? "TurkeyPrograms" : "OmraPrograms";
  var sheet = ss_().getSheetByName(targetSheet);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(p.programName).trim()) {
        sheet.deleteRow(i + 1);
        return responseJSON({ status: "success" });
      }
    }
  }
  return responseJSON({ status: "error", message: "البرنامج غير موجود" });
}

// ------------------------- البرامج المفصلة (جديد) -------------------------
function programsSheet_() {
  return getOrCreateSheet_("Programs", ["ID", "Category", "Name", "DetailsJSON", "Active", "CreatedBy", "CreatedAt"]);
}

function saveProgram(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "إضافة البرامج للمديرين فقط" });
  var name = String(p.name || "").trim();
  var category = String(p.category || "").trim();
  if (!name || !category) return responseJSON({ status: "error", message: "اسم البرنامج ونوع الرحلة مطلوبان" });
  var details = (typeof p.details === "string") ? p.details : JSON.stringify(p.details || {});

  var sheet = programsSheet_();
  var data = sheet.getDataRange().getValues();

  if (p.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[category, name, details]]);
        return responseJSON({ status: "success", message: "تم تحديث البرنامج", id: String(p.id) });
      }
    }
  }
  var newId = "P" + new Date().getTime();
  sheet.appendRow([newId, category, name, details, "نعم", p.createdBy || "", nowStr_()]);
  return responseJSON({ status: "success", message: "تم حفظ البرنامج", id: newId });
}

function deleteProgramById(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "حذف البرامج للمديرين فقط" });
  var sheet = programsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) {
      sheet.deleteRow(i + 1);
      return responseJSON({ status: "success", message: "تم حذف البرنامج" });
    }
  }
  return responseJSON({ status: "error", message: "البرنامج غير موجود" });
}

// ------------------------- طلبات تعديل الرحلة (جديد) -------------------------
function requestsSheet_() {
  return getOrCreateSheet_("TripRequests", ["ID", "ClientId", "ClientName", "RequestedBy", "RequestedByCode", "ProposedJSON", "Reason", "Status", "DecidedBy", "DecidedAt", "CreatedAt"]);
}

function submitTripRequest(p) {
  if (!findClientRow_(p.clientId)) return responseJSON({ status: "error", message: "العميل غير موجود" });

  var sheet = requestsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][1]) === String(p.clientId) && String(data[j][7]) === "pending") {
      return responseJSON({ status: "error", message: "يوجد طلب تعديل معلّق لهذا العميل بالفعل، انتظر رد المدير" });
    }
  }
  var proposed = (typeof p.proposed === "string") ? p.proposed : JSON.stringify(p.proposed || {});
  sheet.appendRow([
    "R" + new Date().getTime(),
    String(p.clientId),
    p.clientName || "",
    p.requestedBy || "",
    p.requestedByCode || "",
    proposed,
    p.reason || "",
    "pending",
    "",
    "",
    nowStr_()
  ]);
  return responseJSON({ status: "success", message: "تم إرسال الطلب للمدير" });
}

function decideTripRequest(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "الموافقة على الطلبات للمديرين فقط" });
  var decision = String(p.decision || "");
  if (decision !== "approved" && decision !== "rejected") return responseJSON({ status: "error", message: "قرار غير صالح" });

  var sheet = requestsSheet_();
  var data = sheet.getDataRange().getValues();
  var decider = p.deciderName || "المدير";

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(p.requestId)) continue;
    if (String(data[i][7]) !== "pending") return responseJSON({ status: "error", message: "تم البت في هذا الطلب مسبقاً" });

    var clientId = String(data[i][1]);
    var clientName = String(data[i][2]);
    var requester = String(data[i][3]);

    if (decision === "approved") {
      var cFound = findClientRow_(clientId);
      if (!cFound) return responseJSON({ status: "error", message: "العميل لم يعد موجوداً" });
      var proposed = safeParse_(data[i][5]) || {};
      proposed._approvedBy = decider;
      proposed._approvedAt = nowStr_();
      cFound.sheet.getRange(cFound.row, 25).setValue(JSON.stringify(proposed));
    }

    sheet.getRange(i + 1, 8, 1, 3).setValues([[decision, decider, nowStr_()]]);

    var msg = (decision === "approved")
      ? "تمت الموافقة على طلب تعديل بيانات الرحلة وتم حفظ البيانات. المدير: " + decider
      : "لم تتم الموافقة على طلب تعديل بيانات الرحلة، وتم الإبقاء على البيانات القديمة. المدير: " + decider;
    appendNote_(decider, requester, clientId, clientName, msg);

    return responseJSON({ status: "success", message: decision === "approved" ? "تمت الموافقة وحفظ البيانات" : "تم رفض الطلب" });
  }
  return responseJSON({ status: "error", message: "الطلب غير موجود" });
}

// ------------------------- نقل جماعي للعملاء (v5.4 - المدير فقط) -------------------------
function transferClients(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "نقل العملاء للمديرين فقط" });
  var ids = String(p.ids || "").split(",").map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
  if (!ids.length) return responseJSON({ status: "error", message: "لم يتم تحديد أي عملاء" });
  var sheet = ss_().getSheetByName("Clients");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Clients غير موجودة" });
  var data = sheet.getDataRange().getValues();
  var moved = 0;
  for (var i = 1; i < data.length; i++) {
    if (ids.indexOf(String(data[i][0])) !== -1) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[String(p.agentCode || ""), String(p.agentName || "")]]);
      moved++;
    }
  }
  if (p.agentName) {
    appendNote_(p.deciderName || "المدير", p.agentName, "", "نقل جماعي", "تم نقل " + moved + " عميل إليك من قبل المدير لمتابعتهم واستكمال الحجز معهم.");
  }
  return responseJSON({ status: "success", message: "تم نقل " + moved + " عميل إلى " + (p.agentName || "الموظف") });
}

// ------------------------- الإعدادات / الأهداف (جديد) -------------------------
function saveSetting(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "الإعدادات للمديرين فقط" });
  var sheet = getOrCreateSheet_("Settings", ["Key", "Value"]);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.key)) {
      sheet.getRange(i + 1, 2).setValue(String(p.value));
      return responseJSON({ status: "success" });
    }
  }
  sheet.appendRow([String(p.key), String(p.value)]);
  return responseJSON({ status: "success" });
}

// ------------------------- الأرشيف: جلب صفحات 50 سجل (v5.4) -------------------------
function getArchive(p) {
  var role = String(p.role || "").trim().toLowerCase();
  var agentCode = String(p.code || "").trim().toLowerCase();
  var username = String(p.username || "").trim().toLowerCase();
  var dispName = String(p.name || "").trim().toLowerCase();
  var mgr = isManager_(role);
  var q = String(p.q || "").trim().toLowerCase();
  var page = Math.max(1, parseInt(p.page || "1", 10) || 1);
  var perPage = 50;

  var clients = [], total = 0;
  var sh = ss_().getSheetByName("ClientsArchive");
  if (sh) {
    var data = sh.getDataRange().getValues();
    var matched = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && row[0] !== 0) continue;
      var rowAgentCode = String(row[1] || "").trim().toLowerCase();
      var rowAgentName = String(row[2] || "").trim().toLowerCase();
      var isOwner = (agentCode !== "" && rowAgentCode === agentCode) ||
                    (rowAgentName !== "" && (rowAgentName === username || rowAgentName === dispName));
      if (!(mgr || isOwner)) continue;
      if (q) {
        var full = (row[3] + " " + row[4] + " " + row[8] + " " + row[11] + " " + row[2]).toLowerCase();
        if (full.indexOf(q) === -1) continue;
      }
      matched.push(mapClientRow_(row));
    }
    total = matched.length;
    clients = matched.slice((page - 1) * perPage, page * perPage);
  }
  return responseJSON({ status: "success", clients: clients, total: total, page: page, perPage: perPage });
}

// ------------------------- v5.5: قائمة الشركة والتسكين (للمديرين) -------------------------
function tripsSheet_() {
  return getOrCreateSheet_("Trips", ["TripId", "TripName", "MembersJSON", "CreatedAt"]);
}

function roomingsSheet_() {
  return getOrCreateSheet_("Roomings", ["RoomingId", "TripId", "Name", "DataJSON", "CreatedAt"]);
}

function findTripRow_(tripId) {
  var sh = tripsSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(tripId)) return { sheet: sh, row: i + 1, data: data[i] };
  }
  return null;
}

function getTripMembers_(row) { return safeParse_(row[2]) || []; }

function saveTripMembers_(tripId, members) {
  var f = findTripRow_(tripId);
  if (f) f.sheet.getRange(f.row, 3).setValue(JSON.stringify(members || []));
}

function ensureTrip_(name) {
  var sh = tripsSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === name) return String(data[i][0]);
  }
  var id = "T" + new Date().getTime();
  sh.appendRow([id, name, "[]", nowStr_()]);
  return id;
}

function addToCompanyTrip_(tripName, member) {
  var tripId = ensureTrip_(tripName);
  var f = findTripRow_(tripId);
  if (!f) return;
  var members = getTripMembers_(f.data);
  if (member.clientId && members.some(function (m) { return String(m.clientId) === String(member.clientId); })) return;
  member.addedAt = nowStr_();
  members.push(member);
  saveTripMembers_(tripId, members);
}

function getTrips(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "قائمة الشركة والتسكين للمديرين فقط" });
  var trips = [];
  var sh = tripsSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    trips.push({
      id: String(data[i][0]),
      name: String(data[i][1] || ""),
      members: safeParse_(data[i][2]) || [],
      createdAt: fmtDateTime_(data[i][3])
    });
  }
  var roomings = [];
  var rsh = roomingsSheet_();
  var rdata = rsh.getDataRange().getValues();
  for (var j = 1; j < rdata.length; j++) {
    if (!rdata[j][0]) continue;
    roomings.push({
      id: String(rdata[j][0]),
      tripId: String(rdata[j][1] || ""),
      name: String(rdata[j][2] || ""),
      data: safeParse_(rdata[j][3]) || {},
      createdAt: fmtDateTime_(rdata[j][4])
    });
  }
  return responseJSON({ status: "success", trips: trips, roomings: roomings });
}

function saveTrip(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "إدارة الرحلات للمديرين فقط" });
  var name = String(p.name || "").trim();
  if (!name) return responseJSON({ status: "error", message: "اسم الرحلة مطلوب" });
  if (p.id) {
    var f = findTripRow_(p.id);
    if (!f) return responseJSON({ status: "error", message: "الرحلة غير موجودة" });
    f.sheet.getRange(f.row, 2).setValue(name);
    return responseJSON({ status: "success", message: "تم تحديث الرحلة", id: String(p.id) });
  }
  var id = ensureTrip_(name);
  return responseJSON({ status: "success", message: "تم إنشاء الرحلة", id: id });
}

function deleteTrip(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "حذف الرحلات للمديرين فقط" });
  var f = findTripRow_(p.tripId);
  if (!f) return responseJSON({ status: "error", message: "الرحلة غير موجودة" });
  f.sheet.deleteRow(f.row);
  // حذف كل تسكينات الرحلة
  var rsh = roomingsSheet_();
  var rdata = rsh.getDataRange().getValues();
  for (var i = rdata.length - 1; i >= 1; i--) {
    if (String(rdata[i][1]) === String(p.tripId)) rsh.deleteRow(i + 1);
  }
  return responseJSON({ status: "success", message: "تم حذف الرحلة وكل تسكيناتها" });
}

function addTripMember(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "إضافة المعتمرين للمديرين فقط" });
  var f = findTripRow_(p.tripId);
  if (!f) return responseJSON({ status: "error", message: "الرحلة غير موجودة" });
  var members = getTripMembers_(f.data);
  var m = {
    name: String(p.name || "").trim(),
    phone: String(p.phone || ""),
    clientId: String(p.clientId || ""),
    members: Number(p.members || 1),
    programName: String(p.programName || ""),
    price: Number(p.price || 0),
    agentName: String(p.agentName || ""),
    addedAt: nowStr_()
  };
  if (!m.name) return responseJSON({ status: "error", message: "اسم العميل مطلوب" });
  if (m.clientId && members.some(function (x) { return String(x.clientId) === m.clientId; })) {
    return responseJSON({ status: "error", message: "العميل موجود في الرحلة بالفعل" });
  }
  members.push(m);
  saveTripMembers_(p.tripId, members);
  return responseJSON({ status: "success", message: "تمت إضافة " + m.name });
}

function removeTripMember(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "للمديرين فقط" });
  var f = findTripRow_(p.tripId);
  if (!f) return responseJSON({ status: "error", message: "الرحلة غير موجودة" });
  var members = getTripMembers_(f.data);
  members = members.filter(function (m, idx) { return idx !== Number(p.index); });
  saveTripMembers_(p.tripId, members);
  return responseJSON({ status: "success", message: "تم الحذف من الرحلة" });
}

function saveRooming(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "التسكين للمديرين فقط" });
  var dataStr = (typeof p.data === "string") ? p.data : JSON.stringify(p.data || {});
  var rsh = roomingsSheet_();
  if (p.id) {
    var rdata = rsh.getDataRange().getValues();
    for (var i = 1; i < rdata.length; i++) {
      if (String(rdata[i][0]) === String(p.id)) {
        rsh.getRange(i + 1, 3, 1, 2).setValues([[String(p.name || ""), dataStr]]);
        return responseJSON({ status: "success", message: "تم تحديث التسكين", id: String(p.id) });
      }
    }
  }
  var id = "R" + new Date().getTime();
  rsh.appendRow([id, String(p.tripId), String(p.name || ""), dataStr, nowStr_()]);
  return responseJSON({ status: "success", message: "تم حفظ التسكين", id: id });
}

function deleteRooming(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "التسكين للمديرين فقط" });
  var rsh = roomingsSheet_();
  var rdata = rsh.getDataRange().getValues();
  for (var i = 1; i < rdata.length; i++) {
    if (String(rdata[i][0]) === String(p.roomingId)) {
      rsh.deleteRow(i + 1);
      return responseJSON({ status: "success", message: "تم حذف التسكين" });
    }
  }
  return responseJSON({ status: "error", message: "التسكين غير موجود" });
}

// ------------------------- v5.6: إدارة الإرسال التلقائي (واتساب - للمديرين) -------------------------
function waRecipientsSheet_() {
  return getOrCreateSheet_("WARecipients", ["ID", "Name", "Phone", "Status", "SentAt", "AddedAt"]);
}

function waTemplatesSheet_() {
  return getOrCreateSheet_("WATemplates", ["ID", "Type", "Text", "CreatedAt"]);
}

function getWA(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "إدارة الإرسال للمديرين فقط" });
  var recipients = [];
  var sh = waRecipientsSheet_();
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    recipients.push({
      id: String(d[i][0]),
      name: String(d[i][1] || ""),
      phone: String(d[i][2] || ""),
      status: String(d[i][3] || "pending"),
      sentAt: fmtDateTime_(d[i][4]),
      addedAt: fmtDateTime_(d[i][5])
    });
  }
  var templates = [];
  var th = waTemplatesSheet_();
  var t = th.getDataRange().getValues();
  for (var j = 1; j < t.length; j++) {
    if (!t[j][0]) continue;
    templates.push({ id: String(t[j][0]), type: String(t[j][1] || ""), text: String(t[j][2] || ""), createdAt: fmtDateTime_(t[j][3]) });
  }
  return responseJSON({ status: "success", recipients: recipients, templates: templates });
}

function addWARecipient(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "إدارة الإرسال للمديرين فقط" });
  var rawPhone = String(p.phone || "").trim();
  var digits = rawPhone.replace(/\D/g, "");
  if (digits.length < 8) return responseJSON({ status: "error", message: "رقم الموبايل غير صالح" });
  var sh = waRecipientsSheet_();
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][2]).replace(/\D/g, "") === digits) {
      return responseJSON({ status: "error", message: "الرقم موجود في قائمة الإرسال بالفعل" });
    }
  }
  sh.appendRow(["W" + new Date().getTime(), String(p.name || "").trim(), rawPhone, "pending", "", nowStr_()]);
  return responseJSON({ status: "success", message: "تمت إضافة الرقم" });
}

function deleteWARecipient(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "للمديرين فقط" });
  var sh = waRecipientsSheet_();
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(p.id)) {
      sh.deleteRow(i + 1);
      return responseJSON({ status: "success", message: "تم حذف الرقم" });
    }
  }
  return responseJSON({ status: "error", message: "الرقم غير موجود" });
}

function addWATemplate(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "للمديرين فقط" });
  var type = String(p.type || "").trim();
  var text = String(p.text || "").trim();
  if (type !== "intro" && type !== "body" && type !== "outro") return responseJSON({ status: "error", message: "نوع القالب غير صالح" });
  if (!text) return responseJSON({ status: "error", message: "نص القالب مطلوب" });
  waTemplatesSheet_().appendRow(["M" + new Date().getTime(), type, text, nowStr_()]);
  return responseJSON({ status: "success", message: "تم حفظ القالب" });
}

function deleteWATemplate(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "للمديرين فقط" });
  var th = waTemplatesSheet_();
  var t = th.getDataRange().getValues();
  for (var i = 1; i < t.length; i++) {
    if (String(t[i][0]) === String(p.id)) {
      th.deleteRow(i + 1);
      return responseJSON({ status: "success", message: "تم حذف القالب" });
    }
  }
  return responseJSON({ status: "error", message: "القالب غير موجود" });
}

function markWASent(p) {
  if (!isManager_(p.role)) return responseJSON({ status: "error", message: "للمديرين فقط" });
  var sh = waRecipientsSheet_();
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(p.id)) {
      sh.getRange(i + 1, 4, 1, 2).setValues([["sent", nowStr_()]]);
      return responseJSON({ status: "success" });
    }
  }
  return responseJSON({ status: "error", message: "الرقم غير موجود" });
}

// ------------------------- v5.7: سجل المتابعات اليومي -------------------------
function followupLogSheet_() {
  return getOrCreateSheet_("FollowupLog", ["ID", "Date", "Time", "ClientId", "ClientName", "Phone", "Program", "Members", "Status", "Feedback1", "Feedback2", "Feedback3", "Feedback4", "DoneBy", "DoneByCode"]);
}

function logFollowup_(clientId, p) {
  var sh = followupLogSheet_();
  var now = new Date();
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 2, 1, 2).setNumberFormat("@");
  sh.getRange(row, 6).setNumberFormat("@");
  var prog = (p.programName || "") + (p.turkeySubName ? " - " + p.turkeySubName : "");
  sh.getRange(row, 1, 1, 15).setValues([[
    "F" + now.getTime(),
    Utilities.formatDate(now, tz_(), "yyyy-MM-dd"),
    Utilities.formatDate(now, tz_(), "HH:mm"),
    String(clientId),
    p.name || "",
    p.phone || "",
    prog,
    Number(p.members || 1),
    p.status || "",
    p.fb1 || "", p.fb2 || "", p.fb3 || "", p.fb4 || "",
    p.editorName || p.agentName || "",
    p.agentCode || ""
  ]]);
}

function getFollowupLog(p) {
  var role = String(p.role || "").trim().toLowerCase();
  var code = String(p.code || "").trim().toLowerCase();
  var username = String(p.username || "").trim().toLowerCase();
  var dispName = String(p.name || "").trim().toLowerCase();
  var mgr = isManager_(role);
  var date = String(p.date || "").trim();
  var sh = followupLogSheet_();
  var d = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    if (date && fmtDate_(d[i][1]) !== date) continue;
    var by = String(d[i][13] || "").trim().toLowerCase();
    var byCode = String(d[i][14] || "").trim().toLowerCase();
    var mine = (by !== "" && (by === username || by === dispName)) || (byCode !== "" && byCode === code);
    if (!(mgr || mine)) continue;
    out.push({
      id: String(d[i][0]),
      date: fmtDate_(d[i][1]),
      time: fmtTime_(d[i][2]),
      clientId: String(d[i][3] || ""),
      name: String(d[i][4] || ""),
      phone: String(d[i][5] || ""),
      program: String(d[i][6] || ""),
      members: d[i][7] || 1,
      status: String(d[i][8] || ""),
      fb1: String(d[i][9] || ""), fb2: String(d[i][10] || ""), fb3: String(d[i][11] || ""), fb4: String(d[i][12] || ""),
      doneBy: String(d[i][13] || "")
    });
  }
  return responseJSON({ status: "success", date: date, entries: out });
}

// ------------------------- الموظفون (للمدير الفني فقط) -------------------------
function getUsers(userRole) {
  var role = String(userRole || "").trim().toLowerCase();
  if (role !== "tech_manager" && role !== "admin") {
    return responseJSON({ status: "error", message: "غير مصرح لك بإدارة الموظفين (للمدير الفني فقط)" });
  }
  var sheet = ss_().getSheetByName("Users");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Users غير موجودة" });

  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      users.push({
        username: String(data[i][0]),
        password: String(data[i][1]),
        code: String(data[i][2] || ""),
        role: String(data[i][3] || "mowazaf"),
        name: String(data[i][4] || data[i][0])
      });
    }
  }
  return responseJSON({ status: "success", users: users });
}

function saveUser(p) {
  var requesterRole = String(p.requesterRole || "").trim().toLowerCase();
  if (requesterRole !== "tech_manager") {
    return responseJSON({ status: "error", message: "عفواً! المدير الفني للسيستم هو الوحيد المصرح له بإضافة أو حذف الموظفين." });
  }
  var sheet = ss_().getSheetByName("Users");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Users غير موجودة" });

  var data = sheet.getDataRange().getValues();
  var username = String(p.username || "").trim();
  var password = String(p.password || "").trim();
  var code = String(p.code || "").trim();
  var role = String(p.role || "mowazaf").trim();
  var name = String(p.name || username).trim();

  if (!username || !password) return responseJSON({ status: "error", message: "اسم المستخدم وكلمة المرور مطلوبان" });

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[username, password, code, role, name]]);
      return responseJSON({ status: "success", message: "تم تحديث بيانات الموظف بنجاح" });
    }
  }

  sheet.appendRow([username, password, code, role, name]);
  return responseJSON({ status: "success", message: "تم إضافة الموظف الجديد بنجاح" });
}

function deleteUser(p) {
  var requesterRole = String(p.requesterRole || "").trim().toLowerCase();
  if (requesterRole !== "tech_manager") {
    return responseJSON({ status: "error", message: "عفواً! المدير الفني للسيستم هو الوحيد المصرح له بإضافة أو حذف الموظفين." });
  }
  var sheet = ss_().getSheetByName("Users");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة Users غير موجودة" });

  var data = sheet.getDataRange().getValues();
  var username = String(p.username || "").trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username) {
      sheet.deleteRow(i + 1);
      return responseJSON({ status: "success", message: "تم حذف الموظف بنجاح" });
    }
  }
  return responseJSON({ status: "error", message: "المستخدم غير موجود" });
}

// ------------------------- التنبيهات الإدارية -------------------------
function sendAdminNote(p) {
  appendNote_(p.fromUser || "المدير", p.toUser || "", p.clientId || "", p.clientName || "", p.noteText || "");
  return responseJSON({ status: "success", message: "تم إرسال التنبيه للموظف بنجاح" });
}

function markNoteRead(noteId) {
  var sheet = ss_().getSheetByName("AdminNotes");
  if (!sheet) return responseJSON({ status: "error", message: "ورقة AdminNotes غير موجودة" });
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(noteId)) {
      sheet.getRange(i + 1, 7).setValue("read");
      return responseJSON({ status: "success" });
    }
  }
  return responseJSON({ status: "error", message: "التنبيه غير موجود" });
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
