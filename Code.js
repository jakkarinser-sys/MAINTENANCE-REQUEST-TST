// ============================================================
//  REP_TST — Repair Request System
//  Google Apps Script (Code.gs)
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const CONFIG = {
  LINE_TOKEN:    "c5YJ9RAyJBvqBLilBeexCFR8qKkh8ZMhl/mNYfc5qtDj6e+PKVHk7CsO+w4nqHto0epCkSSUvud1Xg7c6NdWpjTCL6vRyl25ByP5rKd0y7gtkXGwSEM8NjeJovtvvFM2kNuTWMGzzt2x39ENjfJI0wdB04t89/1O/w1cDnyilFU=",
  SPREADSHEET_ID:"1srwLYPSZSf_WQTyMI9UWPSg1cuRwGWkvZssCM1c5oyE",
  LIFF_FORM_URL: "https://liff.line.me/2009412029-oo37Q1jc",
  LIFF_TECH_URL: "https://liff.line.me/2009412708-aE6M85sA",

  // Dropbox
  DROPBOX_TEMP_TOKEN:    "sl.u.AGV2kvUnGJ0VsLnPHuMn4zbxlrisFc19X6pxlWDGoi9BFWEeZYtjvlWg5Uo6zmfaz3QWqomB8iG1PhwrvYd2NZiQJejzT8qOJ83aHcbrnXKZEGnVxTPYDi60NZE8pzEBJPezc0Qa1h2UVDTKNnYtlkrWF7cBvRadp3KZ7sJPfqYmgyput8UM4DCOyk6sqyW4uT33gsXtXIE0Vq533uUIxx_mwsXpvlSfMgtd67E7G4R3TY8MTtzRFnVjYLUNl41bSgZPWOugKHoN2Qj8L84Aqk7law9ucRldzt0LsQ_5fyxrA9nOZ2r_WdYXacFRElz2wzkmEonPugFY8D9txXACcyWqNeiw2NsoNZQmiNFn7-_Ht7LX-Q4qJ-LirzWTp7eBm1BNjaJU6MsZPoE75FhdZf9yf9L1me2JG70Va24Qsd3EHdr5_wvVvh2YmMFYgcyMwVjcTJQ5WdHAzPIemUTs2kGR-m_jTPvqDp_vFGuhfbJBXr8KbSqFw3F5uq891wdtDVemH4FztAGoJuikF2EFnArTRpAQoeZGQuIwsXlnb0yQUZvhrGk9rCxcrGDBUco_id74fIiqmBKjY0JaN8Kbt5SpAnvLSpu3fyW_CLX8wUhzO1GIT5PrhYPRIGPkTbg6ZUqnEjJPLzzQ_Gz5HuuOUGIjV0jh4SmjlJrCyzLt3t2fVNyFpOMCG5Juh_-K-2ZIs3275kO4zgk78-r98bdAHY4d1spOb0jb1eBvAJtTfFDR5luYiUIJ5wX1E2ZpCC1L6I4MbD7G8lxg4vGyrL1UM9lq9E0pKsHOfKJpDDjZUn8spVuYyoMtjUCQmb09wPWbIqx6f1si8c1AcFCiakTRQz-4VsAx6jMghOXXeou6mnbX-dKp5Hnu-uamiEcEc8g7PERVsfAZnUg4K9cs2YtT7gKaa-5asu3VfPFsZkG_NHpOxdnk1fTtDRtyJq9v1lQNh0lypOi3nlmBh-Fc-qFfAlCajjhh3pT6TvO-TSEEpC8_tmyoe3PO9jZlWyDkaJgoGREixJBIUW8vJCkmXP9z_tPJmEVq_UcA9HQW9Nco7nb2qIbiQ4gVgbxigczWHV1MNLoeODKVyParrlWYuXxnr-PyQnRcvwX616cBhywhKKE4p0tcDrMOanyl2xOZuq5R7ffij5Uo-pUbmiGSrqAHYR5nJiRiJhwmQQ-ecBHdV3I8FOv4A-GXU_vBO4pqu6YkLUq1TO8oRxx0CG8DRmlonrnuP_H8eFNPANgo67O4vljswIJYAdakYSjCkeZhLE_thMQaB-Rcm9oLUNpo4ozJKL3oJdR3Z6BHmdgb-kAV65A8x5bkDvIWHskWEG1AjQut91EbDTPzMwZLcjablykwT8RB9R8zr1yQ4epaivr9BsD53e02g3SrRRky_5rFpiCa2kQ",
  DROPBOX_APP_KEY:       "8d3clb1pmfc3m2c",
  DROPBOX_APP_SECRET:    "i0apr7wpotxb3pz",
  DROPBOX_REFRESH_TOKEN: "swEpNwFWEWsAAAAAAAAAAZrVxOsHIffGz28Ihg7u7skap4bv_nylFQ5RT4aHN_0A",
  DROPBOX_FOLDER:        "/REP-TST",

  // Manager LINE ID แต่ละฝ่าย
  MANAGER_PRODUCTION:   "Ub271fa7b2e4e5f651f2122bee8f333bc",
  MANAGER_ENGINEERING:  "Ue7d684a7b7686776d84d76437bae018f",
  MANAGER_LOGISTICS:    "Ueb559022ba4fc263f85d944cde8823e6",

  // หัวหน้าทีมช่างแต่ละประเภทงาน
  ELEC_LEAD_ID:  "U6c6ff5afedcac4d06874988860e7e21c",
  MECH_LEAD_ID:  "Ue5397a3ab0e915ed2e2429e6d7a6d253",
  CRANE_LEAD_ID: "U6c6ff5afedcac4d06874988860e7e21c",

  // Sheet names
  SHEET_REPAIR:   "RepairOrders",
  SHEET_IMAGES:   "RepairImages",
  SHEET_EMPLOYEE: "Employees",
  SHEET_PARTS:    "Parts",
  SHEET_MACHINES: "Machines",
};

// ── COLUMN INDEX MAP (0-based) ───────────────────────────────
const REPAIR_COLS = {
  REP_NO:            0,
  CREATED_AT:        1,
  EMP_CODE:          2,
  EMP_NAME:          3,
  DEPT:              4,
  DIV:               5,
  MACHINE_ID:        6,
  MACHINE:           7,
  JOB_TYPE:          8,
  DETAIL:            9,
  MACHINE_STOP:      10,
  REQUESTER_LINE_ID: 11,
  STATUS:            12,
  APPROVED_BY:       13,
  TECH_CODE:         14,
  TECH_NAME:         15,
  START_TIME:        16,
  FINISH_TIME:       17,
  REPAIR_DETAIL:     18,
  PARTS_USED:        19,  // รูปแบบ: "สายพาน 2 เส้น, ลูกปืน 1 ลูก"
  TECH_NOTE:         20,
  INSPECT_RESULT:    21,
  REQUESTER_NOTE:    22,
  // ── คอลัมน์ใหม่ (เพิ่มต่อจากเดิม ไม่กระทบข้อมูลเก่า) ───
  APPROVED_AT:       23,  // เวลาที่อนุมัติ
  INSPECT_AT:        24,  // เวลาที่ตรวจรับ
  IMAGES_REPAIR:     25,  // ลิงค์รูปถ่ายจุดเสีย (คั่นด้วย ", ")
  IMAGES_INSPECT:    26,  // ลิงค์รูปถ่ายหลังตรวจรับ (คั่นด้วย ", ")
  TIME_WAIT_APPROVE: 27,  // นาที: CreatedAt → ApprovedAt
  TIME_WAIT_START:   28,  // นาที: ApprovedAt → StartTime
  TIME_REPAIR:       29,  // นาที: StartTime → FinishTime
  TIME_WAIT_INSPECT: 30,  // นาที: FinishTime → InspectAt
  TOTAL_TIME:        31,  // นาที: CreatedAt → InspectAt (ปิดงาน)
};

// ── TIME & FORMAT HELPERS ────────────────────────────────────

/** คำนวณนาทีระหว่าง 2 timestamp (format "yyyy-MM-dd HH:mm:ss") */
function calcMinutes(fromStr, toStr) {
  if (!fromStr || !toStr) return "";
  try {
    const a = new Date(String(fromStr).replace(/-/g, "/"));
    const b = new Date(String(toStr).replace(/-/g, "/"));
    const diff = Math.round((b - a) / 60000);
    return (isNaN(diff) || diff < 0) ? "" : diff;
  } catch(e) { return ""; }
}

/** แปลง partsUsed array → ข้อความอ่านง่าย
 *  input:  [{name:"สายพาน", qty:2, unit:"เส้น"}, ...]
 *  output: "สายพาน 2 เส้น, ลูกปืน 1 ลูก"
 */
function formatParts(partsArray) {
  if (!partsArray || !partsArray.length) return "-";
  return partsArray
    .filter(p => p && p.name)
    .map(p => `${p.name} ${p.qty || 1} ${p.unit || ""}`.trim())
    .join(", ") || "-";
}

/** เพิ่ม URL รูปภาพเข้า column ที่ระบุ (ต่อท้าย ถ้ามีอยู่แล้ว) */
function appendImageUrl(row, col, newUrl) {
  const sh  = getSheet(CONFIG.SHEET_REPAIR);
  const old = sh.getRange(row, col + 1).getValue();
  const updated = old ? `${old}, ${newUrl}` : newUrl;
  sh.getRange(row, col + 1).setValue(updated);
}

// ── DROPBOX HELPERS ──────────────────────────────────────────
function getDropboxAccessToken() {
  const res = UrlFetchApp.fetch("https://api.dropbox.com/oauth2/token", {
    method: "post",
    payload: {
      refresh_token: CONFIG.DROPBOX_REFRESH_TOKEN,
      grant_type:    "refresh_token",
      client_id:     CONFIG.DROPBOX_APP_KEY,
      client_secret: CONFIG.DROPBOX_APP_SECRET,
    },
  });
  return JSON.parse(res.getContentText()).access_token;
}

function uploadToDropbox(base64Data, filename, folder) {
  const token  = getDropboxAccessToken();
  const bytes  = Utilities.base64Decode(base64Data);
  const blob   = Utilities.newBlob(bytes, "image/jpeg", filename);
  const path   = `${folder}/${filename}`;
  const args   = JSON.stringify({ path, mode: "add", autorename: true });

  const res = UrlFetchApp.fetch("https://content.dropboxapi.com/2/files/upload", {
    method: "post",
    headers: {
      Authorization:    "Bearer " + token,
      "Dropbox-API-Arg": args,
      "Content-Type":   "application/octet-stream",
    },
    payload: blob.getBytes(),
  });
  const uploaded = JSON.parse(res.getContentText());

  const linkRes = UrlFetchApp.fetch("https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings", {
    method: "post",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify({ path: uploaded.path_display }),
  });
  const linkData = JSON.parse(linkRes.getContentText());
  return (linkData.url || linkData.error_summary || "").replace("?dl=0", "?raw=1");
}

function saveImageRecord(repNo, stage, imageUrl) {
  const sh  = getSheet(CONFIG.SHEET_IMAGES);
  const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
  sh.appendRow([repNo, stage, imageUrl, now]);
}

// ── SETUP: รันครั้งเดียวเพื่อสร้าง Sheets ───────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const SHEETS = [
    {
      name: "Employees",
      headers: ["EmpCode","Name","Dept","Div","Position","SupervisorLineId"],
      sample: [["E001","สมชาย ใจดี","ผลิต A","ผลิต","พนักงาน","U0000000000000000000000000000001"],
               ["E002","สมหญิง รักงาน","ผลิต A","ผลิต","หัวหน้างาน","U0000000000000000000000000000002"]],
    },
    {
      name: "Machines",
      headers: ["MachineId","MachineName"],
      sample: [["M001","เครื่องปั๊ม A1"],["M002","เครื่องกลึง B2"],["M003","เครนโรงงาน 1"]],
    },
    {
      name: "Parts",
      headers: ["PartId","PartName","Unit"],
      sample: [["P001","สายพาน","เส้น"],["P002","ลูกปืน","ลูก"],["P003","ฟิวส์ 10A","อัน"]],
    },
    {
      name: "RepairImages",
      headers: ["RepNo","Stage","ImageUrl","UploadedAt"],
      sample: [],
    },
    {
      name: "RepairOrders",
      headers: [
        "RepNo","CreatedAt","EmpCode","EmpName","Dept","Div",
        "MachineId","MachineName","JobType","Detail","MachineStop",
        "RequesterLineId","Status","ApprovedBy",
        "TechCode","TechName","StartTime","FinishTime",
        "RepairDetail","PartsUsed","TechNote",
        "InspectResult","RequesterNote",
        // ── คอลัมน์ใหม่ ───────────────────────────────────
        "ApprovedAt","InspectAt",
        "ImagesRepair","ImagesInspect",
        "TimeWaitApprove_min","TimeWaitStart_min",
        "TimeRepair_min","TimeWaitInspect_min","TotalTime_min"
      ],
      sample: [],
    },
  ];

  SHEETS.forEach(({ name, headers, sample }) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      Logger.log(`✅ สร้าง Sheet: ${name}`);
    } else {
      Logger.log(`⚠️ Sheet "${name}" มีอยู่แล้ว — ข้ามการสร้าง`);
      return;
    }
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground("#1565c0")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sh.setFrozenRows(1);
    if (sample.length > 0) {
      sh.getRange(2, 1, sample.length, headers.length).setValues(sample);
    }
    sh.autoResizeColumns(1, headers.length);
  });

  const def = ss.getSheetByName("Sheet1") || ss.getSheetByName("แผ่น1");
  if (def) { ss.deleteSheet(def); Logger.log("🗑️ ลบ Sheet default แล้ว"); }
  Logger.log("🎉 Setup เสร็จสมบูรณ์!");
}

// ── MAP ฝ่าย → Manager LINE ID ──────────────────────────────
function getManagerId(div) {
  const map = {
    "ผลิต":       CONFIG.MANAGER_PRODUCTION,
    "วิศวกรรม":   CONFIG.MANAGER_ENGINEERING,
    "โลจิสติกส์": CONFIG.MANAGER_LOGISTICS,
  };
  return map[div] || null;
}

// ── SPREADSHEET HELPERS ──────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);
}

function getEmployeeByCode(empCode) {
  const sh   = getSheet(CONFIG.SHEET_EMPLOYEE);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(empCode).trim()) {
      return {
        code:         data[i][0],
        name:         data[i][1],
        dept:         data[i][2],
        div:          data[i][3],
        pos:          data[i][4],
        supervisorId: data[i][5],
      };
    }
  }
  return null;
}

function getMachineList() {
  const sh   = getSheet(CONFIG.SHEET_MACHINES);
  const data = sh.getDataRange().getValues();
  return data.slice(1).map(r => ({ id: r[0], name: r[1] }));
}

function getRepairByNo(repNo) {
  const sh   = getSheet(CONFIG.SHEET_REPAIR);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === repNo) return { row: i + 1, data: data[i] };
  }
  return null;
}

function updateRepairRow(row, updates) {
  const sh = getSheet(CONFIG.SHEET_REPAIR);
  for (const [col, val] of Object.entries(updates)) {
    sh.getRange(row, parseInt(col) + 1).setValue(val);
  }
}

// ── REPAIR NUMBER GENERATOR ──────────────────────────────────
function genRepairNo() {
  const sh    = getSheet(CONFIG.SHEET_REPAIR);
  const today = new Date();
  const ymd   = Utilities.formatDate(today, "Asia/Bangkok", "yyyyMMdd");
  const lastRow = sh.getLastRow();
  let seq = 1;
  if (lastRow > 1) {
    const lastNo = sh.getRange(lastRow, 1).getValue();
    if (String(lastNo).startsWith("REP" + ymd)) {
      seq = parseInt(String(lastNo).slice(-3)) + 1;
    }
  }
  return `REP${ymd}${String(seq).padStart(3, "0")}`;
}

// ── LINE API HELPERS ─────────────────────────────────────────
function textMsg(text) {
  return { type: "text", text: String(text) };
}

function lineReply(replyToken, messages) {
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method:      "post",
    contentType: "application/json",
    headers:     { Authorization: "Bearer " + CONFIG.LINE_TOKEN },
    payload:     JSON.stringify({ replyToken, messages }),
  });
}

function linePush(to, messages) {
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method:      "post",
    contentType: "application/json",
    headers:     { Authorization: "Bearer " + CONFIG.LINE_TOKEN },
    payload:     JSON.stringify({ to, messages }),
  });
}

// ── FLEX MESSAGE HELPERS ─────────────────────────────────────
function fRow(label, value, valueColor) {
  return {
    type: "box",
    layout: "horizontal",
    spacing: "sm",
    contents: [
      { type: "text", text: label, size: "sm", color: "#888888", flex: 4, wrap: false },
      { type: "text", text: String(value || "-"), size: "sm", color: valueColor || "#111111", flex: 6, wrap: true, weight: "bold" },
    ],
  };
}

function fSep() {
  return { type: "separator", margin: "sm" };
}

// ── FLEX MESSAGE BUILDERS ─────────────────────────────────────

// 1. ใบแจ้งซ่อมใหม่ → หัวหน้างาน (ฝ่ายผลิต — มีปุ่มอนุมัติ/ไม่อนุมัติ)
function flexApproveMsg(repNo, empName, dept, machineName, jobType, machineStop, detail) {
  const stopText  = machineStop ? "⛔ หยุดเครื่อง" : "✅ ทำงานอยู่";
  const stopColor = machineStop ? "#FF4444"        : "#00C851";
  const detailStr = detail ? String(detail).substring(0, 80) : "-";
  return {
    type: "flex",
    altText: `📬 ใบแจ้งซ่อมใหม่ ${repNo} — รออนุมัติ`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#E65100" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: "📬  ใบแจ้งซ่อมใหม่", color: "#FFFFFF", weight: "bold", size: "lg" },
          { type: "text", text: repNo,                 color: "#FFDDC1", size: "sm",     margin: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
        contents: [
          fRow("👤 ผู้แจ้ง",    empName),
          fRow("🏭 แผนก",       dept),
          fRow("⚙️ เครื่อง",    machineName),
          fRow("🔧 ประเภทงาน",  jobType),
          fSep(),
          {
            type: "box", layout: "horizontal", margin: "sm",
            contents: [
              { type: "text", text: "🚦 สถานะเครื่อง", size: "sm", color: "#888888" },
              { type: "text", text: stopText, size: "sm", color: stopColor, weight: "bold", align: "end" },
            ],
          },
          fSep(),
          { type: "text", text: "📋 " + detailStr, size: "sm", color: "#555555", wrap: true, margin: "sm" },
        ],
      },
      footer: {
        type: "box", layout: "horizontal", spacing: "sm", paddingAll: "12px",
        contents: [
          {
            type: "button", style: "primary", color: "#2E7D32", height: "sm",
            action: { type: "message", label: "✅ อนุมัติ", text: `อนุมัติ ${repNo}` },
          },
          {
            type: "button", style: "primary", color: "#C62828", height: "sm",
            action: { type: "message", label: "❌ ไม่อนุมัติ", text: `ไม่อนุมัติ ${repNo}` },
          },
        ],
      },
    },
  };
}

// 2. งานซ่อมใหม่ → หัวหน้าทีมช่าง
function flexTechMsg(repNo, machineName, jobType, detail, machineStop, liffUrl) {
  const stopText  = machineStop ? "⛔ หยุดเครื่อง" : "✅ ทำงานอยู่";
  const stopColor = machineStop ? "#FF4444"        : "#00C851";
  return {
    type: "flex",
    altText: `🔧 งานซ่อมใหม่ ${repNo}`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#1565C0" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: "🔧  งานซ่อมใหม่", color: "#FFFFFF", weight: "bold", size: "lg" },
          { type: "text", text: repNo,               color: "#BBDEFB", size: "sm",     margin: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
        contents: [
          fRow("⚙️ เครื่อง",    machineName),
          fRow("🔧 ประเภทงาน",  jobType),
          fSep(),
          {
            type: "box", layout: "horizontal", margin: "sm",
            contents: [
              { type: "text", text: "🚦 สถานะเครื่อง", size: "sm", color: "#888888" },
              { type: "text", text: stopText, size: "sm", color: stopColor, weight: "bold", align: "end" },
            ],
          },
          fSep(),
          { type: "text", text: "📋 " + String(detail).substring(0, 100), size: "sm", color: "#555555", wrap: true, margin: "sm" },
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px",
        contents: [
          {
            type: "button", style: "primary", color: "#1565C0",
            action: { type: "uri", label: "🛠️ เปิดใบงานช่าง", uri: liffUrl },
          },
        ],
      },
    },
  };
}

// 3. ซ่อมเสร็จ → ผู้แจ้งซ่อม (รอตรวจรับ)
function flexDoneMsg(repNo, techName, repairDetail, liffUrl) {
  return {
    type: "flex",
    altText: `✅ งานซ่อม ${repNo} เสร็จแล้ว — กรุณาตรวจรับ`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#00897B" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: "✅  ซ่อมเสร็จแล้ว",  color: "#FFFFFF", weight: "bold", size: "lg" },
          { type: "text", text: repNo,                  color: "#B2DFDB", size: "sm",     margin: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
        contents: [
          fRow("👷 ช่างผู้ซ่อม", techName),
          fSep(),
          { type: "text", text: "📋 สรุปงาน",                                    size: "sm", color: "#888888", margin: "sm" },
          { type: "text", text: String(repairDetail).substring(0, 100),           size: "sm", color: "#333333", wrap: true },
          fSep(),
          { type: "text", text: "⏰ กรุณาตรวจรับงานซ่อมภายใน 24 ชั่วโมงครับ",  size: "xs", color: "#E65100", wrap: true, margin: "sm" },
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px",
        contents: [
          {
            type: "button", style: "primary", color: "#00897B",
            action: { type: "uri", label: "🔍 ตรวจรับการซ่อม", uri: liffUrl },
          },
        ],
      },
    },
  };
}

// 4. ไม่อนุมัติ → ผู้แจ้งซ่อม
function flexRejectMsg(repNo) {
  return {
    type: "flex",
    altText: `⚠️ ใบแจ้งซ่อม ${repNo} ไม่ผ่านการอนุมัติ`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#C62828" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: "⚠️  ไม่ผ่านการอนุมัติ", color: "#FFFFFF", weight: "bold", size: "lg" },
          { type: "text", text: repNo,                     color: "#FFCDD2", size: "sm",     margin: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
        contents: [
          { type: "text", text: "ใบแจ้งซ่อมของท่านไม่ได้รับการอนุมัติ",          size: "sm", color: "#333333", wrap: true },
          fSep(),
          { type: "text", text: "กรุณาติดต่อหัวหน้างานของท่านโดยตรงครับ",        size: "sm", color: "#888888", wrap: true, margin: "sm" },
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px",
        contents: [
          {
            type: "button", style: "secondary",
            action: { type: "message", label: "รับทราบ", text: "รับทราบ" },
          },
        ],
      },
    },
  };
}

// 5. ทั่วไป (ส่งใบแจ้งสำเร็จ / ดูสถานะ)
function flexLiffMsg(title, subtitle, btnLabel, url) {
  return {
    type: "flex",
    altText: title,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#1565C0" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: title, color: "#FFFFFF", weight: "bold", size: "md", wrap: true },
        ],
      },
      body: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: subtitle, size: "sm", color: "#555555", wrap: true },
        ],
      },
      footer: {
        type: "box", layout: "vertical", paddingAll: "12px",
        contents: [
          {
            type: "button", style: "primary", color: "#1565C0",
            action: { type: "uri", label: btnLabel, uri: url },
          },
        ],
      },
    },
  };
}

// 6. อัปเดตสถานะ → ส่งทุกครั้งที่สถานะใบแจ้งซ่อมเปลี่ยน
function flexStatusUpdateMsg(repNo, machineName, jobType, machineStop, detail, statusText, headerColor, footerAction) {
  const stopText  = machineStop ? "⛔ หยุดเครื่อง" : "✅ ทำงานอยู่";
  const stopColor = machineStop ? "#FF4444"        : "#00C851";
  const bubble = {
    type: "bubble",
    styles: { header: { backgroundColor: headerColor || "#37474F" } },
    header: {
      type: "box", layout: "vertical", paddingAll: "16px",
      contents: [
        { type: "text", text: "📢  อัปเดตสถานะใบแจ้งซ่อม", color: "#FFFFFF", weight: "bold", size: "lg" },
        { type: "text", text: repNo, color: "#CFD8DC", size: "sm", margin: "xs" },
      ],
    },
    body: {
      type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
      contents: [
        fRow("⚙️ เครื่อง",   machineName),
        fRow("🔧 ประเภทงาน", jobType),
        {
          type: "box", layout: "horizontal", margin: "sm",
          contents: [
            { type: "text", text: "🚦 สถานะเครื่อง", size: "sm", color: "#888888" },
            { type: "text", text: stopText, size: "sm", color: stopColor, weight: "bold", align: "end" },
          ],
        },
        fSep(),
        { type: "text", text: "📋 " + String(detail || "-").substring(0, 80), size: "sm", color: "#555555", wrap: true, margin: "sm" },
        fSep(),
        {
          type: "box", layout: "vertical", margin: "sm", paddingAll: "10px",
          backgroundColor: "#ECEFF1", cornerRadius: "8px",
          contents: [
            { type: "text", text: "🔔 อัปเดตการซ่อม", size: "xs", color: "#607D8B", weight: "bold" },
            { type: "text", text: statusText, size: "sm", color: headerColor || "#37474F", weight: "bold", wrap: true, margin: "xs" },
          ],
        },
      ],
    },
  };
  if (footerAction) {
    bubble.footer = {
      type: "box", layout: "vertical", paddingAll: "12px",
      contents: [
        { type: "button", style: "primary", color: headerColor || "#37474F", action: footerAction },
      ],
    };
  }
  return {
    type: "flex",
    altText: `📢 ${repNo} — ${statusText}`,
    contents: bubble,
  };
}

// 7. แจ้งให้ทราบ → หัวหน้า/ผู้จัดการ ฝ่ายอื่น (ไม่ต้องอนุมัติ)
function flexFYIMsg(repNo, empName, dept, machineName, jobType, machineStop, detail) {
  const stopText  = machineStop ? "⛔ หยุดเครื่อง" : "✅ ทำงานอยู่";
  const stopColor = machineStop ? "#FF4444"        : "#00C851";
  return {
    type: "flex",
    altText: `📋 ${empName} แจ้งซ่อม ${machineName} — ${repNo}`,
    contents: {
      type: "bubble",
      styles: { header: { backgroundColor: "#546E7A" } },
      header: {
        type: "box", layout: "vertical", paddingAll: "16px",
        contents: [
          { type: "text", text: "📋  แจ้งให้ทราบ — ใบแจ้งซ่อมใหม่", color: "#FFFFFF", weight: "bold", size: "lg" },
          { type: "text", text: repNo, color: "#CFD8DC", size: "sm", margin: "xs" },
        ],
      },
      body: {
        type: "box", layout: "vertical", spacing: "sm", paddingAll: "16px",
        contents: [
          fRow("👤 ผู้แจ้ง",    `${empName} (${dept})`),
          fRow("⚙️ เครื่อง",    machineName),
          fRow("🔧 ประเภทงาน",  jobType),
          {
            type: "box", layout: "horizontal", margin: "sm",
            contents: [
              { type: "text", text: "🚦 สถานะเครื่อง", size: "sm", color: "#888888" },
              { type: "text", text: stopText, size: "sm", color: stopColor, weight: "bold", align: "end" },
            ],
          },
          fSep(),
          { type: "text", text: "📋 " + String(detail || "-").substring(0, 80), size: "sm", color: "#555555", wrap: true, margin: "sm" },
          fSep(),
          {
            type: "box", layout: "vertical", margin: "sm", paddingAll: "10px",
            backgroundColor: "#ECEFF1", cornerRadius: "8px",
            contents: [
              { type: "text", text: "ℹ️ ระบบแจ้งทีมช่างโดยตรงแล้วครับ", size: "xs", color: "#546E7A", weight: "bold", wrap: true },
              { type: "text", text: "ไม่ต้องดำเนินการใดๆ เพียงรับทราบครับ", size: "xs", color: "#888888", wrap: true, margin: "xs" },
            ],
          },
        ],
      },
    },
  };
}

// ── WEBHOOK ENTRY POINT ──────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // LINE Webhook มี body.events เสมอ
    if (body.events) {
      body.events.forEach(handleEvent);
      return ContentService.createTextOutput("OK");
    }
    // LIFF form post
    return handleFormPost(body);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function handleEvent(ev) {
  if (ev.type !== "message" || ev.message.type !== "text") return;
  const text       = ev.message.text.trim();
  const userId     = ev.source.userId;
  const replyToken = ev.replyToken;

  // DEBUG: ดัก userId/GroupId
  if (text.toUpperCase() === "MYID") {
    const sourceType = ev.source.type;
    const groupId    = ev.source.groupId || ev.source.roomId || "-";
    lineReply(replyToken, [textMsg(
      `🪪 ข้อมูล ID\n` +
      `ประเภท: ${sourceType}\n` +
      `User ID: ${userId}` +
      (sourceType !== "user" ? `\nGroup ID: ${groupId}` : "")
    )]);
    return;
  }

  // ผู้แจ้งซ่อมพิมพ์ REP
  if (text.toUpperCase() === "REP") {
    lineReply(replyToken, [
      flexLiffMsg(
        "📋 แจ้งซ่อมเครื่องจักร",
        "กดปุ่มด้านล่างเพื่อกรอกรายละเอียดการแจ้งซ่อมครับ",
        "📝 กรอกใบแจ้งซ่อม",
        CONFIG.LIFF_FORM_URL + `?userId=${userId}`
      ),
    ]);
    return;
  }

  // หัวหน้างานอนุมัติ / ไม่อนุมัติ (เฉพาะฝ่ายผลิต)
  const approveMatch = text.match(/^(อนุมัติ|ไม่อนุมัติ)\s+(REP\d+)$/i);
  if (approveMatch) {
    handleApproval(userId, approveMatch[1], approveMatch[2], replyToken);
    return;
  }
}

function handleApproval(supervisorId, decision, repNo, replyToken) {
  const rec = getRepairByNo(repNo);
  if (!rec) {
    lineReply(replyToken, [textMsg(`❌ ไม่พบเลขที่ใบแจ้งซ่อม ${repNo}`)]);
    return;
  }

  const cols            = REPAIR_COLS;
  const requesterUserId = rec.data[cols.REQUESTER_LINE_ID];
  const jobType         = rec.data[cols.JOB_TYPE];

  const machineName = rec.data[cols.MACHINE];
  const machineStop = rec.data[cols.MACHINE_STOP];
  const detail      = rec.data[cols.DETAIL];

  const nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

  if (decision === "ไม่อนุมัติ") {
    const waitApprove = calcMinutes(rec.data[cols.CREATED_AT], nowStr);
    updateRepairRow(rec.row, {
      [cols.STATUS]:            "ไม่อนุมัติ",
      [cols.APPROVED_BY]:       supervisorId,
      [cols.APPROVED_AT]:       nowStr,
      [cols.TIME_WAIT_APPROVE]: waitApprove,
    });
    lineReply(replyToken, [textMsg(`✅ บันทึกการไม่อนุมัติ ${repNo} แล้วครับ`)]);
    // แจ้งผู้แจ้งซ่อม — รายละเอียดครบ
    linePush(requesterUserId, [
      flexStatusUpdateMsg(
        repNo, machineName, jobType, machineStop, detail,
        "❌ หัวหน้างานไม่อนุมัติใบแจ้งซ่อมนี้\nกรุณาติดต่อหัวหน้างานของท่านโดยตรงครับ",
        "#C62828"
      ),
    ]);
    return;
  }

  // อนุมัติ → ส่งต่อหัวหน้าทีมช่าง
  const waitApprove = calcMinutes(rec.data[cols.CREATED_AT], nowStr);
  updateRepairRow(rec.row, {
    [cols.STATUS]:            "อนุมัติ",
    [cols.APPROVED_BY]:       supervisorId,
    [cols.APPROVED_AT]:       nowStr,
    [cols.TIME_WAIT_APPROVE]: waitApprove,
  });
  lineReply(replyToken, [textMsg(`✅ อนุมัติ ${repNo} แล้ว กำลังแจ้งทีมช่างครับ`)]);

  const leadId  = jobType === "ไฟฟ้า"    ? CONFIG.ELEC_LEAD_ID
                : jobType === "เครื่องกล" ? CONFIG.MECH_LEAD_ID
                : CONFIG.CRANE_LEAD_ID;
  const techUrl = CONFIG.LIFF_TECH_URL + `?repNo=${repNo}`;

  // แจ้งหัวหน้าทีมช่าง
  linePush(leadId, [
    flexTechMsg(repNo, machineName, jobType, detail, machineStop, techUrl),
  ]);

  // แจ้งผู้แจ้งซ่อม — รายละเอียดครบ
  linePush(requesterUserId, [
    flexStatusUpdateMsg(
      repNo, machineName, jobType, machineStop, detail,
      "✅ หัวหน้างานอนุมัติแล้ว\nรอทีมช่างรับงานและเริ่มซ่อมครับ",
      "#2E7D32"
    ),
  ]);
}

// ── WEB API สำหรับ LIFF ──────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  if (action === "getEmployee") {
    const emp = getEmployeeByCode(e.parameter.code);
    return jsonOut(emp || { error: "ไม่พบรหัสพนักงาน" });
  }

  if (action === "getMachines") {
    return jsonOut(getMachineList());
  }

  if (action === "getRepair") {
    const rec = getRepairByNo(e.parameter.repNo);
    return jsonOut(rec ? rec.data : { error: "ไม่พบใบแจ้งซ่อม" });
  }

  if (action === "getMyRepairs") {
    const userId = e.parameter.userId || "";
    const sh     = getSheet(CONFIG.SHEET_REPAIR);
    const data   = sh.getDataRange().getValues();
    const result = data.slice(1)
      .filter(r => String(r[REPAIR_COLS.REQUESTER_LINE_ID]) === userId)
      .reverse()          // ล่าสุดขึ้นก่อน
      .slice(0, 30)
      .map(r => ({
        repNo:       r[REPAIR_COLS.REP_NO],
        createdAt:   r[REPAIR_COLS.CREATED_AT],
        machineName: r[REPAIR_COLS.MACHINE],
        jobType:     r[REPAIR_COLS.JOB_TYPE],
        status:      r[REPAIR_COLS.STATUS],
        techName:    r[REPAIR_COLS.TECH_NAME] || "",
      }));
    return jsonOut(result);
  }

  if (action === "getParts") {
    const sh   = getSheet(CONFIG.SHEET_PARTS);
    const data = sh.getDataRange().getValues();
    return jsonOut(data.slice(1).map(r => ({ id: r[0], name: r[1], unit: r[2] })));
  }

  return jsonOut({ error: "unknown action" });
}

// ── FORM HANDLER (POST จาก LIFF) ─────────────────────────────
function handleFormPost(payload) {
  const action = payload.action;

  // ── บันทึกใบแจ้งซ่อมใหม่ ──────────────────────────────────
  if (action === "submitRepair") {
    const repNo = genRepairNo();
    const now   = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const sh    = getSheet(CONFIG.SHEET_REPAIR);
    sh.appendRow([
      repNo,                    // 0  เลขที่ใบ
      now,                      // 1  วันที่แจ้ง
      payload.empCode,          // 2  รหัสพนักงาน
      payload.empName,          // 3  ชื่อ
      payload.dept,             // 4  แผนก
      payload.div,              // 5  ฝ่าย
      payload.machineId,        // 6  รหัสเครื่อง
      payload.machineName,      // 7  ชื่อเครื่อง
      payload.jobType,          // 8  ประเภทงาน
      payload.detail,           // 9  รายละเอียด
      payload.machineStop,      // 10 หยุดเครื่อง true/false
      payload.requesterLineId,  // 11 LINE userId ผู้แจ้ง
      "รอการอนุมัติ",           // 12 สถานะ
      "",                       // 13 อนุมัติโดย
      "", "", "", "",           // 14-17 ข้อมูลช่าง
      "", "", "",               // 18-20 ผลการซ่อม
      "",                       // 21 ผลตรวจรับ
      "",                       // 22 ข้อเสนอแนะผู้แจ้ง
    ]);

    const empData = getEmployeeByCode(payload.empCode);
    const isLeaderOrManager = ["หัวหน้างาน", "หัวหน้าทีม", "ผู้จัดการ", "หัวหน้าฝ่าย"].includes(empData?.pos);

    const leadId  = payload.jobType === "ไฟฟ้า"    ? CONFIG.ELEC_LEAD_ID
                  : payload.jobType === "เครื่องกล" ? CONFIG.MECH_LEAD_ID
                  : CONFIG.CRANE_LEAD_ID;
    const techUrl = CONFIG.LIFF_TECH_URL + `?repNo=${repNo}`;

    // ── ฝ่ายผลิต + ไม่ใช่หัวหน้า → ต้องผ่านการอนุมัติก่อน ──
    if (payload.div === "ผลิต" && !isLeaderOrManager) {
      // แจ้ง supervisor
      if (empData && empData.supervisorId) {
        linePush(empData.supervisorId, [
          flexApproveMsg(repNo, payload.empName, payload.dept, payload.machineName, payload.jobType, payload.machineStop, payload.detail),
        ]);
      }
      // แจ้ง manager (ถ้าไม่ใช่คนเดียวกับ supervisor)
      const managerId = getManagerId(payload.div);
      if (managerId && managerId !== empData?.supervisorId) {
        linePush(managerId, [
          flexApproveMsg(repNo, payload.empName, payload.dept, payload.machineName, payload.jobType, payload.machineStop, payload.detail),
        ]);
      }
      // แจ้งผู้แจ้งว่ารอการอนุมัติ — รายละเอียดครบ
      linePush(payload.requesterLineId, [
        flexStatusUpdateMsg(
          repNo, payload.machineName, payload.jobType, payload.machineStop, payload.detail,
          "📤 ส่งใบแจ้งซ่อมแล้ว\nระบบได้แจ้งหัวหน้างานของท่านแล้ว\n⏳ กรุณารอการอนุมัติครับ",
          "#E65100",
          { type: "uri", label: "📋 ดูสถานะใบแจ้งซ่อม", uri: CONFIG.LIFF_FORM_URL + `?action=status&userId=${payload.requesterLineId}` }
        ),
      ]);

    } else {
      // ── ฝ่ายอื่น หรือ เป็นหัวหน้า/ผู้จัดการ → ข้ามอนุมัติ ไปช่างได้เลย ──
      updateRepairRow(
        getRepairByNo(repNo).row,
        {
          [REPAIR_COLS.STATUS]:            "อนุมัติ",
          [REPAIR_COLS.APPROVED_BY]:       isLeaderOrManager ? payload.empName : "auto",
          [REPAIR_COLS.APPROVED_AT]:       now,   // อนุมัติทันทีพร้อมแจ้ง
          [REPAIR_COLS.TIME_WAIT_APPROVE]: 0,     // ไม่ต้องรออนุมัติ
        }
      );

      // แจ้งหัวหน้าทีมช่าง
      linePush(leadId, [
        flexTechMsg(repNo, payload.machineName, payload.jobType, payload.detail, payload.machineStop, techUrl),
      ]);

      // แจ้ง supervisor เพื่อทราบ (ไม่ต้องอนุมัติ)
      if (empData && empData.supervisorId && empData.supervisorId !== payload.requesterLineId) {
        linePush(empData.supervisorId, [
          flexFYIMsg(repNo, payload.empName, payload.dept, payload.machineName, payload.jobType, payload.machineStop, payload.detail),
        ]);
      }

      // แจ้ง manager เพื่อทราบ (ไม่ต้องอนุมัติ)
      const managerId = getManagerId(payload.div);
      if (managerId && managerId !== payload.requesterLineId && managerId !== empData?.supervisorId) {
        linePush(managerId, [
          flexFYIMsg(repNo, payload.empName, payload.dept, payload.machineName, payload.jobType, payload.machineStop, payload.detail),
        ]);
      }

      // แจ้งผู้แจ้งว่าส่งช่างแล้ว — รายละเอียดครบ
      linePush(payload.requesterLineId, [
        flexStatusUpdateMsg(
          repNo, payload.machineName, payload.jobType, payload.machineStop, payload.detail,
          "✅ ส่งใบแจ้งซ่อมสำเร็จ\nระบบได้แจ้งทีมช่างโดยตรงแล้วครับ",
          "#2E7D32",
          { type: "uri", label: "📋 ดูสถานะใบแจ้งซ่อม", uri: CONFIG.LIFF_FORM_URL + `?action=status&userId=${payload.requesterLineId}` }
        ),
      ]);
    }

    return jsonOut({ success: true, repNo });
  }

  // ── หัวหน้าช่างเพิ่ม/ลบรายชื่อช่าง ───────────────────────
  if (action === "updateTechs") {
    const rec = getRepairByNo(payload.repNo);
    if (!rec) return jsonOut({ error: "ไม่พบใบแจ้งซ่อม" });
    updateRepairRow(rec.row, {
      [REPAIR_COLS.TECH_CODE]: JSON.stringify(payload.techs.map(t => t.code)),
      [REPAIR_COLS.TECH_NAME]: payload.techs.map(t => t.name).join(", "),
    });
    return jsonOut({ success: true, techs: payload.techs });
  }

  // ── บันทึกเริ่มซ่อม ─────────────────────────────────────
  if (action === "startRepair") {
    const rec = getRepairByNo(payload.repNo);
    if (!rec) return jsonOut({ error: "ไม่พบใบแจ้งซ่อม" });
    const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    // คำนวณเวลารอตั้งแต่อนุมัติจนเริ่มซ่อม
    const approvedAt    = rec.data[REPAIR_COLS.APPROVED_AT];
    const waitStart     = calcMinutes(approvedAt || rec.data[REPAIR_COLS.CREATED_AT], now);
    updateRepairRow(rec.row, {
      [REPAIR_COLS.STATUS]:          "กำลังซ่อม",
      [REPAIR_COLS.START_TIME]:      now,
      [REPAIR_COLS.TIME_WAIT_START]: waitStart,
    });
    // แจ้งผู้แจ้งซ่อมว่าช่างเริ่มงานแล้ว
    const requesterLineId = rec.data[REPAIR_COLS.REQUESTER_LINE_ID];
    const techName        = rec.data[REPAIR_COLS.TECH_NAME] || "ทีมช่าง";
    linePush(requesterLineId, [
      flexStatusUpdateMsg(
        payload.repNo,
        rec.data[REPAIR_COLS.MACHINE],
        rec.data[REPAIR_COLS.JOB_TYPE],
        rec.data[REPAIR_COLS.MACHINE_STOP],
        rec.data[REPAIR_COLS.DETAIL],
        `🔧 ทีมช่าง (${techName}) เริ่มดำเนินการซ่อมแล้ว\nเวลาเริ่ม: ${now}`,
        "#1565C0"
      ),
    ]);
    return jsonOut({ success: true });
  }

  // ── ช่างบันทึกซ่อมเสร็จ ────────────────────────────────
  if (action === "finishRepair") {
    const rec = getRepairByNo(payload.repNo);
    if (!rec) return jsonOut({ error: "ไม่พบใบแจ้งซ่อม" });
    const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    // คำนวณเวลาซ่อม (StartTime → FinishTime)
    const startTime  = rec.data[REPAIR_COLS.START_TIME];
    const timeRepair = calcMinutes(startTime, now);
    // แปลง partsUsed → ข้อความอ่านง่าย
    const partsText  = formatParts(payload.partsUsed);
    updateRepairRow(rec.row, {
      [REPAIR_COLS.STATUS]:        "รอตรวจรับ",
      [REPAIR_COLS.FINISH_TIME]:   now,
      [REPAIR_COLS.REPAIR_DETAIL]: payload.repairDetail,
      [REPAIR_COLS.PARTS_USED]:    partsText,   // "สายพาน 2 เส้น, ลูกปืน 1 ลูก"
      [REPAIR_COLS.TECH_NOTE]:     payload.techNote,
      [REPAIR_COLS.TIME_REPAIR]:   timeRepair,
    });

    // แจ้งผู้แจ้งซ่อมให้ตรวจรับ — รายละเอียดครบ
    const requesterLineId = rec.data[REPAIR_COLS.REQUESTER_LINE_ID];
    const techName        = rec.data[REPAIR_COLS.TECH_NAME] || "ทีมช่าง";
    const liffInspectUrl  = CONFIG.LIFF_FORM_URL + `?action=inspect&repNo=${payload.repNo}&userId=${requesterLineId}`;
    linePush(requesterLineId, [
      flexStatusUpdateMsg(
        payload.repNo,
        rec.data[REPAIR_COLS.MACHINE],
        rec.data[REPAIR_COLS.JOB_TYPE],
        rec.data[REPAIR_COLS.MACHINE_STOP],
        rec.data[REPAIR_COLS.DETAIL],
        `✅ ทีมช่าง (${techName}) ซ่อมเสร็จแล้ว\nสรุปงาน: ${String(payload.repairDetail || "-").substring(0, 60)}\n⏰ กรุณาตรวจรับภายใน 24 ชั่วโมงครับ`,
        "#00897B",
        { type: "uri", label: "🔍 กดตรวจรับงานซ่อมที่นี่", uri: liffInspectUrl }
      ),
    ]);
    return jsonOut({ success: true });
  }

  // ── อัปโหลดรูปไป Dropbox ──────────────────────────────
  if (action === "uploadImage") {
    const folder = `${CONFIG.DROPBOX_FOLDER}/${payload.repNo}/${payload.stage}`;
    const url    = uploadToDropbox(payload.base64, payload.filename, folder);
    // บันทึกใน RepairImages sheet (log รายละเอียด)
    saveImageRecord(payload.repNo, payload.stage, url);
    // บันทึก URL ใน RepairOrders ด้วย (รวมในแถวเดียว)
    const rec = getRepairByNo(payload.repNo);
    if (rec) {
      const imgCol = payload.stage === "repair" ? REPAIR_COLS.IMAGES_REPAIR : REPAIR_COLS.IMAGES_INSPECT;
      appendImageUrl(rec.row, imgCol, url);
    }
    return jsonOut({ success: true, url });
  }

  // ── ผู้แจ้งตรวจรับ ─────────────────────────────────────
  if (action === "inspect") {
    const rec = getRepairByNo(payload.repNo);
    if (!rec) return jsonOut({ error: "ไม่พบใบแจ้งซ่อม" });
    const now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    // คำนวณเวลา: รอตรวจรับ (FinishTime → InspectAt) และ รวม (CreatedAt → InspectAt)
    const finishTime      = rec.data[REPAIR_COLS.FINISH_TIME];
    const createdAt       = rec.data[REPAIR_COLS.CREATED_AT];
    const timeWaitInspect = calcMinutes(finishTime, now);
    const totalTime       = calcMinutes(createdAt,  now);
    updateRepairRow(rec.row, {
      [REPAIR_COLS.STATUS]:            payload.pass ? "ปิดงาน"  : "ส่งซ่อมใหม่",
      [REPAIR_COLS.INSPECT_RESULT]:    payload.pass ? "ผ่าน"    : "ไม่ผ่าน",
      [REPAIR_COLS.REQUESTER_NOTE]:    payload.note || "",
      [REPAIR_COLS.INSPECT_AT]:        now,
      [REPAIR_COLS.TIME_WAIT_INSPECT]: timeWaitInspect,
      [REPAIR_COLS.TOTAL_TIME]:        totalTime,
    });

    // แจ้งทีมช่างว่าผู้แจ้งตรวจรับแล้ว
    const jobType = rec.data[REPAIR_COLS.JOB_TYPE];
    const leadId  = jobType === "ไฟฟ้า"    ? CONFIG.ELEC_LEAD_ID
                  : jobType === "เครื่องกล" ? CONFIG.MECH_LEAD_ID
                  : CONFIG.CRANE_LEAD_ID;
    const noteText = payload.note ? `\nหมายเหตุ: ${payload.note}` : "";

    if (payload.pass) {
      linePush(leadId, [
        flexStatusUpdateMsg(
          payload.repNo,
          rec.data[REPAIR_COLS.MACHINE],
          jobType,
          rec.data[REPAIR_COLS.MACHINE_STOP],
          rec.data[REPAIR_COLS.DETAIL],
          `🎉 ผู้แจ้งซ่อมตรวจรับ "ผ่าน" แล้ว\nงานปิดเรียบร้อยครับ${noteText}`,
          "#2E7D32"
        ),
      ]);
    } else {
      linePush(leadId, [
        flexStatusUpdateMsg(
          payload.repNo,
          rec.data[REPAIR_COLS.MACHINE],
          jobType,
          rec.data[REPAIR_COLS.MACHINE_STOP],
          rec.data[REPAIR_COLS.DETAIL],
          `🔄 ผู้แจ้งซ่อมตรวจรับ "ไม่ผ่าน"\nกรุณาดำเนินการซ่อมเพิ่มเติมครับ${noteText}`,
          "#E65100"
        ),
      ]);
    }
    return jsonOut({ success: true });
  }

  return jsonOut({ error: "unknown action" });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
