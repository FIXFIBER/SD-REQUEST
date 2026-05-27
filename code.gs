/**
 * ============================================================
 * FIBERONE SD REQUEST HUB — COMPLETE BACKEND v1.0
 * Google Apps Script for: index.html, auth.html, admin.html, employee.html
 *
 * SHEETS USED:
 *   SD_USERS       — All registered portal users
 *   SD_DEPARTMENTS — Company departments
 *   SD_REQUESTS    — All submitted requests
 *   SD_NEWS        — Employee dashboard news items
 *   SD_ALLOWLIST   — Admin allowlist
 *
 * SESSION SYSTEM:
 *   Uses PropertiesService (UserProperties) — no sheet needed.
 *   Each logged-in browser session stores a token key in UserProperties.
 *   Token → JSON payload (email, name, role, department, expires).
 *
 * HOW TO DEPLOY:
 *   1. Open your Google Spreadsheet
 *   2. Extensions → Apps Script → paste this entire file as Code.gs
 *   3. Also create these HTML files in the same project:
 *      index.html, auth.html, admin.html, employee.html
 *      (paste your existing HTML into each)
 *   4. In each HTML file, find the closing </body> tag and add:
 *      <script>var scriptUrl = "<?= scriptUrl ?>";</script>
 *      (this is already in your auth.html — add it to admin/employee too)
 *   5. Run "runFirstSetup" ONCE from the ⚙️ SD PORTAL ADMIN menu
 *   6. Deploy → New Deployment → Web App
 *      - Execute as: Me
 *      - Who has access: Anyone (even anonymous)
 *   7. Copy the Web App URL — it becomes scriptUrl automatically
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────
var SD = {
  USERS:     "SD_USERS",
  DEPTS:     "SD_DEPARTMENTS",
  PAT_PROJECTS: "SD_PAT_PROJECTS",
  ADMIN:    "admin",
  EMPLOYEE: "employee",
};

/**
 * Configuration Helper
 * Pulls settings from Script Properties or defaults to standard values.
 */
function _getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    TOKEN_TTL: parseInt(props.getProperty("TOKEN_TTL") || (8 * 60 * 60 * 1000)),
    SECRET_KEY: props.getProperty("SECRET_KEY") || "SD_SYSTEM_DEFAULT_SECRET",
    DOMAIN: props.getProperty("ALLOWED_DOMAIN") || "@fob.ng",
    NAVY: props.getProperty("THEME_COLOR") || "#0d1526",
  };
}

/**
 * Utility to update system configuration via UI
 */
function ui_configureSystem() {
  var ui = SpreadsheetApp.getUi();
  var config = _getConfig();
  
  var domain = ui.prompt("Set Allowed Email Domain (currently: " + config.DOMAIN + "):").getResponseText();
  var secret = ui.prompt("Set System Secret Key (for security, currently: " + config.SECRET_KEY + "):").getResponseText();
  
  if (domain) PropertiesService.getScriptProperties().setProperty("ALLOWED_DOMAIN", domain);

  if (secret) PropertiesService.getScriptProperties().setProperty("SECRET_KEY", secret);
  
  // Initialize empty defaults if they don't exist
  if (!PropertiesService.getScriptProperties().getProperty("DEFAULT_ADMIN_DEPT")) {
    PropertiesService.getScriptProperties().setProperty("DEFAULT_ADMIN_DEPT", "Technology Support");
  }
  // Add default request types if the sheet is empty
  if (!PropertiesService.getScriptProperties().getProperty("DEFAULT_REQUEST_TYPES")) {
    PropertiesService.getScriptProperties().setProperty("DEFAULT_REQUEST_TYPES", "Software Request,Automation Idea,Feedback,Bug Report,Other");
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Configuration updated successfully.");
}


// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("⚙️ SD PORTAL ADMIN")
      .addItem("🚀 Run First Setup",        "runFirstSetup")
      .addSeparator()
      .addItem("👑 Create Super Admin",     "ui_createSuperAdmin")
      .addItem("👤 Create Admin User",      "ui_createAdmin")
      .addItem("🔧 Configure System",       "ui_configureSystem")
      .addItem("🔑 Reset User Password",   "ui_resetPassword")
      .addItem("🗑️  Clear All Sessions",    "clearAllSessions")
      .addSeparator()
      .addItem("🧨 Reset & Wipe System",    "ui_wipeSystem")
      .addToUi();
  } catch (e) {
    // This error happens when running onOpen from the script editor.
    // The menu only appears when you refresh the Spreadsheet itself.
  }
}


// ─────────────────────────────────────────────────────────────
// FIRST SETUP
// Creates all sheets with headers. Seeds default data.
// Safe to re-run — will not overwrite existing sheets.
// ─────────────────────────────────────────────────────────────
function runFirstSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = _getConfig();

  // Cleanup legacy sheets no longer in use
  var legacySheets = ["SD_REQUESTS", "SD_NEWS", "SD_ALLOWLIST", "SD_MESSAGES", "SD_MAILS", "SD_REQUEST_TYPES", "SD_WORKFLOWS"];
  legacySheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.deleteSheet(sh);
    }
  });

  // 1. SD_USERS
  // UserID | Name | Email | PasswordHash | Salt | Role | Department | Gender | WorkflowNotes | Status | CreatedAt | LastLoginAt | AllowMessages | AIAutoAnswer
  var users = _sheet(SD.USERS, ["UserID","Name","Email","PasswordHash","Salt","Role","Department","Gender","WorkflowNotes","Status","CreatedAt","LastLoginAt","AllowMessages","AIAutoAnswer"]);
  _repairUserSheet(users);

  // 2. SD_DEPARTMENTS
  // DeptID | Name | HeadEmail | CreatedAt | CreatedBy
  var depts = _sheet(SD.DEPTS, ["DeptID","Name","HeadEmail","CreatedAt","CreatedBy"]);

  // 3. SD_PAT_PROJECTS
  var patSh = _sheet(SD.PAT_PROJECTS, [
    "ProjectID","ProjectName","SiteAddress","Lat","Lon","Phase",
    "WorkDescription","Vendor","InspectionDate","Department",
    "SubmittedBy","SubmittedAt","Status","SnagScore","Verdict",
    "ChecklistJSON","BOQJson","SnagJSON","SignoffJSON","PhotoNotes"
  ]);
  patSh.getRange(1,1,1,patSh.getLastColumn())
    .setBackground(config.NAVY).setFontColor("#fff").setFontWeight("bold");

  // Style all header rows navy
  [users, depts, patSh].forEach(function(sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setBackground(config.NAVY).setFontColor("#ffffff").setFontWeight("bold");
  });

  // Seed a default super-admin if no users exist yet
  if (users.getLastRow() === 1) {
    var adminDept = PropertiesService.getScriptProperties().getProperty("DEFAULT_ADMIN_DEPT") || "Technology Support";
    _createUserRow("Portal Admin", "admin" + config.DOMAIN, "admin123", SD.ADMIN, adminDept);
  }

  ss.toast("✅ SD Portal Setup Complete — v1.0 Ready", "Setup", 5);
}

/**
 * Standardizes the SD_USERS sheet to the 12-column format.
 * Fixes malformed/shifted data from older versions.
 */
function _repairUserSheet(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  var data = sh.getDataRange().getValues();
  var headers = ["UserID","Name","Email","PasswordHash","Salt","Role","Department","Gender","WorkflowNotes","Status","CreatedAt","LastLoginAt","AllowMessages","AIAutoAnswer"];
  var repaired = [headers];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r.filter(String).length === 0) continue; 

    var row = new Array(14).fill("");

    var emailIdx = -1;
    if (String(r[0]).indexOf("@") !== -1) emailIdx = 0;
    else if (String(r[1]).indexOf("@") !== -1) emailIdx = 1;
    else if (String(r[2]).indexOf("@") !== -1) emailIdx = 2;

    if (emailIdx === -1) continue; 

    if (String(r[0]).startsWith("USR-")) {
      for (var j = 0; j < Math.min(r.length, 14); j++) row[j] = r[j];
      if (!row[7]) row[7] = "Other";
      if (!row[9]) row[9] = "active";
      if (!row[12]) row[12] = "TRUE";
      if (!row[13]) row[13] = "FALSE";
    } 
    else if (emailIdx === 1) {
      row[0] = _genId("USR");
      row[1] = r[0]; // Name
      row[2] = r[1]; // Email
      row[3] = r[2]; // Hash
      row[4] = r[3]; // Salt
      row[5] = r[4]; // Role
      row[6] = r[5]; // Dept
      
      var val6 = String(r[6]).toLowerCase();
      if (val6 === "male" || val6 === "female") {
        row[7] = r[6];
        row[9] = r[7] || "active";
        row[10] = r[8] || "";
      } else {
        row[7] = "Other";
        row[9] = r[6] || "active";
        row[10] = r[7] || "";
        row[11] = r[8] || "";
        row[12] = "TRUE";
        row[13] = "FALSE";
      }
    }
    else {
       row[0] = _genId("USR");
       row[1] = r[emailIdx - 1] || "Unknown";
       row[2] = r[emailIdx];
       for (var k = 3; k < 14; k++) { if (r[k-1]) row[k] = r[k-1]; }
    }

    repaired.push(row);
  }
  
  sh.clear();
  sh.getRange(1, 1, repaired.length, 14).setValues(repaired);
  var config = _getConfig();
  sh.getRange(1, 1, 1, 14).setBackground(config.NAVY).setFontColor("#ffffff").setFontWeight("bold");
}

// Create a sheet only if it doesn't exist; returns the sheet either way
function _sheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    // Ensure getLastColumn() is not 0 for an existing sheet before getting range
    if (sh.getLastColumn() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      var curH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (curH.length < headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}


// ─────────────────────────────────────────────────────────────
// PASSWORD HELPERS
// ─────────────────────────────────────────────────────────────
function _salt() {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
      String(new Date().getTime()) + String(Math.random()))
  ).substring(0, 16);
}

function _hash(password, salt) {
  var secret = _getConfig().SECRET_KEY;
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      password + salt + secret)
  );
}


// ─────────────────────────────────────────────────────────────
// SESSION SYSTEM  (PropertiesService — UserProperties)
//
// Why UserProperties?
//   • Scoped to the Google account running the script —
//     each user gets isolated storage, no sheet required.
//   • Survives page refreshes; cleared on logout or expiry.
//   • 9KB per key is more than enough for our token payload.
//
// Token format stored as a JSON string:
//   { email, name, role, department, expires (ms timestamp) }
//
// The token itself is a random 32-char hex string that we
// use as the property key: "token_<hex>"
// ─────────────────────────────────────────────────────────────

function _makeToken() {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(Math.random()) + String(new Date().getTime()));
  return bytes.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('').substring(0, 32);
}

/**
 * Determines if a user is currently "active" based on last login.
 */
function _isUserActive(lastLoginIso) {
  if (!lastLoginIso) return false;
  try {
    var last = new Date(lastLoginIso).getTime();
    var now = new Date().getTime();
    return (now - last) < (10 * 60 * 1000); // 10 minute active threshold
  } catch(e) { return false; }
}

/**
 * Create a session. Stores the payload in UserProperties.
 * Returns the token string (sent to the browser via sessionStorage).
 */
function _createSession(email, name, role, department) {
  console.log("Creating session for: " + email);
  var lock = LockService.getUserLock();
  try {
    lock.waitLock(10000); // 10 second timeout
    var token   = _makeToken();
    var config  = _getConfig();
    var payload = JSON.stringify({
      email:      email,
      name:       name,
      role:       role,
      department: department,
      expires:    new Date().getTime() + config.TOKEN_TTL,
    });
    PropertiesService.getUserProperties().setProperty("tok_" + token, payload);
  } catch (e) {
    throw new Error("Unable to create session: " + e.message);
  } finally {
    lock.releaseLock();
  }
  return token;
}

/**
 * Validate a token. Returns the session object or throws.
 */
function _session(token) {
  if (!token) throw new Error("Not logged in. Please sign in again.");
  var raw = PropertiesService.getUserProperties().getProperty("tok_" + token);
  if (!raw) throw new Error("Session not found. Please sign in again.");
  var s = JSON.parse(raw);
  if (new Date().getTime() > s.expires) {
    PropertiesService.getUserProperties().deleteProperty("tok_" + token);
    throw new Error("Session expired. Please sign in again.");
  }
  return s;
}

/**
 * Validate token AND require a specific role.
 */
function _adminSession(token) {
  var s = _session(token);
  var role = String(s.role || "").toLowerCase();
  if (role !== SD.ADMIN && role !== "super admin") throw new Error("Admin access required.");
  return s;
}

/**
 * Validate token AND require Super Admin role.
 */
function _superAdminSession(token) {
  var s = _session(token);
  var role = String(s.role || "").toLowerCase();
  if (role !== "super admin") throw new Error("Super Admin access required.");
  return s;
}

/**
 * Delete a session (logout).
 */
function _destroySession(token) {
  if (token) PropertiesService.getUserProperties().deleteProperty("tok_" + token);
}

/**
 * Menu action — wipes all "tok_*" keys from UserProperties.
 * Useful for testing or locked-out admins.
 */
function clearAllSessions() {
  var props = PropertiesService.getUserProperties().getProperties();
  Object.keys(props).forEach(function(k) {
    if (k.indexOf("tok_") === 0) PropertiesService.getUserProperties().deleteProperty(k);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("All sessions cleared.", "Sessions", 3);
}

// ─────────────────────────────────────────────────────────────
// ID / REF GENERATORS
// ─────────────────────────────────────────────────────────────
function _genId(prefix) {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var rand  = "";
  for (var i = 0; i < 6; i++) rand += chars[Math.floor(Math.random() * chars.length)];
  return prefix + "-" + rand;
}

// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  AUTH  —  called from auth.html via google.script.run
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────

/**
 * Sign up a new employee.
 * @param {string} name
 * @param {string} email     must end with @fob.ng
 * @param {string} password  min 4 chars
 * @param {string} department
 * @param {string} gender
 * @returns {{ success, message? }}
 */
/**
 * Admin-only: Creates a user with a specific role and gender.
 */
function adminCreateUser(token, name, email, password, department, role, gender) {
  try {
    _superAdminSession(token);
    var config = _getConfig();
    email = String(email || "").toLowerCase().trim();
    role = String(role || SD.EMPLOYEE).toLowerCase().trim();
    gender = String(gender || "Other").trim();

    if (!name || !email || !password || !department)
      throw new Error("Name, email, password, and department are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup = data.slice(1).find(function(r){ 
      return String(r[2] || "").toLowerCase().trim() === String(email).toLowerCase().trim(); 
    });
    if (dup) throw new Error("An account with this email already exists.");

    _createUserRow(name, email, password, role, department, gender, "TRUE", "FALSE");

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function registerUser(name, email, password, department, gender) {
  var lock = LockService.getScriptLock();
  try {
    var config = _getConfig();
    email = String(email || "").toLowerCase().trim();
    if (!name || !email || !password || !department || !gender)
      throw new Error("All fields are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");
    if (String(password).length < 4)
      throw new Error("Password must be at least 4 characters.");

    lock.waitLock(15000);

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup  = data.slice(1).find(function(r){ 
      return String(r[2] || "").toLowerCase().trim() === String(email).toLowerCase().trim(); 
    });
    if (dup) throw new Error("An account with this email already exists.");

    _createUserRow(name, email, password, SD.EMPLOYEE, department, gender);
    
    lock.releaseLock();
    return { success: true };
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: e.message };
  }
}

function _createUserRow(name, email, password, role, department, gender, allowMsgs, aiAuto) {
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var s    = _salt();
  var uid  = _genId("USR"); // Ensure unique ID
  sh.appendRow([uid, name, email.toLowerCase(), _hash(password, s), s, role, department, gender, "", "active", new Date().toISOString(), "", allowMsgs || "TRUE", aiAuto || "FALSE"]);
  return uid;
}

/**
 * Log in.
 * @param {string} email
 * @param {string} password
 * @returns {{ success, token?, role?, name?, department?, message? }}
 */
function loginUser(email, password) {
  try {
    console.log("Login attempt: " + email);
    email = String(email).toLowerCase().trim();
    if (!email || !password) throw new Error("Email and password are required.");

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var lastRow = sh.getLastRow();
    if (!sh || lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    // cols: UserID(0) Name(1) Email(2) Hash(3) Salt(4) Role(5) Dept(6) Gender(7) Notes(8) Status(9) CreatedAt(10) LastLogin(11) Allow(12) AI(13)

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (String(r[2] || "").toLowerCase().trim() !== email) continue;
      if (r[9] === "banned") throw new Error("Your account has been suspended. Contact your admin.");
      if (_hash(String(password), String(r[4])) !== String(r[3])) throw new Error("Incorrect password.");

      sh.getRange(i + 2, 12).setValue(new Date().toISOString()); // Update login time
      var token = _createSession(email, String(r[1]), String(r[5]), String(r[6]));
      return { 
        success: true, token: token, role: r[5], name: r[1], department: r[5] === "super admin" ? "N/A" : r[6], 
        gender: r[7], notes: r[8], allowMessages: r[12] === "TRUE", aiAutoAnswer: r[13] === "TRUE" 
      };
    }
    throw new Error("No account found with that email address.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Log out — destroys the session.
 * @param {string} token
 */
function logoutUser(token) {
  _destroySession(token);
  return { success: true };
}

/**
 * Get department names for the signup dropdown.
 * Public — no token needed.
 * @returns {{ success, data: string[] }}
 */
function getPublicDepartments() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("public_departments");
  if (cached) return { success: true, data: JSON.parse(cached), cached: true };

  try {
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    var names = data.map(function(r){ return String(r[0]).trim(); })
                    .filter(Boolean).sort();
    
    cache.put("public_departments", JSON.stringify(names), 600); // Cache for 10 mins (better UX)
    return { success: true, data: names };
  } catch(e) {
    return { success: false, data: [], message: e.message };
  }
}

/**
 * Get the allowed email domain from configuration.
 * Public — no token needed.
 * @returns {{ success, domain: string }}
 */
function getDomainConfig() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("domain_config");
  if (cached) return { success: true, ...JSON.parse(cached), cached: true };

  try {
    var config = _getConfig();
    var data = { domain: config.DOMAIN };
    cache.put("domain_config", JSON.stringify(data), 1500);
    return { success: true, ...data };
  } catch(e) {
    return { success: false, domain: "@fob.ng", message: e.message };
  }
}


// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  ADMIN  —  all require a valid admin token (first argument)
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────

// ── ALL REQUESTS + STATS ──────────────────────────────────────

// ── DEPARTMENTS ───────────────────────────────────────────────

/**
 * List departments with live request + user counts.
 * @param {string} token  admin token
 */
function getDepartments(token) {
  try {
    _adminSession(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var deptSh = ss.getSheetByName(SD.DEPTS);
    var userSh = ss.getSheetByName(SD.USERS);

    var deptData = deptSh.getLastRow() > 1 ? deptSh.getDataRange().getValues().slice(1) : [];
    var userData = userSh.getLastRow() > 1 ? userSh.getDataRange().getValues().slice(1) : [];

    // Count active users per department name
    var userMap = {};
    userData.forEach(function(r){
      if (r[7]==="active"){ var d=String(r[6]).trim(); userMap[d]=(userMap[d]||0)+1; }
    });

    var departments = deptData.map(function(r) {
      return {
        deptId:       r[0],
        name:         r[1],
        headEmail:    r[2],
        createdAt:    r[3],
        activeUsers:  userMap[r[1]] || 0,
      };
    });
    return { success: true, departments: departments };
  } catch(e) {
    return { success: false, message: e.message, departments: [] };
  }
}

/**
 * Add a department.
 * @param {string} token
 * @param {string} name
 * @param {string} headEmail  (optional)
 */
function addDepartment(token, name, headEmail) {
  try {
    var sess = _adminSession(token);
    name = String(name || "").trim();
    if (!name) throw new Error("Department name is required.");

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues().slice(1);
    var dup  = data.find(function(r){ return String(r[1]).toLowerCase()===name.toLowerCase(); });
    if (dup) throw new Error("A department with the name '" + name + "' already exists.");

    var deptId = _genId("DEPT");
    sh.appendRow([deptId, name, headEmail||"", new Date().toISOString(), sess.email]);
    CacheService.getScriptCache().remove("public_departments"); // Invalidate cache for signup
    return { success: true, deptId: deptId, message: "Department added." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update a department's name and/or head email.
 * @param {string} token
 * @param {string} deptId
 * @param {string} name
 * @param {string} headEmail
 */
function updateDepartment(token, deptId, name, headEmail) {
  try {
    _adminSession(token);
    name = String(name || "").trim();
    if (!name) throw new Error("Department name is required.");

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.getRange(i+1, 2).setValue(name);
        sh.getRange(i+1, 3).setValue(headEmail || "");
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department updated." };
      }
    }
    throw new Error("Department not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete a department.
 * @param {string} token
 * @param {string} deptId
 */
function deleteDepartment(token, deptId) {
  try {
    _adminSession(token);
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.deleteRow(i + 1);
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department deleted." };
      }
    }
    throw new Error("Department not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ── USERS ─────────────────────────────────────────────────────

/**
 * List all registered users.
 * @param {string} token  admin token
 */
function getUsers(token) {
  try {
    _adminSession(token);
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues().slice(1);
    var users = data.map(function(r) {
      return {
        userId:      r[0], name:        r[1],
        email:       r[2], role:        r[5],
        department:  r[5] === "super admin" ? "N/A" : r[6],
        gender:      r[7], workflowNotes: r[8],
        status:      r[9], createdAt:   r[10], lastLoginAt: r[11],
      };
    });
    return { success: true, users: users };
  } catch(e) {
    return { success: false, message: e.message, users: [] };
  }
}

/**
 * Promote or demote a user's role.
 * @param {string} token
 * @param {string} userId
 * @param {string} newRole  "admin" | "employee"
 */
function updateUserRole(token, userId, newRole) {
  try {
    _superAdminSession(token);
    if (newRole !== SD.ADMIN && newRole !== SD.EMPLOYEE && newRole !== "super admin")
      throw new Error("Role must be 'admin', 'employee', or 'super admin'.");

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sh      = ss.getSheetByName(SD.USERS);
    var data    = sh.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sh.getRange(i+1, 6).setValue(newRole);
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Activate or ban a user account.
 * @param {string} token
 * @param {string} userId
 * @param {string} status  "active" | "banned"
 */
function updateUserStatus(token, userId, status) {
  var lock = LockService.getScriptLock();
  try {
    _superAdminSession(token);
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sh.getRange(i+1, 8).setValue(status);
        lock.releaseLock();
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: e.message };
  }
}

/**
 * Permanently deletes a user from the sheet.
 */
function adminDeleteUser(token, userId) {
  try {
    _superAdminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sh.deleteRow(i + 1);
        return { success: true, message: "User deleted from sheet." };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Edits user details in the sheet.
 */
function adminUpdateUser(token, userId, updates) {
  try {
    _superAdminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        if (updates.name) sh.getRange(i+1, 2).setValue(updates.name);
        if (updates.email) sh.getRange(i+1, 3).setValue(updates.email);
        if (updates.role) sh.getRange(i+1, 6).setValue(updates.role);
        if (updates.department) sh.getRange(i+1, 7).setValue(updates.department);
        if (updates.gender) sh.getRange(i+1, 8).setValue(updates.gender);
        return { success: true, message: "User updated in sheet." };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ── PROFILE SETTINGS (admin + employee) ──────────────────────

/**
 * Update the logged-in user's display name.
 * Token identifies who is updating — email must match session.
 *
 * @param {string} token
 * @param {string} email    the user's own email (must match session)
 * @param {string} newName  new display name
 */
function updateUserProfile(token, email, newName) {
  try {
    var sess = _session(token);
    if (sess.email !== String(email).toLowerCase().trim())
      throw new Error("You can only update your own profile.");
    newName = String(newName || "").trim();
    if (!newName) throw new Error("Name cannot be empty.");

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || "").toLowerCase().trim() === String(sess.email || "").toLowerCase().trim()) {
        sh.getRange(i+1, 2).setValue(newName);
        // Update the stored session payload so the nav reflects the change immediately
        var raw = PropertiesService.getUserProperties().getProperty("tok_" + token);
        if (raw) {
          var s = JSON.parse(raw);
          s.name = newName;
          PropertiesService.getUserProperties().setProperty("tok_" + token, JSON.stringify(s));
        }
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  EMPLOYEE  —  called from employee.html via google.script.run
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────
/**
 * Gets members of the same department (for employees) or all members (for admins).
 */
function getDepartmentMembers(token) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, members: [] };
    
    var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var members = data.filter(function(r) {
      var role = String(sess.role || "").toLowerCase();
      if (role === 'admin' || role === 'super admin') return true;
      return String(r[6]) === sess.department;
    }).map(function(r) {
      return { name: r[1], email: r[2], role: String(r[5] || "").toLowerCase() };
    });
    
    return { success: true, members: members };
  } catch(e) {
    return { success: false, message: e.message, members: [] };
  }
}

function handleRequest(e) {
  console.log("Request received: " + JSON.stringify(e.parameter));
  var action, args;
  var responseHeaders = { "Access-Control-Allow-Origin": "*" }; // Prepping for advanced CORS
  
  if (e.postData && e.postData.contents) {
    try {
      var payload = JSON.parse(e.postData.contents);
      action = payload.action;
      args = payload.args || [];
    } catch(err) {
      return _json({ success: false, message: "API ERROR: Invalid JSON payload" });
    }
  } else {
    action = e.parameter.action;
    try {
      args = e.parameter.data ? JSON.parse(e.parameter.data) : [];
    } catch(err) {
      args = [];
    }
  }

  if (!action) return _json({ success: false, message: "API ERROR: No action specified." });

  try {
    var result;
    switch(action) {
      case "loginUser": result = loginUser.apply(null, args); break;
      case "registerUser": result = registerUser.apply(null, args); break;
      case "getDepartments": result = getDepartments.apply(null, args); break;
      case "getUsers": result = getUsers.apply(null, args); break;
      case "getDepartmentMembers": result = getDepartmentMembers.apply(null, args); break;
      case "adminCreateUser": result = adminCreateUser.apply(null, args); break;
      case "updateUserProfile": result = updateUserProfile.apply(null, args); break;
      case "adminDeleteUser": result = adminDeleteUser.apply(null, args); break;
      case "adminUpdateUser": result = adminUpdateUser.apply(null, args); break;
      case "savePATProject": result = savePATProject.apply(null, args); break;
      case "getPATProjects": result = getPATProjects.apply(null, args); break;
      case "getPATProjectById": result = getPATProjectById.apply(null, args); break;
      case "addDepartment": result = addDepartment.apply(null, args); break;
      case "updateDepartment": result = updateDepartment.apply(null, args); break;
      case "deleteDepartment": result = deleteDepartment.apply(null, args); break;
      case "updateUserRole": result = updateUserRole.apply(null, args); break;
      case "updateUserStatus": result = updateUserStatus.apply(null, args); break;
      case "getPublicDepartments": result = getPublicDepartments.apply(null, args); break;
      case "resetUserPassword": result = resetUserPassword.apply(null, args); break;
      case "getDomainConfig": result = getDomainConfig.apply(null, args); break;
      default: throw new Error("Unknown action: " + action);
    }
    return _json(result);
  } catch (err) {
    return _json({ success: false, message: "SYSTEM CRASH: " + err.message });
  }
}

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ─────────────────────────────────────────────────────────────
// ADMIN MENU HELPERS  (run from the spreadsheet UI)
// ─────────────────────────────────────────────────────────────

/**
 * UI Helper to create a Super Admin.
 */
function ui_createSuperAdmin() {
  _manualCreateUser("super admin");
}

/**
 * UI Helper to create a standard Admin.
 */
function ui_createAdmin() {
  _manualCreateUser(SD.ADMIN);
}

/**
 * Internal logic for creating privileged users from the Spreadsheet UI.
 */
function _manualCreateUser(role) {
  var ui = SpreadsheetApp.getUi();
  var config = _getConfig();
  var email = ui.prompt("New " + role + " email (must end with " + config.DOMAIN + "):").getResponseText().trim();
  if (!email || !email.endsWith(config.DOMAIN)) { ui.alert("Invalid domain."); return; }
  var name  = ui.prompt("Full name:").getResponseText().trim();
  var pass  = ui.prompt("Password (min 4 chars):").getResponseText();
  
  var deptsRes = getPublicDepartments();
  var dept = "";
  if (role !== "super admin" && deptsRes.success && deptsRes.data.length > 0) {
    var deptMsg = "Available Departments:\n" + deptsRes.data.join(", ") + "\n\nEnter Department:";
    dept = ui.prompt(deptMsg).getResponseText().trim();
  } else if (role !== "super admin") {
    dept = ui.prompt("No departments found. Enter a new department name:").getResponseText().trim();
  } else { dept = "N/A"; }
  
  if (role !== "super admin" && !dept) { ui.alert("Department is required."); return; }

  try {
    var res = registerUser(name, email, pass, dept, "Other"); // Default gender for manual admin
    if (!res.success) throw new Error(res.message);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh     = ss.getSheetByName(SD.USERS);
    var data   = sh.getDataRange().getValues();
    var lo     = email.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === lo) { sh.getRange(i+1,6).setValue(role); break; }
    }
    ui.alert("✅ " + role + " created!\n\nEmail: " + email);
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

/**
 * Securely resets a user's password. Restricted to Super Admins.
 */
function resetUserPassword(token, targetEmail, newPassword) {
  try {
    _superAdminSession(token);
    targetEmail = String(targetEmail || "").toLowerCase().trim();
    if (!targetEmail || !newPassword) throw new Error("Email and new password are required.");
    if (newPassword.length < 4) throw new Error("Password must be at least 4 characters.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase().trim() === targetEmail) {
        var s = _salt();
        sh.getRange(i + 1, 4).setValue(_hash(newPassword, s));
        sh.getRange(i + 1, 5).setValue(s);
        return { success: true, message: "Password reset successfully for " + targetEmail };
      }
    }
    throw new Error("User not found: " + targetEmail);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function ui_resetPassword() {
  var ui    = SpreadsheetApp.getUi();
  var email = ui.prompt("Email of user to reset:").getResponseText().trim().toLowerCase();
  if (!email) return;
  var pass  = ui.prompt("New password (min 4 chars):").getResponseText();
  if (!pass || pass.length < 4) { ui.alert("❌ Password too short."); return; }

  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) {
      var s = _salt();
      sh.getRange(i+1, 4).setValue(_hash(pass, s));
      sh.getRange(i+1, 5).setValue(s);
      ui.alert("✅ Password reset for: " + email);
      return;
    }
  }
  ui.alert("❌ User not found: " + email);
}

function ui_wipeSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("🚨 DANGER: SYSTEM WIPE", 
    "This will delete ALL Users and Departments. \n\nAre you absolutely sure you want to proceed?", 
    ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [SD.USERS, SD.DEPTS];
    sheets.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 1) {
        sh.deleteRows(2, sh.getLastRow() - 1);
      }
    });
    PropertiesService.getScriptProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();
    ss.toast("System wiped successfully. Run 'First Setup' to initialize again.", "Reset Complete");
  }
}

// ── PAT PROJECTS ──────────────────────────────────────────────

function savePATProject(token, projectData) {
  try {
    var sess = _session(token);
    // Only MEC team (or admin) can save
    var dept = String(sess.department || "").toLowerCase();
    var role = String(sess.role || "").toLowerCase();
    if (dept !== "mec" && dept !== "mech" && role !== "admin" && role !== "super admin")
      throw new Error("Only MEC department can save PAT projects.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SD_PAT_PROJECTS");
    var projectId = _genId("PAT");
    var now = new Date().toISOString();

    sh.appendRow([
      projectId,
      projectData.projectName || "",
      projectData.siteAddress || "",
      projectData.lat || "",
      projectData.lon || "",
      projectData.phase || "",
      projectData.workDescription || "",
      projectData.vendor || "",
      projectData.inspectionDate || "",
      sess.department,
      sess.name,
      now,
      projectData.status || "Pending",
      projectData.snagScore || 0,
      projectData.verdict || "",
      JSON.stringify(projectData.checklist || {}),
      JSON.stringify(projectData.boq || []),
      JSON.stringify(projectData.snags || []),
      JSON.stringify(projectData.signoff || {}),
      projectData.photoNotes || ""
    ]);

    return { success: true, projectId: projectId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getPATProjects(token) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SD_PAT_PROJECTS");
    if (!sh || sh.getLastRow() < 2) return { success: true, projects: [] };

    var data = sh.getDataRange().getValues().slice(1);
    var projects = data.map(function(r) {
      return {
        projectId: r[0], projectName: r[1], siteAddress: r[2],
        lat: r[3], lon: r[4], phase: r[5], workDescription: r[6],
        vendor: r[7], inspectionDate: r[8], department: r[9],
        submittedBy: r[10], submittedAt: r[11], status: r[12],
        snagScore: r[13], verdict: r[14]
        // checklist/boq/snag/signoff JSON intentionally excluded from list view (too heavy)
      };
    });
    return { success: true, projects: projects };
  } catch(e) {
    return { success: false, message: e.message, projects: [] };
  }
}

function getPATProjectById(token, projectId) {
  try {
    _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SD_PAT_PROJECTS");
    if (!sh || sh.getLastRow() < 2) throw new Error("No projects found.");

    var data = sh.getDataRange().getValues().slice(1);
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === projectId) {
        var r = data[i];
        return {
          success: true,
          project: {
            projectId: r[0], projectName: r[1], siteAddress: r[2],
            lat: r[3], lon: r[4], phase: r[5], workDescription: r[6],
            vendor: r[7], inspectionDate: r[8], department: r[9],
            submittedBy: r[10], submittedAt: r[11], status: r[12],
            snagScore: r[13], verdict: r[14],
            checklist: JSON.parse(r[15] || "{}"),
            boq: JSON.parse(r[16] || "[]"),
            snags: JSON.parse(r[17] || "[]"),
            signoff: JSON.parse(r[18] || "{}"),
            photoNotes: r[19]
          }
        };
      }
    }
    throw new Error("Project not found: " + projectId);
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// END OF FILE
// ─────────────────────────────────────────────────────────────