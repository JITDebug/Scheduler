// ============================================================
//  SCHEDULER — Google Apps Script
//  Paste this entire file into Extensions ▸ Apps Script
//  then run  setupScheduler()  once from the toolbar (▶).
// ============================================================

const CONFIG = {
  schedulerSheetName: "Schedule",
  ministersSheetName: "Ministers",
  roleGroupsSheetName: "Role Groups",
  rolesSheetName: "Roles",
  vacationsSheetName: "Vacations",
  eventsSheetName: "Events",
  // Starting Friday – change to whatever date you need.
  // Format: year, month (0‑based!), day
  startFriday: new Date(2026, 4, 8),   // May 8 2026 (Friday)
  numberOfWeeks: 26,                    // how many Friday columns to generate (≈6 months)

  // Seed roles — only used during initial Full Setup.
  // After setup, edit the "Roles" sheet directly.
  seedRoles: [
    "Cooking",
    "Uniform",
    "Preacher",
    "Presider",
    "Worship Leader 1",
    "Worship Leader 2",
    "Backup Singer 1",
    "Backup Singer 2",
    "Backup Singer 3",
    "Prayer for Newcomer",
    "Breaking of the Bread",
    "Tithes and Offering",
    "Closing Prayer",
    "Guitarist (Acoustic)",
    "Guitarist (Electric)",
    "Guitarist (Bass)",
    "Drummer",
    "Keyboardist",
    "Mixer",
    "Chord Arranger",
    "Projectionist",
    "Cameraman",
    "Usher 1 (Lead)",
    "Usher 2",
    "Usher 3",
    "Usher 4",
    "Teacher (Kids)",
    "Teacher (Pre-teens)",
    "Co-teacher 1 (Kids)",
    "Co-teacher 2 (Kids)"
  ],

  // Role Groups – each group contains a subset of roles.
  // When a group is selected, only its roles are visible on the Schedule.
  roleGroups: {
    "Pulpit Ministry": [
      "Preacher",
      "Presider",
      "Prayer for Newcomer",
      "Breaking of the Bread",
      "Tithes and Offering",
      "Closing Prayer"
    ],
    "Music Ministry": [
      "Worship Leader 1",
      "Worship Leader 2",
      "Backup Singer 1",
      "Backup Singer 2",
      "Backup Singer 3",
      "Guitarist (Acoustic)",
      "Guitarist (Electric)",
      "Guitarist (Bass)",
      "Drummer",
      "Keyboardist",
      "Mixer",
      "Chord Arranger",
      "Projectionist",
      "Cameraman"
    ],
    "Children Ministry": [
      "Teacher (Kids)",
      "Teacher (Pre-teens)",
      "Co-teacher 1 (Kids)",
      "Co-teacher 2 (Kids)"
    ],
    "Ushering Ministry": [
      "Usher 1 (Lead)",
      "Usher 2",
      "Usher 3",
      "Usher 4"
    ],
    "Communion": [
      "Communion Server 1",
      "Communion Server 2",
      "Communion Server 3",
      "Communion Server 4"
    ],
    "General Announcement" : [
      "Cooking",
      "Uniform",
      "Preacher",
      "Presider",
      "Worship Leader 1",
      "Worship Leader 2",
      "Backup Singer 1",
      "Backup Singer 2",
      "Backup Singer 3",
      "Tithes and Offering",
      "Projectionist",
      "Cameraman"
    ],
    "Others": [
      "Uniform",
      "Cooking"
    ]
  },

  // Sample ministers with their eligible roles.
  // Edit the Ministers sheet later to add/remove people.
  sampleMinisters: [
    { name: "David", roles: ["Worship Leader 1", "Worship Leader 2"] },
    { name: "Jewell", roles: ["Worship Leader 1", "Worship Leader 2"] },
    { name: "Barong", roles: ["Uniform"] },
    { name: "Red", roles: ["Uniform"] },
    { name: "Blue", roles: ["Uniform"] },
    { name: "Green", roles: ["Uniform"] },
    { name: "Yellow", roles: ["Uniform"] },
    { name: "Pink", roles: ["Uniform"] },
    { name: "BS (Jerold)", roles: ["Cooking"] },
    { name: "BS (Caryl)", roles: ["Cooking"] },
    { name: "BS (Beth)", roles: ["Cooking"] },
    { name: "BS (Ailene)", roles: ["Cooking"] },
    { name: "BS (Clifford)", roles: ["Cooking"] }
  ],

  // URLs for header and footer images (PNG/JPG)
  // You can replace these with your own URLs
  printHeaderUrl: "https://via.placeholder.com/1000x150.png?text=CHURCH+MONTHLY+SCHEDULE",
  printFooterUrl: "https://via.placeholder.com/1000x100.png?text=GOD+BLESS+OUR+MINISTRY"
};

// ─── MENU ───────────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const groupMenu = ui.createMenu("📂 Filter by Group");

  // Dynamically build group submenu from the Role Groups sheet (or CONFIG fallback)
  const groupNames = getGroupNames_();
  groupNames.forEach((name, i) => {
    // Each group gets its own menu function via a dispatcher
    groupMenu.addItem(name, "filterGroup_" + i);
    if (name === "General Announcement") {
      groupMenu.addSeparator();
    }
  });
  groupMenu.addSeparator();
  groupMenu.addItem("👁 Show All Roles", "showAllRoles");

  ui.createMenu("⚙️ Configs")
    .addItem("🎉 Create/Repair Events Sheet", "initializeEventsSheet")
    .addItem("🏖 Create/Repair Vacations Sheet", "initializeVacationsSheet")
    .addItem("🛠 Create/Repair Settings Sheet", "initializeSettingsSheet")
    .addSeparator()
    .addItem("⚙️ Full Setup (WARNING: DELETES ALL DATA)", "setupScheduler")
    .addToUi();

  ui.createMenu("⏱ Scheduler")
    .addSubMenu(groupMenu)
    .addSeparator()
    .addItem("🖨 Export for Printing",        "exportForPrinting")
    .addItem("🤍 Clear Formatting",           "clearFormatting")
    .addItem("🎨 Re‑apply Formatting",        "applyFormatting")
    .addItem("⚠️ Check Conflicts",            "checkConflicts")
    .addSeparator()
    .addItem("➕ Add Month",                   "addMonth")
    .addItem("📦 Archive Month...",           "showArchiveDialog")
    .addItem("📦 Unarchive Month...",         "showUnarchiveDialog")
    .addSeparator()
    .addItem("👥 Update Minister Dropdowns",    "refreshAllDropdowns")
    .addItem("🔄 Sync New/Removed Roles",      "syncRoles")
    .addItem("🎉 Sync New/Removed Events",     "syncEvents")
    .addToUi();
}

// ─── MAIN SETUP ─────────────────────────────────────────────
function setupScheduler() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Setup Scheduler",
    "This will create / overwrite the Schedule, Ministers, Roles, Vacations, Events, Settings, and Role Groups sheets.\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Create or clear sheets (clear() does NOT remove data validation, so we do it explicitly)
  let rolesSheet = ss.getSheetByName(CONFIG.rolesSheetName);
  if (rolesSheet) { rolesSheet.getRange(1, 1, rolesSheet.getMaxRows(), rolesSheet.getMaxColumns()).clearDataValidations(); rolesSheet.clear(); }
  else rolesSheet = ss.insertSheet(CONFIG.rolesSheetName);

  let scheduleSheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (scheduleSheet) { scheduleSheet.getRange(1, 1, scheduleSheet.getMaxRows(), scheduleSheet.getMaxColumns()).clearDataValidations(); scheduleSheet.clear(); }
  else scheduleSheet = ss.insertSheet(CONFIG.schedulerSheetName);

  let ministersSheet = ss.getSheetByName(CONFIG.ministersSheetName);
  if (ministersSheet) { ministersSheet.getRange(1, 1, ministersSheet.getMaxRows(), ministersSheet.getMaxColumns()).clearDataValidations(); ministersSheet.clear(); }
  else ministersSheet = ss.insertSheet(CONFIG.ministersSheetName);

  let roleGroupsSheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  if (roleGroupsSheet) { roleGroupsSheet.getRange(1, 1, roleGroupsSheet.getMaxRows(), roleGroupsSheet.getMaxColumns()).clearDataValidations(); roleGroupsSheet.clear(); }
  else roleGroupsSheet = ss.insertSheet(CONFIG.roleGroupsSheetName);

  let settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet) { settingsSheet.clear(); }
  else settingsSheet = ss.insertSheet("Settings");

  let vacationsSheet = ss.getSheetByName(CONFIG.vacationsSheetName);
  if (vacationsSheet) { vacationsSheet.getRange(1, 1, vacationsSheet.getMaxRows(), vacationsSheet.getMaxColumns()).clearDataValidations(); vacationsSheet.clear(); }
  else vacationsSheet = ss.insertSheet(CONFIG.vacationsSheetName);

  let eventsSheet = ss.getSheetByName(CONFIG.eventsSheetName);
  if (eventsSheet) { eventsSheet.getRange(1, 1, eventsSheet.getMaxRows(), eventsSheet.getMaxColumns()).clearDataValidations(); eventsSheet.clear(); }
  else eventsSheet = ss.insertSheet(CONFIG.eventsSheetName);

  // 2. Build Settings sheet
  buildSettingsSheet_(settingsSheet);

  // 2.5 Build Vacations sheet
  buildVacationsSheet_(vacationsSheet);

  // 2.6 Build Events sheet
  buildEventsSheet_(eventsSheet);

  // 3. Build Roles sheet first (source of truth)
  buildRolesSheet_(rolesSheet);
  SpreadsheetApp.flush();  // Ensure roles data is written before other sheets read it

  // 3. Build Role Groups sheet
  buildRoleGroupsSheet_(roleGroupsSheet);

  // 4. Build Ministers sheet (referenced by Schedule)
  buildMinistersSheet_(ministersSheet);

  // 5. Build Schedule sheet
  buildScheduleSheet_(scheduleSheet);
  applyEventHeadersToSchedule_();

  // 6. Apply formatting
  applyFormatting();

  // 7. Set up dropdown validations
  refreshAllDropdowns();

  // 8. Move Schedule sheet to first position
  ss.setActiveSheet(scheduleSheet);
  ss.moveActiveSheet(1);

  // 9. Install edit trigger (only once)
  installTrigger_();

  ui.alert(
    "✅ Scheduler is ready!\n\n" +
    "• Edit the Roles sheet to add/remove positions.\n" +
    "• Edit the Ministers sheet to manage people and role eligibility.\n" +
    "• Edit the Role Groups sheet to organize roles into groups.\n" +
    "• Edit the Vacations sheet to track unavailabilities.\n" +
    "• Edit the Events sheet to highlight schedule date headers.\n" +
    "• Run Scheduler ▸ Sync Roles after changing the Roles sheet.\n" +
    "• Dropdowns auto‑update when you change the Ministers sheet."
  );
}

// ─── ROLES HELPER ───────────────────────────────────────────
/**
 * Read the list of roles from the Roles sheet.
 * Falls back to CONFIG.seedRoles if the sheet doesn't exist yet.
 */
function getRoles_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.rolesSheetName);
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return data
      .map(r => String(r[0]).trim())
      .filter(v => v !== "" && !v.startsWith("ℹ️"));
  }
  return CONFIG.seedRoles;
}

// ─── BUILD ROLES SHEET ──────────────────────────────────────
function buildRolesSheet_(sheet) {
  const roles = CONFIG.seedRoles;

  // Header
  sheet.appendRow(["Role Name"]);

  // Role rows
  roles.forEach(role => sheet.appendRow([role]));

  // Format header
  sheet.getRange(1, 1)
    .setFontWeight("bold")
    .setBackground("#0d47a1")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  // Format role cells
  const roleRange = sheet.getRange(2, 1, roles.length, 1);
  roleRange
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // Column width
  sheet.setColumnWidth(1, 200);

  // Freeze header
  sheet.setFrozenRows(1);

  // Instructions
  const instrRow = roles.length + 3;
  sheet.getRange(instrRow, 1)
    .setValue("ℹ️ Add or remove roles here. Each row is one role/position.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange(instrRow + 1, 1)
    .setValue("ℹ️ After editing, run Scheduler ▸ Sync Roles to update all sheets.")
    .setFontColor("#666666")
    .setFontStyle("italic");
}

// ─── SYNC ROLES ─────────────────────────────────────────────
// Reads the Roles sheet and propagates changes to the Ministers
// sheet (adds/removes columns) and Schedule sheet (adds/removes rows).
function syncRoles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolesSheet     = ss.getSheetByName(CONFIG.rolesSheetName);
  const ministersSheet = ss.getSheetByName(CONFIG.ministersSheetName);
  const scheduleSheet  = ss.getSheetByName(CONFIG.schedulerSheetName);

  if (!rolesSheet || !ministersSheet || !scheduleSheet) {
    SpreadsheetApp.getUi().alert("Cannot find required sheets. Run Full Setup first.");
    return;
  }

  const newRoles = getRoles_();
  if (newRoles.length === 0) {
    SpreadsheetApp.getUi().alert("No roles found in the Roles sheet. Please add at least one role.");
    return;
  }

  // ── Sync Ministers sheet (columns) ──
  const mData = ministersSheet.getDataRange().getValues();
  const mHeader = mData[0]; // ["Name", "Role1", "Role2", ...]
  const currentMinisterRoles = mHeader.slice(1); // roles currently as columns

  // Find roles to add and roles to remove
  const rolesToAdd    = newRoles.filter(r => !currentMinisterRoles.includes(r));
  const rolesToRemove = currentMinisterRoles.filter(r => !newRoles.includes(r));

  // Add new role columns to Ministers
  rolesToAdd.forEach(role => {
    const newCol = ministersSheet.getLastColumn() + 1;
    ministersSheet.getRange(1, newCol).setValue(role)
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    ministersSheet.setColumnWidth(newCol, 120);

    // Add checkboxes for existing minister rows
    const numMinisters = mData.slice(1).filter(row =>
      row[0] !== "" && typeof row[0] === "string" && !row[0].startsWith("ℹ️")
    ).length;
    if (numMinisters > 0) {
      const cbRange = ministersSheet.getRange(2, newCol, numMinisters, 1);
      cbRange.setDataValidation(
        SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()
      );
      cbRange.setValue(false);
      cbRange.setHorizontalAlignment("center");
    }
  });

  // Remove old role columns from Ministers (iterate right-to-left)
  if (rolesToRemove.length > 0) {
    // Re-read header after adds
    const updatedHeader = ministersSheet.getRange(1, 1, 1, ministersSheet.getLastColumn()).getValues()[0];
    for (let i = updatedHeader.length - 1; i >= 1; i--) {
      if (rolesToRemove.includes(updatedHeader[i])) {
        ministersSheet.deleteColumn(i + 1);
      }
    }
  }

  // ── Sync Schedule sheet (rows) ──
  const sData = scheduleSheet.getDataRange().getValues();
  const currentScheduleRoles = sData.slice(2).map(row => row[0]);
  const numCols = sData[0].length;

  // Add new role rows to Schedule
  const sRolesToAdd = newRoles.filter(r => !currentScheduleRoles.includes(r));
  sRolesToAdd.forEach(role => {
    const newRow = scheduleSheet.getLastRow() + 1;
    scheduleSheet.getRange(newRow, 1).setValue(role)
      .setFontWeight("bold")
      .setBackground("#e8f0fe")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    scheduleSheet.setRowHeight(newRow, 32);
  });

  // Remove old role rows from Schedule (iterate bottom-up)
  const sRolesToRemove = currentScheduleRoles.filter(r => !newRoles.includes(r));
  if (sRolesToRemove.length > 0) {
    // Re-read after adds
    const updatedData = scheduleSheet.getDataRange().getValues();
    for (let r = updatedData.length - 1; r >= 2; r--) {
      if (sRolesToRemove.includes(updatedData[r][0])) {
        scheduleSheet.deleteRow(r + 1);
      }
    }
  }

  // ── Update Role Groups dropdown validation ──
  const rgSheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  if (rgSheet && rgSheet.getLastRow() > 1 && rgSheet.getLastColumn() > 0) {
    // Clear old validation first, then re-apply only to the data area
    rgSheet.getDataRange().clearDataValidations();
    const rgDataCols = rgSheet.getLastColumn();
    const rgRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(newRoles, true)
      .setAllowInvalid(true)   // lenient — avoids errors if list is long
      .build();
    rgSheet.getRange(2, 1, newRoles.length, rgDataCols)
      .setDataValidation(rgRule);
  }

  // Refresh dropdowns and formatting
  refreshAllDropdowns();
  applyFormatting();

  const added = rolesToAdd.length + sRolesToAdd.length;
  const removed = rolesToRemove.length + sRolesToRemove.length;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Roles synced ✅  Added: " + rolesToAdd.length + " | Removed: " + rolesToRemove.length,
    "Scheduler", 5
  );
}

// ─── BUILD ROLE GROUPS SHEET ────────────────────────────────
function buildRoleGroupsSheet_(sheet) {
  const groups = CONFIG.roleGroups;
  const groupNames = Object.keys(groups);

  // Find the max number of roles in any group (for row count)
  let maxRoles = 0;
  groupNames.forEach(g => { maxRoles = Math.max(maxRoles, groups[g].length); });

  // Header row: group names across the top
  sheet.appendRow(groupNames);

  // Role rows under each group
  for (let r = 0; r < maxRoles; r++) {
    const row = [];
    groupNames.forEach(g => {
      row.push(groups[g][r] || "");
    });
    sheet.appendRow(row);
  }

  // Format header
  const headerRange = sheet.getRange(1, 1, 1, groupNames.length);
  headerRange
    .setFontWeight("bold")
    .setBackground("#7b1fa2")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  // Format data area
  if (maxRoles > 0) {
    const dataRange = sheet.getRange(2, 1, maxRoles, groupNames.length);
    dataRange.setHorizontalAlignment("center");

    // Add data validation — dropdown helper for role names.
    // setAllowInvalid(true) avoids errors when the role list is long
    // (Google Sheets can truncate requireValueInList with many items).
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(getRoles_(), true)
      .setAllowInvalid(true)
      .build();
    dataRange.setDataValidation(roleRule);
  }

  // Column widths
  if (groupNames.length > 0) {
    sheet.setColumnWidths(1, groupNames.length, 160);
  }

  // Freeze header
  sheet.setFrozenRows(1);

  // Instructions
  const instrRow = maxRoles + 3;
  sheet.getRange(instrRow, 1)
    .setValue("ℹ️ Each column is a group. List the role names under each group header.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange(instrRow + 1, 1)
    .setValue("ℹ️ Use Scheduler ▸ Filter by Group to show only the selected group's roles on the Schedule.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange(instrRow + 2, 1)
    .setValue("ℹ️ You can add new groups by adding new columns. Role names must match exactly.")
    .setFontColor("#666666")
    .setFontStyle("italic");
}

// ─── BUILD MINISTERS SHEET ───────────────────────────────────
function buildMinistersSheet_(sheet) {
  const roles = getRoles_();
  const ministers = CONFIG.sampleMinisters;

  // Header row: Name | Role1 | Role2 | ...
  const header = ["Name", ...roles];
  sheet.appendRow(header);

  // Minister rows
  ministers.forEach(person => {
    const row = [person.name];
    roles.forEach(role => {
      row.push(person.roles.includes(role));  // TRUE / FALSE
    });
    sheet.appendRow(row);
  });

  // Format header
  const headerRange = sheet.getRange(1, 1, 1, header.length);
  headerRange
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  // Freeze header row and Name column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Auto‑resize Name column
  sheet.autoResizeColumn(1);

  // Set checkbox data validation for role columns
  const numMinisters = ministers.length;
  if (numMinisters > 0) {
    const checkboxRange = sheet.getRange(2, 2, numMinisters, roles.length);
    const cbRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
    checkboxRange.setDataValidation(cbRule);
    checkboxRange.setHorizontalAlignment("center");
  }

  // Column widths for role columns
  if (roles.length > 0) {
    sheet.setColumnWidths(2, roles.length, 120);
  }

  // Add instructions note
  sheet.getRange(numMinisters + 3, 1)
    .setValue("ℹ️ Check the boxes to indicate which roles each minister can fill.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange(numMinisters + 4, 1)
    .setValue("ℹ️ Add new ministers by adding rows above this note. Then run Scheduler ▸ Refresh Dropdowns.")
    .setFontColor("#666666")
    .setFontStyle("italic");

  // Protect header row
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(protection => {
    if (protection.getDescription() === "Ministers sheet structure") {
      protection.remove();
    }
  });
  const protection = sheet.protect().setDescription("Ministers sheet structure");
  protection.setUnprotectedRanges([sheet.getRange(2, 1, sheet.getMaxRows() - 1, roles.length + 1)]);
  protection.setWarningOnly(true);
}

// ─── BUILD SCHEDULE SHEET ───────────────────────────────────
function buildScheduleSheet_(sheet) {
  const roles = getRoles_();

  // Generate Friday dates
  const fridays = generateFridays_(CONFIG.startFriday, CONFIG.numberOfWeeks);

  // Row 1: Monthly Theme
  const themeRow = ["Monthly Theme"];
  for (let i = 0; i < fridays.length; i++) themeRow.push("");
  sheet.appendRow(themeRow);

  // Row 2: Header — "Role / Date" + Friday dates
  const headerRow = ["Role / Date"];
  fridays.forEach(d => headerRow.push(d));
  sheet.appendRow(headerRow);

  // Format the date header cells (Row 2)
  const dateRange = sheet.getRange(2, 2, 1, fridays.length);
  dateRange.setNumberFormat("MMM d");  // e.g. "May 8"
  dateRange
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Role / Date header cell (Row 2)
  sheet.getRange(2, 1)
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  // Format the Monthly Theme row (Row 1)
  sheet.getRange(1, 1).setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  const themeRange = sheet.getRange(1, 2, 1, fridays.length);
  themeRange.setBackground("#e8f0fe").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");

  // Role rows
  roles.forEach(role => {
    const row = [role];
    // Fill with empty cells for now (dropdowns will be added later)
    for (let i = 0; i < fridays.length; i++) row.push("");
    sheet.appendRow(row);
  });

  // Format role column (Row 3 onwards)
  const roleRange = sheet.getRange(3, 1, roles.length, 1);
  roleRange
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  // Merge months in Row 1
  let currentMonth = -1;
  let startCol = 2;
  let count = 0;
  for (let i = 0; i < fridays.length; i++) {
    const month = fridays[i].getMonth();
    if (currentMonth === -1) {
      currentMonth = month;
      count = 1;
    } else if (month === currentMonth) {
      count++;
    } else {
      if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();
      currentMonth = month;
      startCol = startCol + count;
      count = 1;
    }
  }
  if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();

  // Freeze header rows and role column
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  // Column widths
  sheet.setColumnWidth(1, 160);
  if (fridays.length > 0) {
    sheet.setColumnWidths(2, fridays.length, 130);
  }

  // Row heights
  sheet.setRowHeights(1, 2, 36); // Theme and Dates
  if (roles.length > 0) {
    sheet.setRowHeights(3, roles.length, 32);
  }
}

// ─── GENERATE FRIDAY DATES ─────────────────────────────────
function generateFridays_(startDate, count) {
  const fridays = [];
  const d = new Date(startDate);

  // Make sure we start on a Friday (day 5)
  while (d.getDay() !== 5) {
    d.setDate(d.getDate() + 1);
  }

  for (let i = 0; i < count; i++) {
    fridays.push(new Date(d));
    d.setDate(d.getDate() + 7);
  }
  return fridays;
}

const SCHEDULE_HEADER_DATE_NOTE_PREFIX_ = "SCHEDULER_DATE:";

function dateKey_(date) {
  return Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

function parseDateKey_(key) {
  const parts = String(key || "").split("-");
  if (parts.length !== 3) return null;
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function addDays_(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function sameMonthDay_(date, month, day) {
  return date && date.getMonth() === month && date.getDate() === day;
}

function parseHeaderMonthDay_(displayText) {
  const firstLine = String(displayText || "").split("\n")[0].trim();
  const match = firstLine.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2})$/);
  if (!match) return null;

  const monthNames = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  const month = monthNames[match[1].toLowerCase()];
  const day = Number(match[2]);
  if (month === undefined || !day) return null;

  return { month: month, day: day };
}

function dateFromMonthDayNear_(monthDay, referenceDate) {
  const reference = referenceDate || CONFIG.startFriday;
  let best = null;
  let bestDistance = Infinity;

  for (let year = reference.getFullYear() - 1; year <= reference.getFullYear() + 1; year++) {
    const candidate = new Date(year, monthDay.month, monthDay.day);
    const distance = Math.abs(candidate.getTime() - reference.getTime());
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }

  return best;
}

function getHeaderDate_(value, note) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof note === "string" && note.indexOf(SCHEDULE_HEADER_DATE_NOTE_PREFIX_) === 0) {
    return parseDateKey_(note.replace(SCHEDULE_HEADER_DATE_NOTE_PREFIX_, "").trim());
  }
  return null;
}

function getScheduleHeaderDates_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return [];

  const range = sheet.getRange(2, 2, 1, lastCol - 1);
  const values = range.getValues()[0];
  const notes = range.getNotes()[0];
  const displays = range.getDisplayValues()[0];
  const dates = values.map((value, i) => getHeaderDate_(value, notes[i]));

  let lastKnownIndex = -1;
  let lastKnownDate = null;
  for (let i = 0; i < dates.length; i++) {
    if (dates[i]) {
      lastKnownIndex = i;
      lastKnownDate = dates[i];
      continue;
    }

    const monthDay = parseHeaderMonthDay_(displays[i]);
    if (!monthDay) continue;

    const expected = lastKnownDate
      ? addDays_(lastKnownDate, 7 * (i - lastKnownIndex))
      : addDays_(CONFIG.startFriday, 7 * i);
    dates[i] = sameMonthDay_(expected, monthDay.month, monthDay.day)
      ? expected
      : dateFromMonthDayNear_(monthDay, expected);

    lastKnownIndex = i;
    lastKnownDate = dates[i];
  }

  let nextKnownIndex = -1;
  let nextKnownDate = null;
  for (let i = dates.length - 1; i >= 0; i--) {
    if (dates[i]) {
      nextKnownIndex = i;
      nextKnownDate = dates[i];
      continue;
    }

    const monthDay = parseHeaderMonthDay_(displays[i]);
    if (!monthDay || !nextKnownDate) continue;

    const expected = addDays_(nextKnownDate, -7 * (nextKnownIndex - i));
    dates[i] = sameMonthDay_(expected, monthDay.month, monthDay.day)
      ? expected
      : dateFromMonthDayNear_(monthDay, expected);
  }

  return dates;
}

function remergeScheduleThemeRow_(sheet) {
  if (!sheet || sheet.getLastColumn() < 2) return;

  const allDates = getScheduleHeaderDates_(sheet);
  const themeRowRange = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1);
  themeRowRange.breakApart();

  let currentMonth = -1;
  let startCol = 2;
  let count = 0;
  for (let i = 0; i < allDates.length; i++) {
    const d = allDates[i];
    if (!(d instanceof Date) || isNaN(d)) {
      if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();
      currentMonth = -1;
      count = 0;
      continue;
    }

    const month = d.getMonth();
    if (currentMonth === -1) {
      currentMonth = month;
      startCol = 2 + i;
      count = 1;
    } else if (month === currentMonth) {
      count++;
    } else {
      if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();
      currentMonth = month;
      startCol = 2 + i;
      count = 1;
    }
  }

  if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();
}

function getEventsByDate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.eventsSheetName);
  const eventsByDate = {};
  if (!sheet || sheet.getLastRow() < 2) return eventsByDate;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(row => {
    const eventName = row[0];
    const eventDate = row[1];
    if (!eventName || !(eventDate instanceof Date) || isNaN(eventDate)) return;

    const key = dateKey_(eventDate);
    const label = String(eventName).trim();
    if (!label) return;
    if (!eventsByDate[key]) eventsByDate[key] = [];
    eventsByDate[key].push(label);
  });

  return eventsByDate;
}

function applyEventHeadersToSchedule_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet || sheet.getLastColumn() < 2) return;

  const dates = getScheduleHeaderDates_(sheet);
  if (dates.length === 0) return;

  const dateRange = sheet.getRange(2, 2, 1, dates.length);
  dateRange.setValues([dates])
    .setNumberFormat("MMM d")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(false);
  dateRange.setNotes([dates.map(d => d ? SCHEDULE_HEADER_DATE_NOTE_PREFIX_ + dateKey_(d) : "")]);

  const eventsByDate = getEventsByDate_();
  let hasEvent = false;

  dates.forEach((date, i) => {
    if (!date) return;

    const eventNames = eventsByDate[dateKey_(date)];
    if (!eventNames || eventNames.length === 0) return;

    hasEvent = true;
    const dateText = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM d");
    const eventText = eventNames.join(" / ").toUpperCase();
    const headerText = dateText + "\n" + eventText;
    const dateStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(10).build();
    const eventStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(8).build();
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(headerText)
      .setTextStyle(0, dateText.length, dateStyle)
      .setTextStyle(dateText.length + 1, headerText.length, eventStyle)
      .build();

    sheet.getRange(2, i + 2)
      .setRichTextValue(richText)
      .setBackground("#185abc")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setNote(SCHEDULE_HEADER_DATE_NOTE_PREFIX_ + dateKey_(date));
  });

  remergeScheduleThemeRow_(sheet);
  sheet.setRowHeight(2, hasEvent ? 44 : 36);
}

function normalizeEventsSheetHeaders_(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = sheet || ss.getSheetByName(CONFIG.eventsSheetName);
  if (!eventsSheet) return;

  eventsSheet.getRange(1, 1, 1, 4)
    .setValues([["Event Name", "Date", "", "ℹ️ Instructions"]])
    .setFontWeight("bold")
    .setBackground("#455a64")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
}

function syncEvents() {
  normalizeEventsSheetHeaders_();
  applyEventHeadersToSchedule_();
  SpreadsheetApp.getActiveSpreadsheet().toast("Events synced ✅", "Scheduler", 3);
}

// ─── REFRESH DROPDOWNS ─────────────────────────────────────
// Reads the Ministers sheet and rebuilds data‑validation dropdowns
// on the Schedule sheet so each role only shows eligible ministers.
function refreshAllDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet  = ss.getSheetByName(CONFIG.schedulerSheetName);
  const ministersSheet = ss.getSheetByName(CONFIG.ministersSheetName);

  if (!scheduleSheet || !ministersSheet) {
    SpreadsheetApp.getUi().alert("Cannot find Schedule or Ministers sheet. Run Full Setup first.");
    return;
  }

  // Read ministers data
  const ministersData   = ministersSheet.getDataRange().getValues();
  const ministersHeader = ministersData[0]; // ["Name", "Role1", "Role2", ...]
  const ministersRows   = ministersData.slice(1).filter(row =>
    row[0] !== "" && typeof row[0] === "string" && !row[0].startsWith("ℹ️")
  );

  // Build a map: roleName → [eligible minister names]
  const roleToNames = {};
  getRoles_().forEach(role => { roleToNames[role] = []; });

  ministersRows.forEach(row => {
    const name = row[0];
    for (let c = 1; c < ministersHeader.length; c++) {
      const roleName = ministersHeader[c];
      const isAssigned = row[c] === true || String(row[c]).toLowerCase() === "true";
      if (isAssigned) {
        if (!roleToNames[roleName]) roleToNames[roleName] = [];
        roleToNames[roleName].push(name);
      }
    }
  });

  // Read vacations data
  let vacations = [];
  const vacationsSheet = ss.getSheetByName(CONFIG.vacationsSheetName);
  if (vacationsSheet) {
    const vacData = vacationsSheet.getDataRange().getValues();
    for (let i = 1; i < vacData.length; i++) {
      const name = vacData[i][0];
      const start = vacData[i][1];
      const end = vacData[i][2];
      if (name && start instanceof Date && end instanceof Date) {
        const sDate = new Date(start); sDate.setHours(0,0,0,0);
        const eDate = new Date(end); eDate.setHours(23,59,59,999);
        vacations.push({name: name.toString().trim(), start: sDate.getTime(), end: eDate.getTime()});
      }
    }
  }

  // Read schedule structure
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const scheduleHeaderDates = getScheduleHeaderDates_(scheduleSheet);
  const numCols = scheduleData[0].length;
  if (numCols < 2) return;

  const scheduleDates = [];
  
  // Pre-fetch hidden states
  const isHidden = [];
  for (let c = 1; c < numCols; c++) {
    isHidden[c] = scheduleSheet.isColumnHiddenByUser(c + 1);
  }

  for (let c = 1; c < numCols; c++) {
    const headerDate = scheduleHeaderDates[c - 1];
    if (headerDate && !isHidden[c]) {
      const d = new Date(headerDate);
      d.setHours(12,0,0,0);
      scheduleDates[c] = d.getTime();
    } else {
      scheduleDates[c] = null;
    }
  }

  // For each role row (starting from top to scrub headers too), set data validation
  const validRoles = getRoles_();
  for (let r = 0; r < scheduleData.length; r++) {
    let role = scheduleData[r][0];
    if (!role || typeof role !== "string") continue;
    role = role.trim();
    if (!validRoles.includes(role)) {
      // Actively scrub orphaned notes and validations from non-role rows!
      const nonRoleRange = scheduleSheet.getRange(r + 1, 2, 1, numCols - 1);
      nonRoleRange.clearDataValidations();
      if (r >= 2) nonRoleRange.clearNote();
      continue;
    }

    const eligible = roleToNames[role] || [];
    const range = scheduleSheet.getRange(r + 1, 2, 1, numCols - 1);

    if (eligible.length > 0) {
      const rulesRow = [];
      let hasAnyWarning = false;

      for (let c = 1; c < numCols; c++) {
        const sDate = scheduleDates[c];
        let validMinisters = eligible;
        
        if (sDate) {
          validMinisters = eligible.filter(m => {
            return !vacations.some(v => v.name === m && sDate >= v.start && sDate <= v.end);
          });
        }

        if (validMinisters.length > 0) {
          rulesRow.push(
            SpreadsheetApp.newDataValidation()
              .requireValueInList(validMinisters, true)
              .setAllowInvalid(false)
              .build()
          );
        } else {
          rulesRow.push(null);
          hasAnyWarning = true;
        }
      }

      range.setDataValidations([rulesRow]);
      range.clearNote(); // Important: remove the warning if it existed!
    } else {
      // No eligible ministers — clear validation but add a note
      range.clearDataValidations();
      range.setNote("No minister assigned to this role. Update the Ministers sheet.");
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Dropdowns updated ✅", "Scheduler", 3);
}

// ─── ROLE GROUP FILTERING ───────────────────────────────────
// Reads group definitions from the Role Groups sheet and
// hides/shows rows on the Schedule sheet accordingly.

/**
 * Read group names from the Role Groups sheet header row.
 * Falls back to CONFIG if the sheet doesn't exist yet.
 */
function getGroupNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  if (sheet) {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return header.filter(h => h !== "");
  }
  return Object.keys(CONFIG.roleGroups);
}

/**
 * Read the roles belonging to a group (by group name) from the Role Groups sheet.
 * @param {string} groupName - The name of the group to fetch.
 * @param {Array<Array<any>>} [optData] - Optional pre-fetched data from the Role Groups sheet.
 */
function getGroupRoles_(groupName, optData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  if (!sheet) {
    return CONFIG.roleGroups[groupName] || getRoles_();
  }

  const data = optData || sheet.getDataRange().getValues();
  const header = data[0];
  const colIdx = header.indexOf(groupName);
  if (colIdx === -1) return getRoles_();

  const roles = [];
  for (let r = 1; r < data.length; r++) {
    const val = data[r][colIdx];
    if (val && typeof val === "string" && val.trim() !== "") {
      roles.push(val.trim());
    }
  }
  return roles.length > 0 ? roles : getRoles_();
}

/**
 * Core filter function — shows only the rows whose role is in the given list.
 * Optimized: batch-reads all role names in one call, then groups contiguous
 * rows for bulk show/hide to minimize API round-trips.
 */
function filterScheduleByRoles_(rolesToShow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  // Convert rolesToShow to a Set for O(1) lookups instead of O(n) Array.includes
  const roleSet = new Set(rolesToShow);

  // Batch-read all role names in column A (rows 3..lastRow) in one API call
  const roleValues = sheet.getRange(3, 1, lastRow - 2, 1).getValues();

  // Build an array of show/hide decisions
  // Then group contiguous rows with the same decision into batches
  let batchStart = 3;
  let batchShouldShow = roleSet.has(String(roleValues[0][0]).trim());

  for (let i = 1; i <= roleValues.length; i++) {
    const r = i + 3; // current sheet row (1-based)
    const shouldShow = (i < roleValues.length)
      ? roleSet.has(String(roleValues[i][0]).trim())
      : !batchShouldShow; // force flush of last batch

    if (shouldShow !== batchShouldShow) {
      // Flush the previous batch
      const count = r - batchStart;
      if (batchShouldShow) {
        sheet.showRows(batchStart, count);
      } else {
        sheet.hideRows(batchStart, count);
      }
      batchStart = r;
      batchShouldShow = shouldShow;
    }
  }

  // Store the active group name in a document property for reference
  ss.toast("Group filter applied ✅", "Scheduler", 3);
}

/**
 * Show all roles (remove any group filter).
 */
function showAllRoles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  sheet.showRows(3, lastRow - 2);
  ss.toast("All roles visible ✅", "Scheduler", 3);
}

/**
 * Filter by group at a specific index.
 * The menu calls filterGroup_0, filterGroup_1, etc.
 */
function filterByGroupIndex_(index) {
  const groupNames = getGroupNames_();
  if (index >= 0 && index < groupNames.length) {
    const groupName = groupNames[index];
    const roles = getGroupRoles_(groupName);
    filterScheduleByRoles_(roles);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Showing group: " + groupName, "Scheduler", 3
    );
  }
}

// ── Menu dispatcher functions (Google Apps Script requires named functions) ──
// These are auto‑called by the dynamic menu. Supports up to 20 groups.
function filterGroup_0()  { filterByGroupIndex_(0);  }
function filterGroup_1()  { filterByGroupIndex_(1);  }
function filterGroup_2()  { filterByGroupIndex_(2);  }
function filterGroup_3()  { filterByGroupIndex_(3);  }
function filterGroup_4()  { filterByGroupIndex_(4);  }
function filterGroup_5()  { filterByGroupIndex_(5);  }
function filterGroup_6()  { filterByGroupIndex_(6);  }
function filterGroup_7()  { filterByGroupIndex_(7);  }
function filterGroup_8()  { filterByGroupIndex_(8);  }
function filterGroup_9()  { filterByGroupIndex_(9);  }
function filterGroup_10() { filterByGroupIndex_(10); }
function filterGroup_11() { filterByGroupIndex_(11); }
function filterGroup_12() { filterByGroupIndex_(12); }
function filterGroup_13() { filterByGroupIndex_(13); }
function filterGroup_14() { filterByGroupIndex_(14); }
function filterGroup_15() { filterByGroupIndex_(15); }
function filterGroup_16() { filterByGroupIndex_(16); }
function filterGroup_17() { filterByGroupIndex_(17); }
function filterGroup_18() { filterByGroupIndex_(18); }
function filterGroup_19() { filterByGroupIndex_(19); }

// ─── ADD MONTH ─────────────────────────────────────────
function addMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  const scheduleHeaderDates = getScheduleHeaderDates_(sheet);
  const lastDate = scheduleHeaderDates[scheduleHeaderDates.length - 1];

  let nextFriday;
  if (lastDate instanceof Date) {
    nextFriday = new Date(lastDate);
    nextFriday.setDate(nextFriday.getDate() + 7);
  } else {
    nextFriday = CONFIG.startFriday;
  }

  const targetMonth = nextFriday.getMonth();
  const targetYear = nextFriday.getFullYear();
  
  const newFridays = [];
  let d = new Date(nextFriday);
  while (d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
    newFridays.push(new Date(d));
    d.setDate(d.getDate() + 7);
  }

  const numRows = sheet.getLastRow();

  // Add new Friday headers
  newFridays.forEach((friday, i) => {
    const col = lastCol + 1 + i;
    sheet.getRange(2, col).setValue(friday)
      .setNumberFormat("MMM d")
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    sheet.getRange(1, col).setBackground("#e8f0fe")
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  });
  if (newFridays.length > 0) {
    sheet.setColumnWidths(lastCol + 1, newFridays.length, 130);
  }

  applyEventHeadersToSchedule_();

  // Refresh dropdowns to cover new columns
  refreshAllDropdowns();

  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  ss.toast(`Added ${monthNames[targetMonth]} ${targetYear} (${newFridays.length} weeks) ✅`, "Scheduler", 3);
}

// ─── FORMATTING ─────────────────────────────────────────────
function applyFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();

  // Alternating row colors for readability and Data cells alignment
  if (numRows >= 3 && numCols > 1) {
    const dataRange = sheet.getRange(3, 2, numRows - 2, numCols - 1);
    
    // Batch set backgrounds instead of calling setBackground per row
    const bgs = [];
    for (let r = 3; r <= numRows; r++) {
      const rowBg = (r % 2 !== 0) ? "#f8f9fa" : "#ffffff";
      bgs.push(Array(numCols - 1).fill(rowBg));
    }
    dataRange.setBackgrounds(bgs);
    
    dataRange.setHorizontalAlignment("center");
    dataRange.setVerticalAlignment("middle");
    dataRange.setFontSize(10);
  }

  // Border around the entire data area
  const fullRange = sheet.getRange(1, 1, numRows, numCols);
  fullRange.setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  // ── Conflict detection via conditional formatting ──
  // These rules use COUNTIF to detect when the same minister name
  // appears multiple times in the same column (same Friday).
  // Rules are ordered by priority: highest severity first.

  function getColLetter(col) {
    let temp, letter = '';
    while (col > 0) {
      temp = (col - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
  }

  const visibleCols = [];
  const baseA1s = [];
  for (let c = 2; c <= numCols; c++) {
    if (!sheet.isColumnHiddenByUser(c)) {
      const letter = getColLetter(c);
      visibleCols.push(letter);
      baseA1s.push(`${letter}3:${letter}${numRows}`);
    }
  }

  if (visibleCols.length === 0) {
    sheet.clearConditionalFormatRules();
    ss.toast("Formatting applied (all weeks archived) ✅", "Scheduler", 3);
    return;
  }

  const baseRanges = sheet.getRangeList(baseA1s).getRanges();
  const fCol = visibleCols[0];
  const refCell = fCol + "3";

  const isConf3 = `COUNTIF(${fCol}$3:${fCol}$${numRows},${refCell})>=3`;
  const isConf2 = `COUNTIF(${fCol}$3:${fCol}$${numRows},${refCell})>=2`;

  // Base general rules
  const conf3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${refCell}<>"", ${isConf3})`)
    .setBackground("#ffcdd2").setFontColor("#000000").setBold(false).setRanges(baseRanges).build();

  const conf2Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(${refCell}<>"", ${isConf2})`)
    .setBackground("#ffe0b2").setFontColor("#000000").setBold(false).setRanges(baseRanges).build();

  const filledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellNotEmpty()
    .setBackground("#c8e6c9").setFontColor("#000000").setBold(false).setRanges(baseRanges).build();

  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground("#fff9c4").setRanges(baseRanges).build();

  // Dynamic Fatigue Warning rules per Role Group
  const rulesConf3Consec = [];
  const rulesConf2Consec = [];
  const rulesConsec = [];

  const sData = sheet.getDataRange().getValues();
  const roleToRow = {};
  for (let r = 2; r < sData.length; r++) {
    const role = sData[r][0];
    if (role && typeof role === "string" && role.trim() !== "") {
      roleToRow[role.trim()] = r + 1; // 1-based sheet row index
    }
  }

  const groupNames = getGroupNames_();
  
  // Pre-fetch role groups data to avoid redundant API calls inside the loop
  const roleGroupsSheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  const rgData = roleGroupsSheet ? roleGroupsSheet.getDataRange().getValues() : null;

  groupNames.forEach(gName => {
    // "All Roles" and "General Announcement" contradict specific group-based checking, skip them.
    if (gName === "All Roles" || gName === "General Announcement") return;
    const roles = getGroupRoles_(gName, rgData);
    const rows = [];
    roles.forEach(role => {
      if (roleToRow[role]) rows.push(roleToRow[role]);
    });

    if (rows.length === 0) return;

    const groupA1s = [];
    rows.forEach(r => {
      visibleCols.forEach(letter => {
        groupA1s.push(`${letter}${r}`);
      });
    });
    const ranges = sheet.getRangeList(groupA1s).getRanges();
    const topLeftCell = fCol + rows[0];

    const fLeft2 = "IF(COLUMN()>3, OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()-2)=${topLeftCell}`).join(",") + "), FALSE)";
    const fLeft1 = "IF(COLUMN()>2, OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()-1)=${topLeftCell}`).join(",") + "), FALSE)";
    const fRight1 = "OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()+1)=${topLeftCell}`).join(",") + ")";
    const fRight2 = "OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()+2)=${topLeftCell}`).join(",") + ")";
    const isConsecutive = `OR(AND(${fLeft2},${fLeft1}), AND(${fLeft1},${fRight1}), AND(${fRight1},${fRight2}))`;

    const gConf3 = `COUNTIF(${fCol}$3:${fCol}$${numRows},${topLeftCell})>=3`;
    const gConf2 = `COUNTIF(${fCol}$3:${fCol}$${numRows},${topLeftCell})>=2`;

    rulesConf3Consec.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${topLeftCell}<>"", ${gConf3}, ${isConsecutive})`)
        .setBackground("#ffcdd2").setFontColor("#d32f2f").setBold(false).setRanges(ranges).build()
    );
    rulesConf2Consec.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${topLeftCell}<>"", ${gConf2}, ${isConsecutive})`)
        .setBackground("#ffe0b2").setFontColor("#d32f2f").setBold(false).setRanges(ranges).build()
    );
    rulesConsec.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${topLeftCell}<>"", ${isConsecutive})`)
        .setBackground("#c8e6c9").setFontColor("#d32f2f").setBold(false).setRanges(ranges).build()
    );
  });

  // Order matters! First matching rule wins.
  sheet.setConditionalFormatRules([
    ...rulesConf3Consec, conf3Rule, 
    ...rulesConf2Consec, conf2Rule, 
    ...rulesConsec, filledRule, emptyRule
  ]);

  // Add Color Legend Note to the header cell
  const legendNote = "🎨 COLOR LEGEND:\n\n" +
    "🟩 Green BG: Assigned normally (no conflicts).\n" +
    "🟨 Yellow BG: Role is unfilled.\n" +
    "🟧 Orange BG: Double-booked (scheduled 2 times on this day).\n" +
    "🟥 Red BG: Triple-booked (scheduled 3+ times on this day).\n\n" +
    "🔴 Red Text: Fatigue warning! This person has been scheduled in this same Ministry Group for 3 or more consecutive weeks.";
  sheet.getRange(2, 1).setNote(legendNote);

  ss.toast("Formatting applied ✅", "Scheduler", 3);
}

// ─── AUTO‑UPDATE TRIGGER ────────────────────────────────────
// When setup sheets are edited, automatically refresh the dependent schedule areas.
function onEditTrigger(e) {
  try {
    const sheetName = e.range.getSheet().getName();
    if (sheetName === CONFIG.ministersSheetName || sheetName === CONFIG.vacationsSheetName) {
      refreshAllDropdowns();
    } else if (sheetName === CONFIG.eventsSheetName) {
      applyEventHeadersToSchedule_();
    }
  } catch (err) {
    // Silently ignore trigger errors
    Logger.log("onEdit error: " + err.message);
  }
}

// ─── ARCHIVE / UNARCHIVE DIALOG ──────────────────────────────
function showArchiveDialog() { showActionDialog_('archive'); }
function showUnarchiveDialog() { showActionDialog_('unarchive'); }

function showActionDialog_(action) {
  const result = getAllScheduleMonths_();
  if (result.error) {
    SpreadsheetApp.getUi().alert(result.error);
    return;
  }
  
  const validMonths = result.months.filter(m => action === 'archive' ? m.visibleCount > 0 : m.hiddenCount > 0);

  if (validMonths.length === 0) {
    const msg = action === 'archive' ? "There are no visible months left to archive." : "There are no archived months to unarchive.";
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  const optionsHtml = validMonths.map((opt) =>
    '<option value="' + opt.month + '_' + opt.year + '">' + opt.label + '</option>'
  ).join('');

  const title = action === 'archive' ? "Archive Month" : "Unarchive Month";
  const btnText = action === 'archive' ? "Archive" : "Unarchive";

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        * { box-sizing: border-box; }
        body {
          margin: 0; padding: 24px;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
          background-color: #ffffff; color: #3c4043;
          overflow: hidden;
        }
        .label { font-size: 14px; font-weight: 500; margin-bottom: 12px; display: block; }
        select {
          width: 100%; height: 40px; padding: 0 12px; border-radius: 4px; border: 1px solid #dadce0;
          font-size: 14px; background-color: #fff; outline: none; transition: border 0.2s;
          cursor: pointer; -webkit-appearance: none; -moz-appearance: none; margin-bottom: 24px;
        }
        select:focus { border: 2px solid #1a73e8; }
        .buttons { display: flex; justify-content: flex-end; gap: 12px; margin-top: 12px; }
        button {
          height: 36px; padding: 0 20px; border-radius: 4px; font-size: 14px; font-weight: 500;
          cursor: pointer; border: 1px solid transparent; transition: background 0.2s;
        }
        .btn-action { background-color: #1a73e8; color: #ffffff; }
        .btn-action:hover { background-color: #1b66c9; }
        .btn-action:disabled { opacity: 0.6; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <label class="label" for="monthSelect">Select Month to ${btnText}:</label>
      <select id="monthSelect">
        ${optionsHtml}
      </select>
      <div class="buttons">
        <button id="btnAction" class="btn-action" onclick="doAction()">${btnText}</button>
      </div>

      <script>
        function doAction() {
          document.getElementById('btnAction').disabled = true;
          document.getElementById('btnAction').textContent = 'Working...';
          const val = document.getElementById('monthSelect').value;
          const parts = val.split('_');
          const month = parseInt(parts[0]);
          const year = parseInt(parts[1]);
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .withFailureHandler((err) => { alert(err.message); google.script.host.close(); })
            .runArchiveActionByVal(month, year, '${action}');
        }
      </script>
    </body>
    </html>
  `;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(340).setHeight(180),
    title
  );
}

function runArchiveActionByVal(month, year, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  
  const lastCol = sheet.getLastColumn();
  const dateHeaderRow = getScheduleHeaderDates_(sheet);

  let startCol = -1;
  let count = 0;
  let modifiedCount = 0;

  for (let i = 0; i < dateHeaderRow.length; i++) {
    const d = dateHeaderRow[i];
    let isTarget = false;
    
    if (d && d.getMonth() === month && d.getFullYear() === year) {
      if (action === 'archive' && !sheet.isColumnHiddenByUser(i + 2)) isTarget = true;
      if (action === 'unarchive' && sheet.isColumnHiddenByUser(i + 2)) isTarget = true;
    }

    if (isTarget) {
      if (startCol === -1) startCol = i + 2;
      count++;
    } else {
      if (startCol !== -1) {
        if (action === 'archive') sheet.hideColumns(startCol, count);
        else sheet.showColumns(startCol, count);
        modifiedCount += count;
        startCol = -1;
        count = 0;
      }
    }
  }

  if (startCol !== -1) {
    if (action === 'archive') sheet.hideColumns(startCol, count);
    else sheet.showColumns(startCol, count);
    modifiedCount += count;
  }

  const actStr = action === 'archive' ? 'Archived' : 'Unarchived';
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const label = monthNames[month] + " " + year;

  if (modifiedCount > 0) {
    ss.toast(`${actStr} ${modifiedCount} week(s) for ${label} ✅`, "Scheduler", 4);
  } else {
    ss.toast(`Nothing to change for ${label}.`, "Scheduler", 4);
  }
}

/**
 * Get all months from the schedule, tracking their hidden vs visible status.
 */
function getAllScheduleMonths_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return { months: [], error: "Schedule sheet not found." };

  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return { months: [], error: "No dates found." };

  const dateHeaderRow = getScheduleHeaderDates_(sheet);
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];

  // Pre-fetch hidden states
  const isHidden = [];
  for (let i = 0; i < dateHeaderRow.length; i++) {
    isHidden[i] = sheet.isColumnHiddenByUser(i + 2);
  }

  const monthMap = new Map();

  for (let i = 0; i < dateHeaderRow.length; i++) {
    const d = dateHeaderRow[i];
    if (d) {
      const label = monthNames[d.getMonth()] + " " + d.getFullYear();
      if (!monthMap.has(label)) {
        monthMap.set(label, { label: label, month: d.getMonth(), year: d.getFullYear(), visibleCount: 0, hiddenCount: 0 });
      }
      if (isHidden[i]) {
        monthMap.get(label).hiddenCount++;
      } else {
        monthMap.get(label).visibleCount++;
      }
    }
  }
  
  return { months: Array.from(monthMap.values()), error: null };
}

// ─── REUSABLE MONTH RANGE HELPERS ───────────────────────────
// Shared utilities for any feature that needs a 2-month range picker.
// Reads ONLY the date header row (row 2) — not the entire sheet — for speed.

/**
 * Read only the date header row and derive paired 2-month ranges.
 * Returns { pairs: [{label, startMonth, startYear, endMonth, endYear}], error: string|null }
 * Fast: reads a single row instead of getDataRange().getValues().
 */
function getScheduleMonthPairs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return { pairs: [], error: "Schedule sheet not found. Run Full Setup first." };

  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return { pairs: [], error: "No dates found on the Schedule sheet." };

  // Read ONLY row 2 (date headers) — much faster than getDataRange()
  const dateHeaderRow = getScheduleHeaderDates_(sheet);
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];

  // Extract unique Month-Year labels, ignoring hidden columns
  const months = [];
  const seen = new Set();

  // Pre-fetch hidden states
  const isHidden = [];
  for (let i = 0; i < dateHeaderRow.length; i++) {
    isHidden[i] = sheet.isColumnHiddenByUser(i + 2);
  }

  for (let i = 0; i < dateHeaderRow.length; i++) {
    if (isHidden[i]) continue; // Skip archived columns

    const d = dateHeaderRow[i];
    if (d) {
      const label = monthNames[d.getMonth()] + " " + d.getFullYear();
      if (!seen.has(label)) {
        months.push({ label: label, month: d.getMonth(), year: d.getFullYear() });
        seen.add(label);
      }
    }
  }

  if (months.length === 0) return { pairs: [], error: "No dates found on the Schedule sheet." };

  // Build paired options: each starting month + next month
  const pairs = [];
  for (let i = 0; i < months.length; i++) {
    const start = months[i];
    const next = (i + 1 < months.length) ? months[i + 1] : null;
    if (next) {
      pairs.push({
        label: start.label + " – " + next.label,
        startMonth: start.month, startYear: start.year,
        endMonth: next.month, endYear: next.year
      });
    } else {
      pairs.push({
        label: start.label + " (single month)",
        startMonth: start.month, startYear: start.year,
        endMonth: start.month, endYear: start.year
      });
    }
  }

  return { pairs: pairs, error: null };
}

/**
 * Show a reusable 2-month range picker dialog.
 * @param {string} title - Dialog window title
 * @param {string} labelText - Prompt label shown above the dropdown
 * @param {string} hint - Smaller hint text (or "" to hide)
 * @param {string} buttonLabel - Text on the action button
 * @param {string} serverCallback - Name of the server-side function to call with (pairedIndex)
 * @param {boolean} includeAllOption - If true, adds a "Check Entire Schedule" option at index -1
 */
function showMonthRangeDialog_(title, labelText, hint, buttonLabel, serverCallback, includeAllOption) {
  const result = getScheduleMonthPairs_();
  if (result.error) {
    SpreadsheetApp.getUi().alert(result.error);
    return;
  }

  const pairs = result.pairs;
  const allOptionHtml = includeAllOption
    ? '<option value="-1">Check Entire Schedule</option>'
    : '';
  const optionsHtml = pairs.map((opt, i) =>
    '<option value="' + i + '">' + opt.label + '</option>'
  ).join('');
  const hintHtml = hint ? '<p class="hint">' + hint + '</p>' : '';

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        * { box-sizing: border-box; }
        body {
          margin: 0; padding: 24px;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
          background-color: #ffffff; color: #3c4043;
          overflow: hidden;
        }
        .label { font-size: 14px; font-weight: 500; margin-bottom: 12px; display: block; }
        .hint  { font-size: 12px; color: #80868b; margin-bottom: 16px; margin-top: 0; }
        select {
          width: 100%; height: 40px; padding: 0 12px; border-radius: 4px; border: 1px solid #dadce0;
          font-size: 14px; background-color: #fff; outline: none; transition: border 0.2s;
          cursor: pointer; -webkit-appearance: none; -moz-appearance: none;
        }
        select:focus { border: 2px solid #1a73e8; }
        .buttons { display: flex; justify-content: flex-end; gap: 12px; margin-top: 28px; }
        button {
          height: 36px; padding: 0 24px; border-radius: 4px; font-size: 14px; font-weight: 500;
          cursor: pointer; border: 1px solid transparent; transition: background 0.2s;
        }
        .btn-cancel { background: none; color: #1a73e8; border: 1px solid #dadce0; }
        .btn-cancel:hover { background-color: #f8f9fa; }
        .btn-action { background-color: #1a73e8; color: #ffffff; }
        .btn-action:hover { background-color: #1b66c9; box-shadow: 0 1px 2px 0 rgba(60,64,67,0.302), 0 1px 3px 1px rgba(60,64,67,0.149); }
      </style>
    </head>
    <body>
      <label class="label" for="rangeSelect">${labelText}</label>
      ${hintHtml}
      <select id="rangeSelect">
        ${allOptionHtml}
        ${optionsHtml}
      </select>
      <div class="buttons">
        <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-action" onclick="doAction()">${buttonLabel}</button>
      </div>
      <script>
        function doAction() {
          var btn = document.querySelector('.btn-action');
          btn.disabled = true; btn.textContent = 'Working…';
          var idx = parseInt(document.getElementById('rangeSelect').value);
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) { btn.disabled = false; btn.textContent = '${buttonLabel}'; alert(e.message); })
            .${serverCallback}(idx);
        }
      </script>
    </body>
    </html>
  `;

  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(hint ? 230 : 200)
    .setTitle(title);

  SpreadsheetApp.getUi().showModalDialog(userInterface, title);
}

// ─── CONFLICT CHECKER ───────────────────────────────────────
// Scans the Schedule sheet and reports all ministers who are
// booked into multiple roles on the same Friday.
function checkConflicts() {
  showMonthRangeDialog_(
    "⚠️ Conflict Checker",
    "Select period to analyze:",
    "",
    "Check Now",
    "runConflictCheck",
    true  // include "Check Entire Schedule" option
  );
}

/**
 * Server-side function called by the dialog.
 * pairedIndex: -1 means "all", otherwise the index into the paired month ranges.
 */
function runConflictCheck(pairedIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  const data = sheet.getDataRange().getValues();
  const scheduleHeaderDates = getScheduleHeaderDates_(sheet);
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];

  // Determine which months to include
  let chosen = null;
  let filterLabel = "the entire schedule";
  if (pairedIndex >= 0) {
    const result = getScheduleMonthPairs_();
    if (result.error || pairedIndex >= result.pairs.length) return;
    chosen = result.pairs[pairedIndex];
    filterLabel = chosen.label;
  }

  const allConflicts = [];
  const numRows = data.length;
  const numCols = data[0].length;

  // Pre-calculate which columns are hidden
  const isHidden = [];
  for (let c = 1; c < numCols; c++) {
    isHidden[c] = sheet.isColumnHiddenByUser(c + 1);
  }

  for (let c = 1; c < numCols; c++) {
    if (isHidden[c]) continue; // Ignore archived/hidden weeks

    const dateHeader = scheduleHeaderDates[c - 1];

    // Filtering logic — skip columns outside the chosen 2-month range
    if (chosen) {
      if (dateHeader) {
        const m = dateHeader.getMonth();
        const y = dateHeader.getFullYear();
        const inRange =
          (y === chosen.startYear && m === chosen.startMonth) ||
          (y === chosen.endYear && m === chosen.endMonth);
        if (!inRange) continue;
      } else {
        continue;
      }
    }

    const nameCount = {};
    for (let r = 2; r < numRows; r++) {
      const name = data[r][c];
      const role = data[r][0];
      if (name && typeof name === "string" && name.trim() !== "") {
        if (!nameCount[name]) nameCount[name] = [];
        nameCount[name].push(role);
      }
    }

    for (const [name, roles] of Object.entries(nameCount)) {
      if (roles.length >= 2) {
        let dateStr = (dateHeader)
          ? Utilities.formatDate(dateHeader, ss.getSpreadsheetTimeZone(), "MMM d, yyyy")
          : String(dateHeader);

        allConflicts.push({
          date: dateStr,
          name: name,
          roles: roles,
          count: roles.length
        });
      }
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (allConflicts.length === 0) {
    ui.alert("✅ No Conflicts", "No conflicts found for " + filterLabel + ".", ui.ButtonSet.OK);
    return;
  }

  allConflicts.sort((a, b) => b.count - a.count || a.date.localeCompare(b.date));
  let report = "Found " + allConflicts.length + " conflict(s) in " + filterLabel + ":\n\n";

  allConflicts.forEach(c => {
    const severity = c.count >= 3 ? "🔴" : "🟠";
    report += severity + " " + c.name + " — " + c.date + "\n";
    report += "   Assigned to " + c.count + " roles: " + c.roles.join(", ") + "\n\n";
  });

  const severe = allConflicts.filter(c => c.count >= 3).length;
  const moderate = allConflicts.filter(c => c.count === 2).length;
  report += "──────────────\n";
  report += "🟠 Double-booked: " + moderate + "  |  🔴 Triple+ booked: " + severe;

  ui.alert("⚠️ Conflict Report", report, ui.ButtonSet.OK);
}

function installTrigger_() {
  // Remove any existing scheduler triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onEditTrigger") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Install new onEdit trigger
  ScriptApp.newTrigger("onEditTrigger")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

// ─── CLEAR FORMATTING ───────────────────────────────────────
function clearFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();

  if (numRows > 2 && numCols > 1) {
    const dataRange = sheet.getRange(3, 2, numRows - 2, numCols - 1);

    // Clear conditional format rules (removes conflict/fatigue indicators)
    sheet.clearConditionalFormatRules();

    // Set background to alternating zebra stripes using batch operation
    const bgs = [];
    for (let r = 3; r <= numRows; r++) {
      const bg = (r % 2 !== 0) ? "#f8f9fa" : "#ffffff";
      bgs.push(Array(numCols - 1).fill(bg));
    }
    dataRange.setBackgrounds(bgs);

    // Ensure font color is black and not bold
    dataRange.setFontColor("#000000");
    dataRange.setFontWeight("normal");
  }

  ss.toast("Formatting cleared to plain zebra stripes ✅", "Scheduler", 3);
}

// ─── PRINTING EXPORT ────────────────────────────────────────
/**
 * Shows the shared 2-month range picker, then generates the print sheet.
 * Hidden rows (from group filtering) are automatically excluded.
 */
function exportForPrinting() {
  showMonthRangeDialog_(
    "🖨 Export for Printing",
    "Select 2-month range to print:",
    "Hidden rows (filtered groups) will be excluded automatically.",
    "🖨 Generate Print Sheet",
    "runPrintRange",
    false  // no "all" option for printing
  );
}

/**
 * Server-side function called by the print dialog.
 * Uses the shared getScheduleMonthPairs_() helper instead of re-deriving months.
 */
function runPrintRange(pairedIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const result = getScheduleMonthPairs_();
  if (result.error || pairedIndex < 0 || pairedIndex >= result.pairs.length) return;
  const chosen = result.pairs[pairedIndex];

  // Read only the date header row to find matching columns
  const lastCol = sheet.getLastColumn();
  const dateHeaderRow = getScheduleHeaderDates_(sheet);

  // Determine which columns (dates) fall within the chosen range
  const colIndices = []; // 1-based column indices on the sheet
  
  // Pre-fetch hidden states
  const isHidden = [];
  for (let i = 0; i < dateHeaderRow.length; i++) {
    isHidden[i] = sheet.isColumnHiddenByUser(i + 2);
  }

  for (let i = 0; i < dateHeaderRow.length; i++) {
    if (isHidden[i]) continue; // Ignore archived weeks

    const d = dateHeaderRow[i];
    if (d) {
      const m = d.getMonth();
      const y = d.getFullYear();
      const inRange =
        (y === chosen.startYear && m === chosen.startMonth) ||
        (y === chosen.endYear && m === chosen.endMonth);
      if (inRange) {
        colIndices.push(i + 2); // +2: 0-based array offset + column 1 is roles
      }
    }
  }

  if (colIndices.length === 0) {
    SpreadsheetApp.getUi().alert("No dates found for the selected range.");
    return;
  }

  // Always include column 1 (Role names) as the first column
  const allCols = [1, ...colIndices];

  // Determine which rows to include — skip hidden rows
  // Include row 1 (Monthly Theme) and row 2 (Date headers), then visible role rows
  const visibleRows = [1, 2]; // theme + header
  const lastRow = sheet.getLastRow();
  for (let r = 3; r <= lastRow; r++) {
    if (!sheet.isRowHiddenByUser(r)) {
      visibleRows.push(r);
    }
  }

  if (visibleRows.length <= 2) {
    SpreadsheetApp.getUi().alert("No visible role rows to print. Show some roles first.");
    return;
  }

  generatePrintableFromRange_(sheet, visibleRows, allCols);
}

/**
 * Public function to create the Settings sheet without a full reset
 */
function initializeSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  if (sheet) {
    const response = SpreadsheetApp.getUi().alert("Settings sheet already exists. Overwrite it?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (response !== SpreadsheetApp.getUi().Button.YES) return;
    sheet.clear();
  } else {
    sheet = ss.insertSheet("Settings");
  }
  buildSettingsSheet_(sheet);
  ss.setActiveSheet(sheet);
  ss.toast("Settings sheet created! ✅", "Scheduler", 3);
}

// ─── SETTINGS HELPER ────────────────────────────────────────
function buildSettingsSheet_(sheet) {
  const settings = [
    ["Setting Name", "Image", "Instructions"],
    ["Header Image", "", "Go to B2 and select Insert ▸ Image ▸ Image in cell"],
    ["Footer Image", "", "Go to B3 and select Insert ▸ Image ▸ Image in cell"]
  ];
  sheet.getRange(1, 1, settings.length, settings[0].length).setValues(settings);
  
  // Formatting
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#455a64").setFontColor("#ffffff");
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 600);
  sheet.setColumnWidth(3, 400);
  sheet.setRowHeight(2, 100); 
  sheet.setRowHeight(3, 100);
  sheet.setFrozenRows(1);

  // Center the image cells
  sheet.getRange("B2:B3").setHorizontalAlignment("center").setVerticalAlignment("middle");
}

function getPrintImages_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return { header: null, footer: null };

  return {
    header: sheet.getRange(2, 2).getValue(),
    footer: sheet.getRange(3, 2).getValue()
  };
}

/**
 * Generates a print sheet from specific rows and columns of the Schedule sheet.
 * Uses batch read/write operations to avoid slow cell-by-cell API calls.
 * @param {Sheet} srcSheet - The source Schedule sheet
 * @param {number[]} rows - Array of 1-based row indices to include
 * @param {number[]} cols - Array of 1-based column indices to include
 */
function generatePrintableFromRange_(srcSheet, rows, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const numRows = rows.length;
  const numCols = cols.length;
  if (numRows === 0 || numCols === 0) return;

  // 1. Batch-read ALL formatting data from source sheet in one go
  const fullRange = srcSheet.getDataRange();
  const allValues      = fullRange.getValues();
  const allBgs         = fullRange.getBackgrounds();
  const allFontColors  = fullRange.getFontColors();
  const allFontWeights = fullRange.getFontWeights();
  const allFontSizes   = fullRange.getFontSizes();
  const allHAligns     = fullRange.getHorizontalAlignments();
  const allVAligns     = fullRange.getVerticalAlignments();
  const allNumFormats  = fullRange.getNumberFormats();
  const allRichText    = fullRange.getRichTextValues();

  // 2. Assemble only the needed rows/cols into output arrays (in-memory)
  const values = [], bgs = [], fontColors = [], fontWeights = [];
  const fontSizes = [], hAligns = [], vAligns = [], numFormats = [];

  for (let r = 0; r < numRows; r++) {
    const srcRow = rows[r] - 1; // convert to 0-based index
    const vRow = [], bgRow = [], fcRow = [], fwRow = [];
    const fsRow = [], haRow = [], vaRow = [], nfRow = [];
    for (let c = 0; c < numCols; c++) {
      const srcCol = cols[c] - 1; // convert to 0-based index
      vRow.push(allValues[srcRow][srcCol]);
      bgRow.push(allBgs[srcRow][srcCol]);
      fcRow.push(allFontColors[srcRow][srcCol]);
      fwRow.push(allFontWeights[srcRow][srcCol]);
      fsRow.push(allFontSizes[srcRow][srcCol]);
      haRow.push(allHAligns[srcRow][srcCol]);
      vaRow.push(allVAligns[srcRow][srcCol]);
      nfRow.push(allNumFormats[srcRow][srcCol]);
    }
    values.push(vRow);       bgs.push(bgRow);
    fontColors.push(fcRow);  fontWeights.push(fwRow);
    fontSizes.push(fsRow);   hAligns.push(haRow);
    vAligns.push(vaRow);     numFormats.push(nfRow);
  }

  // 3. Create or clean the print sheet
  const printSheetName = "Print - Selection";
  let printSheet = ss.getSheetByName(printSheetName);
  if (printSheet) {
    ss.deleteSheet(printSheet);
  }
  printSheet = ss.insertSheet(printSheetName);

  // 4. Batch-write everything to the print sheet
  const headerOffset = 7;
  const targetRange = printSheet.getRange(headerOffset, 1, numRows, numCols);
  targetRange.setValues(values);
  targetRange.setBackgrounds(bgs);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  targetRange.setFontSizes(fontSizes);
  targetRange.setHorizontalAlignments(hAligns);
  targetRange.setVerticalAlignments(vAligns);
  targetRange.setNumberFormats(numFormats);
  targetRange.clearDataValidations();
  targetRange.setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  const sourceHeaderDates = getScheduleHeaderDates_(srcSheet);
  const headerRowOffset = rows.indexOf(2);
  if (headerRowOffset !== -1) {
    const printHeaderRow = headerOffset + headerRowOffset;
    cols.forEach((col, i) => {
      const richText = allRichText[1][col - 1];
      if (richText && richText.getText && richText.getText().indexOf("\n") !== -1) {
        printSheet.getRange(printHeaderRow, i + 1).setRichTextValue(richText);
      }
    });
  }

  // 4.5. Re-apply merges in Row 1 of the print sheet (Monthly Theme)
  if (numCols > 2 && values.length > 1) {
    const printDates = cols.slice(1).map(col => sourceHeaderDates[col - 2]);
    let currentMonth = -1;
    let startCol = 2;
    let count = 0;
    
    for (let i = 0; i < printDates.length; i++) {
      const d = printDates[i];
      if (d instanceof Date) {
        const month = d.getMonth();
        if (currentMonth === -1) {
          currentMonth = month;
          count = 1;
        } else if (month === currentMonth) {
          count++;
        } else {
          if (count > 1) {
            printSheet.getRange(headerOffset, startCol, 1, count).mergeAcross();
          }
          currentMonth = month;
          startCol = 2 + i;
          count = 1;
        }
      } else {
        if (count > 1) {
          printSheet.getRange(headerOffset, startCol, 1, count).mergeAcross();
        }
        currentMonth = -1;
        count = 0;
      }
    }
    if (count > 1) {
      printSheet.getRange(headerOffset, startCol, 1, count).mergeAcross();
    }
  }

  // 5. Match row heights and column widths
  // Group identical row heights to batch writes
  if (numRows > 0) {
    let currentHeight = srcSheet.getRowHeight(rows[0]);
    let startRow = headerOffset;
    let count = 1;
    for (let r = 1; r < numRows; r++) {
      const h = srcSheet.getRowHeight(rows[r]);
      if (h === currentHeight) {
        count++;
      } else {
        printSheet.setRowHeights(startRow, count, currentHeight);
        currentHeight = h;
        startRow = headerOffset + r;
        count = 1;
      }
    }
    printSheet.setRowHeights(startRow, count, currentHeight);
  }

  // Group identical column widths to batch writes
  if (numCols > 0) {
    let currentWidth = srcSheet.getColumnWidth(cols[0]);
    let startCol = 1;
    let count = 1;
    for (let c = 1; c < numCols; c++) {
      const w = srcSheet.getColumnWidth(cols[c]);
      if (w === currentWidth) {
        count++;
      } else {
        printSheet.setColumnWidths(startCol, count, currentWidth);
        currentWidth = w;
        startCol = 1 + c;
        count = 1;
      }
    }
    printSheet.setColumnWidths(startCol, count, currentWidth);
  }

  SpreadsheetApp.flush();

  // 6. Insert Header Image
  try {
    const settingsSheet = ss.getSheetByName("Settings");
    if (settingsSheet) {
      const headerRange = printSheet.getRange(1, 1, headerOffset - 1, numCols);
      settingsSheet.getRange(2, 2).copyTo(printSheet.getRange(1, 1));
      headerRange.merge().setHorizontalAlignment("center").setVerticalAlignment("middle");
    }
  } catch (e) { console.log("Header image failed: " + e); }

  // 7. Insert Footer Image
  try {
    const settingsSheet = ss.getSheetByName("Settings");
    if (settingsSheet) {
      const footerRow = headerOffset + numRows;
      const footerHeight = headerOffset - 1;
      const footerRange = printSheet.getRange(footerRow, 1, footerHeight, numCols);
      settingsSheet.getRange(3, 2).copyTo(printSheet.getRange(footerRow, 1));
      footerRange.merge().setHorizontalAlignment("center").setVerticalAlignment("middle");
    }
  } catch (e) { console.log("Footer image failed: " + e); }

  // 8. Polish UI
  try {
    printSheet.setHideGridlines(true);
  } catch (e) {}

  ss.setActiveSheet(printSheet);
  ss.toast("Print sheet generated! ✅", "Success", 5);
}

// ─── VACATIONS HELPER ────────────────────────────────────────
function initializeVacationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.vacationsSheetName || "Vacations");
  if (sheet) {
    const response = SpreadsheetApp.getUi().alert("Vacations sheet already exists. Overwrite it?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (response !== SpreadsheetApp.getUi().Button.YES) return;
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.vacationsSheetName || "Vacations");
  }
  buildVacationsSheet_(sheet);
  ss.setActiveSheet(sheet);
  ss.toast("Vacations sheet created! ✅", "Scheduler", 3);
}

function buildVacationsSheet_(sheet) {
  sheet.appendRow(["Minister Name", "Start Date", "End Date", "", "ℹ️ Instructions"]);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange
    .setFontWeight("bold")
    .setBackground("#fbbc04")
    .setFontColor("#000000")
    .setHorizontalAlignment("center");
    
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 50); // spacer
  sheet.setColumnWidth(5, 500); // instructions
  sheet.setFrozenRows(1);
  
  // Data validation for Minister Name (dropdown from Ministers sheet)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ministersSheet = ss.getSheetByName(CONFIG.ministersSheetName);
  if (ministersSheet) {
    // We only need column A starting from row 2
    // But since the number of rows might change, we define a dynamic range
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(ministersSheet.getRange("A2:A"), true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange("A2:A").setDataValidation(rule);
  }
  
  // Date validation for Start Date and End Date
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange("B2:C").setDataValidation(dateRule);

  // Instructions
  sheet.getRange("E2").setValue("Add minister vacations here. They will be removed from dropdowns for dates within their vacation period.").setFontColor("#666666").setFontStyle("italic");
  sheet.getRange("E3").setValue("Start Date and End Date must be valid dates (e.g. 5/14/2026).").setFontColor("#666666").setFontStyle("italic");
  sheet.getRange("E4").setValue("Make sure to run Scheduler ▸ Update Minister Dropdowns after adding vacations!").setFontColor("#666666").setFontStyle("italic");
}

function initializeEventsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.eventsSheetName || "Events");
  let existingEvents = [];
  if (sheet) {
    if (sheet.getLastRow() > 1) {
      existingEvents = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
        .filter(row => row[0] !== "" || row[1] !== "");
    }
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.eventsSheetName || "Events");
  }

  buildEventsSheet_(sheet);
  if (existingEvents.length > 0) {
    sheet.getRange(2, 1, existingEvents.length, 2).setValues(existingEvents);
  }
  applyEventHeadersToSchedule_();
  ss.setActiveSheet(sheet);
  ss.toast("Events sheet created/repaired! ✅", "Scheduler", 3);
}

function buildEventsSheet_(sheet) {
  normalizeEventsSheetHeaders_(sheet);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(4, 520);
  sheet.setFrozenRows(1);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  sheet.getRange("B2:B").setDataValidation(dateRule).setNumberFormat("mmm d, yyyy");

  sheet.getRange("D2")
    .setValue("Add events here. If an event date matches a Schedule date, that date header will show the date and event name.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange("D3")
    .setValue("Column A is the event name. Column B is the event date.")
    .setFontColor("#666666")
    .setFontStyle("italic");
  sheet.getRange("D4")
    .setValue("After editing events, the Schedule headers refresh automatically if the onEdit trigger is installed.")
    .setFontColor("#666666")
    .setFontStyle("italic");
}
