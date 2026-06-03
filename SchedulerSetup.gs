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

  ui.createMenu("⏱ Scheduler")
    .addSubMenu(groupMenu)
    .addSeparator()
    .addItem("⚠️ Check Conflicts",            "checkConflicts")
    .addItem("🎨 Re‑apply Formatting",        "applyFormatting")
    .addItem("🤍 Clear Formatting",           "clearFormatting")
    .addItem("🖨 Export for Printing",        "exportForPrinting")
    .addSeparator()
    .addItem("➕ Add 4 More Weeks",            "addMoreWeeks")
    .addItem("👥 Update Minister Dropdowns",    "refreshAllDropdowns")
    .addItem("🔄 Sync New/Removed Roles",      "syncRoles")
    .addSeparator()
    .addItem("🛠 Create/Repair Settings Sheet", "initializeSettingsSheet")
    .addItem("🏖 Create/Repair Vacations Sheet", "initializeVacationsSheet")
    .addItem("⚙️ Full Setup (WARNING: DELETES ALL DATA)", "setupScheduler")
    .addToUi();
}

// ─── MAIN SETUP ─────────────────────────────────────────────
function setupScheduler() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Setup Scheduler",
    "This will create / overwrite the Schedule, Ministers, Roles, Vacations, and Role Groups sheets.\nContinue?",
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

  // 2. Build Settings sheet
  buildSettingsSheet_(settingsSheet);

  // 2.5 Build Vacations sheet
  buildVacationsSheet_(vacationsSheet);

  // 3. Build Roles sheet first (source of truth)
  buildRolesSheet_(rolesSheet);
  SpreadsheetApp.flush();  // Ensure roles data is written before other sheets read it

  // 3. Build Role Groups sheet
  buildRoleGroupsSheet_(roleGroupsSheet);

  // 4. Build Ministers sheet (referenced by Schedule)
  buildMinistersSheet_(ministersSheet);

  // 5. Build Schedule sheet
  buildScheduleSheet_(scheduleSheet);

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
  for (let c = 1; c <= groupNames.length; c++) {
    sheet.setColumnWidth(c, 160);
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
  for (let c = 2; c <= roles.length + 1; c++) {
    sheet.setColumnWidth(c, 120);
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
  for (let c = 2; c <= fridays.length + 1; c++) {
    sheet.setColumnWidth(c, 130);
  }

  // Row heights
  sheet.setRowHeight(1, 36); // Theme
  sheet.setRowHeight(2, 36); // Dates
  for (let r = 3; r <= roles.length + 2; r++) {
    sheet.setRowHeight(r, 32);
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
  const numCols = scheduleData[0].length;
  if (numCols < 2) return;

  const scheduleDates = [];
  for (let c = 1; c < numCols; c++) {
    const cellVal = scheduleData[1][c];
    if (cellVal instanceof Date) {
      const d = new Date(cellVal);
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
      scheduleSheet.getRange(r + 1, 2, 1, numCols - 1).clearDataValidations().clearNote();
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
 */
function getGroupRoles_(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.roleGroupsSheetName);
  if (!sheet) {
    return CONFIG.roleGroups[groupName] || getRoles_();
  }

  const data = sheet.getDataRange().getValues();
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
 */
function filterScheduleByRoles_(rolesToShow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  // Iterate over role rows (row 3 onward)
  for (let r = 3; r <= lastRow; r++) {
    const role = sheet.getRange(r, 1).getValue();
    if (rolesToShow.includes(role)) {
      sheet.showRows(r);
    } else {
      sheet.hideRows(r);
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

// ─── ADD MORE WEEKS ─────────────────────────────────────────
function addMoreWeeks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  const lastDate = sheet.getRange(2, lastCol).getValue();

  let nextFriday;
  if (lastDate instanceof Date) {
    nextFriday = new Date(lastDate);
    nextFriday.setDate(nextFriday.getDate() + 7);
  } else {
    nextFriday = CONFIG.startFriday;
  }

  const newFridays = generateFridays_(nextFriday, 4);
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
    sheet.setColumnWidth(col, 130);
  });

  // Re-merge row 1 for the dates
  const allDatesRange = sheet.getRange(2, 2, 1, sheet.getLastColumn() - 1);
  const allDates = allDatesRange.getValues()[0];
  const themeRowRange = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1);
  themeRowRange.breakApart(); // unmerge first

  let currentMonth = -1;
  let startCol = 2;
  let count = 0;
  for (let i = 0; i < allDates.length; i++) {
    const d = allDates[i];
    if (Object.prototype.toString.call(d) !== "[object Date]" || isNaN(d)) {
      if (count > 0) {
        if (count > 1) sheet.getRange(1, startCol, 1, count).mergeAcross();
        count = 0;
      }
      currentMonth = -1;
      continue;
    }
    const month = d.getMonth();
    if (currentMonth === -1) {
      currentMonth = month;
      count = 1;
      startCol = 2 + i;
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

  // Refresh dropdowns to cover new columns
  refreshAllDropdowns();

  ss.toast("Added 4 more weeks ✅", "Scheduler", 3);
}

// ─── FORMATTING ─────────────────────────────────────────────
function applyFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const numRows = sheet.getLastRow();
  const numCols = sheet.getLastColumn();

  // Alternating row colors for readability (start at row 3)
  for (let r = 3; r <= numRows; r++) {
    const bg = (r % 2 !== 0) ? "#f8f9fa" : "#ffffff";
    sheet.getRange(r, 2, 1, numCols - 1).setBackground(bg);
  }

  // Border around the entire data area
  const fullRange = sheet.getRange(1, 1, numRows, numCols);
  fullRange.setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  // Data cells alignment
  if (numRows > 2 && numCols > 1) {
    const dataRange = sheet.getRange(3, 2, numRows - 2, numCols - 1);
    dataRange.setHorizontalAlignment("center");
    dataRange.setVerticalAlignment("middle");
    dataRange.setFontSize(10);
  }

  // ── Conflict detection via conditional formatting ──
  // These rules use COUNTIF to detect when the same minister name
  // appears multiple times in the same column (same Friday).
  // Rules are ordered by priority: highest severity first.

  const dataRange = sheet.getRange(3, 2, numRows - 2, numCols - 1);

  const isConf3 = 'COUNTIF(B$3:B$' + numRows + ',B3)>=3';
  const isConf2 = 'COUNTIF(B$3:B$' + numRows + ',B3)>=2';

  // Base general rules
  const conf3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(B3<>"", ' + isConf3 + ')')
    .setBackground("#ffcdd2").setFontColor("#000000").setBold(false).setRanges([dataRange]).build();

  const conf2Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(B3<>"", ' + isConf2 + ')')
    .setBackground("#ffe0b2").setFontColor("#000000").setBold(false).setRanges([dataRange]).build();

  const filledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellNotEmpty()
    .setBackground("#c8e6c9").setFontColor("#000000").setBold(false).setRanges([dataRange]).build();

  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground("#fff9c4").setRanges([dataRange]).build();

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
  groupNames.forEach(gName => {
    // "All Roles" and "General Announcement" contradict specific group-based checking, skip them.
    if (gName === "All Roles" || gName === "General Announcement") return;
    const roles = getGroupRoles_(gName);
    const rows = [];
    roles.forEach(role => {
      if (roleToRow[role]) rows.push(roleToRow[role]);
    });

    if (rows.length === 0) return;

    const topLeftCell = "B" + rows[0];
    const ranges = rows.map(r => sheet.getRange(r, 2, 1, numCols - 1));

    const fLeft2 = "IF(COLUMN()>3, OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()-2)=${topLeftCell}`).join(",") + "), FALSE)";
    const fLeft1 = "IF(COLUMN()>2, OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()-1)=${topLeftCell}`).join(",") + "), FALSE)";
    const fRight1 = "OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()+1)=${topLeftCell}`).join(",") + ")";
    const fRight2 = "OR(" + rows.map(r => `INDEX($A$1:$ZZ, ${r}, COLUMN()+2)=${topLeftCell}`).join(",") + ")";
    const isConsecutive = `OR(AND(${fLeft2},${fLeft1}), AND(${fLeft1},${fRight1}), AND(${fRight1},${fRight2}))`;

    const gConf3 = `COUNTIF(B$3:B$${numRows},${topLeftCell})>=3`;
    const gConf2 = `COUNTIF(B$3:B$${numRows},${topLeftCell})>=2`;

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
// When the Ministers or Vacations sheet is edited, automatically refresh dropdowns.
function onEditTrigger(e) {
  try {
    const sheetName = e.source.getActiveSheet().getName();
    if (sheetName === CONFIG.ministersSheetName || sheetName === CONFIG.vacationsSheetName) {
      refreshAllDropdowns();
    }
  } catch (err) {
    // Silently ignore trigger errors
    Logger.log("onEdit error: " + err.message);
  }
}

// ─── CONFLICT CHECKER ───────────────────────────────────────
// Scans the Schedule sheet and reports all ministers who are
// booked into multiple roles on the same Friday.
function checkConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Schedule sheet not found. Run Full Setup first.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 3 || data[0].length < 2) return;

  const dateHeaderRow = data[1];
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  
  // Extract unique Month Year combinations
  const options = [];
  const seen = new Set();
  for (let c = 1; c < dateHeaderRow.length; c++) {
    const d = dateHeaderRow[c];
    if (d instanceof Date) {
      const label = monthNames[d.getMonth()] + " " + d.getFullYear();
      if (!seen.has(label)) {
        options.push(label);
        seen.add(label);
      }
    }
  }

  // If there's only one month, just run it immediately
  if (options.length <= 1) {
    runConflictCheck_("all");
    return;
  }

  // Create a clean, modern HTML modal for the dropdown
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
          cursor: pointer; -webkit-appearance: none; -moz-appearance: none;
        }
        select:focus { border: 2px solid #1a73e8; }
        .buttons { display: flex; justify-content: flex-end; gap: 12px; margin-top: 32px; }
        button { 
          height: 36px; padding: 0 24px; border-radius: 4px; font-size: 14px; font-weight: 500; 
          cursor: pointer; border: 1px solid transparent; transition: background 0.2s;
        }
        .btn-cancel { background: none; color: #1a73e8; border: 1px solid #dadce0; }
        .btn-cancel:hover { background-color: #f8f9fa; }
        .btn-check { background-color: #1a73e8; color: #ffffff; }
        .btn-check:hover { background-color: #1b66c9; box-shadow: 0 1px 2px 0 rgba(60,64,67,0.302), 0 1px 3px 1px rgba(60,64,67,0.149); }
      </style>
    </head>
    <body>
      <label class="label" for="period">Select period to analyze:</label>
      <select id="period">
        <option value="all">Check Entire Schedule</option>
        ${options.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
      </select>
      <div class="buttons">
        <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-check" onclick="startCheck()">Check Now</button>
      </div>
      <script>
        function startCheck() {
          const val = document.getElementById('period').value;
          google.script.run.withSuccessHandler(() => google.script.host.close()).runConflictCheck(val);
        }
      </script>
    </body>
    </html>
  `;

  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(200)
    .setTitle("Conflict Checker");
    
  SpreadsheetApp.getUi().showModalDialog(userInterface, " ");
}

/**
 * Server-side function called by the HTML dialog
 */
function runConflictCheck(filter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  const data = sheet.getDataRange().getValues();
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  
  const allConflicts = [];
  const numRows = data.length;
  const numCols = data[0].length;

  for (let c = 1; c < numCols; c++) {
    const dateHeader = data[1][c];
    
    // Filtering logic
    if (filter !== "all") {
      if (dateHeader instanceof Date) {
        const label = monthNames[dateHeader.getMonth()] + " " + dateHeader.getFullYear();
        if (label !== filter) continue;
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
        let dateStr = (dateHeader instanceof Date) 
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
    ui.alert("✅ No Conflicts", "No conflicts found for " + (filter === "all" ? "the entire schedule" : filter) + ".", ui.ButtonSet.OK);
    return;
  }

  allConflicts.sort((a, b) => b.count - a.count || a.date.localeCompare(b.date));
  let report = "Found " + allConflicts.length + " conflict(s) in " + (filter === "all" ? "the entire schedule" : filter) + ":\n\n";

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
    
    // Set background to alternating zebra stripes
    for (let r = 3; r <= numRows; r++) {
      const bg = (r % 2 !== 0) ? "#f8f9fa" : "#ffffff";
      sheet.getRange(r, 2, 1, numCols - 1).setBackground(bg);
    }
    
    // Ensure font color is black and not bold
    dataRange.setFontColor("#000000");
    dataRange.setFontWeight("normal");
  }

  ss.toast("Formatting cleared to plain zebra stripes ✅", "Scheduler", 3);
}

// ─── PRINTING EXPORT ────────────────────────────────────────
/**
 * Shows a modal dialog to select a 2-month range for printing.
 * The user picks a starting month and the print covers that month + the next.
 * Hidden rows (from group filtering) are automatically excluded.
 */
function exportForPrinting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Schedule sheet not found. Run Full Setup first.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const dateHeaderRow = data[1]; // Row 2 contains dates
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];

  // Extract unique Month-Year labels from the date header
  const options = [];
  const seen = new Set();
  for (let c = 1; c < dateHeaderRow.length; c++) {
    const d = dateHeaderRow[c];
    if (d instanceof Date) {
      const label = monthNames[d.getMonth()] + " " + d.getFullYear();
      if (!seen.has(label)) {
        options.push({ label: label, month: d.getMonth(), year: d.getFullYear() });
        seen.add(label);
      }
    }
  }

  if (options.length === 0) {
    SpreadsheetApp.getUi().alert("No dates found on the Schedule sheet.");
    return;
  }

  // Build the paired options: each starting month + next month
  const pairedOptions = [];
  for (let i = 0; i < options.length; i++) {
    const startOpt = options[i];
    // Find the next month that exists in the schedule
    const nextOpt = (i + 1 < options.length) ? options[i + 1] : null;
    if (nextOpt) {
      pairedOptions.push({
        label: startOpt.label + " – " + nextOpt.label,
        startMonth: startOpt.month,
        startYear: startOpt.year,
        endMonth: nextOpt.month,
        endYear: nextOpt.year
      });
    } else {
      // Last month available — single month print
      pairedOptions.push({
        label: startOpt.label + " (single month)",
        startMonth: startOpt.month,
        startYear: startOpt.year,
        endMonth: startOpt.month,
        endYear: startOpt.year
      });
    }
  }

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
        .hint  { font-size: 12px; color: #80868b; margin-bottom: 16px; }
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
        .btn-print { background-color: #1a73e8; color: #ffffff; }
        .btn-print:hover { background-color: #1b66c9; box-shadow: 0 1px 2px 0 rgba(60,64,67,0.302), 0 1px 3px 1px rgba(60,64,67,0.149); }
      </style>
    </head>
    <body>
      <label class="label" for="range">Select 2-month range to print:</label>
      <p class="hint">Hidden rows (filtered groups) will be excluded automatically.</p>
      <select id="range">
        ${pairedOptions.map((opt, i) => '<option value="' + i + '">' + opt.label + '</option>').join('')}
      </select>
      <div class="buttons">
        <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-print" onclick="startPrint()">🖨 Generate Print Sheet</button>
      </div>
      <script>
        function startPrint() {
          const idx = parseInt(document.getElementById('range').value);
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .runPrintRange(idx);
        }
      </script>
    </body>
    </html>
  `;

  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(420)
    .setHeight(230)
    .setTitle("Print Range");

  SpreadsheetApp.getUi().showModalDialog(userInterface, "🖨 Export for Printing");
}

/**
 * Server-side function called by the print dialog.
 * Receives the index of the selected paired option, re-derives the months,
 * and generates the print sheet.
 */
function runPrintRange(pairedIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.schedulerSheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const dateHeaderRow = data[1];
  const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"];

  // Re-derive the paired options (same logic as exportForPrinting)
  const options = [];
  const seen = new Set();
  for (let c = 1; c < dateHeaderRow.length; c++) {
    const d = dateHeaderRow[c];
    if (d instanceof Date) {
      const label = monthNames[d.getMonth()] + " " + d.getFullYear();
      if (!seen.has(label)) {
        options.push({ label: label, month: d.getMonth(), year: d.getFullYear() });
        seen.add(label);
      }
    }
  }

  const pairedOptions = [];
  for (let i = 0; i < options.length; i++) {
    const startOpt = options[i];
    const nextOpt = (i + 1 < options.length) ? options[i + 1] : null;
    if (nextOpt) {
      pairedOptions.push({
        startMonth: startOpt.month, startYear: startOpt.year,
        endMonth: nextOpt.month, endYear: nextOpt.year
      });
    } else {
      pairedOptions.push({
        startMonth: startOpt.month, startYear: startOpt.year,
        endMonth: startOpt.month, endYear: startOpt.year
      });
    }
  }

  if (pairedIndex < 0 || pairedIndex >= pairedOptions.length) return;
  const chosen = pairedOptions[pairedIndex];

  // Determine which columns (dates) fall within the chosen range
  const colIndices = []; // 1-based column indices on the sheet
  for (let c = 1; c < dateHeaderRow.length; c++) {
    const d = dateHeaderRow[c];
    if (d instanceof Date) {
      const m = d.getMonth();
      const y = d.getFullYear();
      const inRange =
        (y === chosen.startYear && m === chosen.startMonth) ||
        (y === chosen.endYear && m === chosen.endMonth);
      if (inRange) {
        colIndices.push(c + 1); // +1 because data array is 0-based but sheet is 1-based
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

  // 4.5. Re-apply merges in Row 1 of the print sheet (Monthly Theme)
  if (numCols > 2 && values.length > 1) {
    const printDates = values[1].slice(1);
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
  for (let r = 0; r < numRows; r++) {
    printSheet.setRowHeight(headerOffset + r, srcSheet.getRowHeight(rows[r]));
  }
  for (let c = 0; c < numCols; c++) {
    printSheet.setColumnWidth(c + 1, srcSheet.getColumnWidth(cols[c]));
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
