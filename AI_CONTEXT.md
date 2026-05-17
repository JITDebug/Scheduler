# Scheduler App Context for AI Assistants

This document provides a technical overview of the Google Sheets Scheduler Add-on (`SchedulerSetup.gs`) to help AI assistants understand the architecture, data relationships, and complex logic before making further modifications.

## 1. Architecture Overview

This is a monolithic Google Apps Script (`.gs`) designed to be bound to a Google Sheet. 
The entire configuration is driven by the `CONFIG` object at the top of the file.
The application creates and manages 6 different sheets, acting as an automated UI and relational database inside Google Sheets.

### Core Sheets
1. **Schedule**: The main UI. Roles are rows (starting at Row 3), Dates are columns (starting at Column 2). Row 2 contains `Date` objects representing the Fridays.
2. **Ministers**: A matrix of Minister Names (Column A) vs Roles (Columns B onwards). Cells contain `TRUE/FALSE` checkboxes to denote eligibility.
3. **Vacations**: Tracks unavailability. Columns are `Minister Name`, `Start Date`, and `End Date`.
4. **Roles**: The source of truth for the list of available roles.
5. **Role Groups**: Defines categories of roles for the "Filter by Group" UI menu.
6. **Settings**: Used for defining header and footer images for printing.

## 2. Core Mechanisms & Data Flows

### A. Dynamic Dropdowns (`refreshAllDropdowns()`)
This is the most complex logic in the script. It is responsible for creating in-cell dropdowns (`DataValidation`) for the Schedule sheet.
*   **Trigger**: It runs manually via the menu, or automatically via `onEditTrigger(e)` whenever the `Ministers` or `Vacations` sheets are edited.
*   **Logic**:
    1. Reads `Ministers` sheet to map `roleName -> [eligible ministers]`.
    2. Reads `Vacations` sheet and parses Start/End Date objects.
    3. Reads the Date headers from Row 2 of the `Schedule` sheet.
    4. Iterates over every cell (Role Row x Date Column).
    5. Filters the eligible ministers list by checking if the Date falls within any vacation periods for that minister.
    6. Uses `range.setDataValidations([rulesRow])` to bulk-apply the validation rules to an entire row at once. *Note: If no ministers are available for a cell, `null` is passed into the array to clear the validation.*

### B. Conditional Formatting & Conflict Detection (`applyFormatting()`)
The script avoids running heavy iterative conflict-checking on every edit. Instead, it relies on Google Sheets native Conditional Formatting.
*   **Conflicts**: Uses `COUNTIF(B$3:B$MAX, B3)` to check if a person is scheduled multiple times in the same column (same day). Yellow/Orange/Red backgrounds are applied dynamically.
*   **Fatigue Warning**: Checks if a minister is scheduled in the same `Role Group` for 3 consecutive weeks. It uses complex string-building in Apps Script to generate `OR(AND(...))` formulas referencing the surrounding columns, setting the text color to red if true.

### C. Role Syncing (`syncRoles()`)
Instead of wiping the schedule, `syncRoles()` gracefully adds/removes:
*   Columns in the `Ministers` sheet.
*   Rows in the `Schedule` sheet.

### D. Printing Engine (`exportForPrinting()`)
Because Google Sheets printing is limited, the script copies the user's current selection to a temporary `Print - Selection` sheet. It retrieves Header/Footer images configured in the `Settings` sheet and merges them above and below the selected data block, maintaining exact row heights and column widths.

## 3. Important Nuances & Gotchas

*   **Date Handling**: Google Sheets can have weird timezone offset behaviors. In `refreshAllDropdowns()`, the Schedule column dates are forcibly set to `12:00:00` noon to ensure they safely fall between a vacation's `00:00:00` and `23:59:59` boundaries.
*   **Performance**: Apps Script API calls (`getValue()`, `setValue()`) are slow. The script heavily relies on `getValues()`, array manipulation in memory, and `setValues()` or `setDataValidations()` for performance. Avoid making cell-by-cell API calls.
*   **Triggers**: We use an installable trigger (setup via `installTrigger_()` during `setupScheduler()`) instead of a simple `onEdit(e)` function to ensure the script has the necessary permissions to read/write across multiple sheets and display toasts.

## 4. Extending the Application

If you are asked to add a new sheet or setting:
1. Add its name to the `CONFIG` object.
2. If it's a new sheet, add a `buildMySheet_(sheet)` function.
3. Update `setupScheduler()` to construct it on a full reset.
4. If it affects dropdowns or schedule data, hook it into `refreshAllDropdowns()` or `applyFormatting()`.
5. Ensure any new logic respects the batch-processing design pattern (read all data into memory, process, write back in a single call).
