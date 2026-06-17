# ⏱ Church Ministry Scheduler

A powerful, automated scheduling system built with Google Apps Script for Google Sheets. Designed to manage church ministry rotations, prevent scheduling conflicts, and generate beautiful, print-ready schedules.

## 🌟 Key Features

- **Automated Setup**: Generates all necessary sheets (Schedule, Ministers, Roles, Role Groups, Settings, Vacations, Events) with a single click.
- **Conflict Detection**: Built-in logic to prevent assigning the same person to multiple roles on the same day. Filterable by Year/Month.
- **Ministry Group Filtering**: Easily toggle visibility between different groups (e.g., Music Ministry, Ushering, Children's Ministry).
- **In-Cell Image Printing**: Export specific selections to a dedicated print sheet with custom header and footer branding.
- **Zebra Striping**: Quickly strip formatting and apply a clean, high-readability "zebra stripe" pattern for printing.
- **Dynamic Month Archiving**: Explicitly archive or unarchive targeted months. Archived months are excluded from all real-time evaluations (like conflict detection) to drastically improve Google Sheets performance.
- **Event Headers**: Maintain an `Events` sheet so matching schedule dates show a two-line date and event name header.
- **Dynamic Sync**: Add or remove roles/ministers and sync the dropdown menus across the entire schedule instantly.

## 🚀 Getting Started

### 1. Installation
1. Create a new Google Sheet.
2. Go to **Extensions ▸ Apps Script**.
3. Delete any code in the editor and paste the contents of `SchedulerSetup.gs`.
4. Click **Save** and name the project "Ministry Scheduler".
5. Refresh your Google Sheet.

### 2. Initial Setup
1. In your Google Sheet, find the new **⚙️ Configs** menu.
2. Select **⚙️ Full Setup (WARNING: DELETES ALL DATA)**.
3. Authenticate the script when prompted.
4. The script will generate 7 sheets: `Schedule`, `Ministers`, `Roles`, `Role Groups`, `Settings`, `Vacations`, and `Events`.

### 3. Customization
- **Ministers**: Add your team members and their eligible roles in the `Ministers` sheet.
- **Roles**: List all roles in the `Roles` sheet.
- **Groups**: Define which roles belong to which ministry in the `Role Groups` sheet.
- **Settings**: Insert your church logo or banner into cells **B2** and **B3** (Insert ▸ Image ▸ Image in cell).
- **Events**: Add event names in column A and event dates in column B. Matching Schedule headers will show the date plus event name.

## 🖨 How to Print
1. Highlight the cells you want to print.
2. Go to **⏱ Scheduler ▸ 🖨 Export for Printing**.
3. A new sheet `Print - Selection` will be created with your headers and footers.
4. Go to **File ▸ Print** and set to "Fit to Width".

## 🛠 Maintenance
- **Add Month**: Use **➕ Add Month** to dynamically extend your schedule for the entirety of the next month (automatically calculating 4 or 5 weeks).
- **Archive Month**: Use **📦 Archive Month...** to selectively hide past schedules, clean up your view, and instantly boost Google Sheets performance by excluding those columns from background evaluations.
- **Sync Roles**: If you add new roles to the `Roles` sheet, use **🔄 Sync New/Removed Roles** to update the schedule layout.
- **Refresh Dropdowns**: Use **👥 Update Minister Dropdowns** to sync the latest list of eligible ministers to each row.

## 📄 License
This project is open-source and free to use for ministry purposes.
