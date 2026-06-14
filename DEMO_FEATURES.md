# ⏱️ Ministry Scheduler Demo Guide & Feature Showcase

Welcome to the **Church Ministry Scheduler**! This guide is designed to walk you through a live demonstration of the scheduler's automated features. It is structured to help you explain *what* each feature does and *how* to showcase it in a beginner-friendly way during a live presentation on a pre-configured Google Sheet.

---

## 🚀 The Core Vision
The Church Ministry Scheduler is a smart, automated scheduling system built inside Google Sheets. It replaces manual, error-prone spreadsheets with built-in logic that:
*   Prevents double-booking (conflict detection).
*   Avoids volunteer burnout (fatigue warnings).
*   Respects unavailability/vacations automatically.
*   Adapts dynamically when roles or people are added.
*   Generates beautiful, print-ready sheets with a single click.

---

## 💎 The Showcase Sheet Structure
The system is built on **6 interconnected sheets**:

| Sheet | Purpose | Demonstration Value |
| :--- | :--- | :--- |
| **`Schedule`** | The main dashboard where dates run horizontally as columns and roles run vertically as rows. | The visual "command center" showing colors, warnings, and schedules. |
| **`Ministers`** | A matrix of people and checkboxes representing which roles they are eligible to fill. | Show how easy it is to manage staff/volunteer credentials. |
| **`Vacations`**| Unavailability tracking sheet (Name, Start Date, End Date). | Show how the dropdowns filter out unavailable people. |
| **`Roles`** | The master list of all positions (e.g., Preacher, Guitarist, Usher). | Simple single-column data entry that propagates changes. |
| **`Role Groups`** | Custom groups of roles (e.g., Music Ministry, Pulpit Ministry) used for filtering. | Keep the schedule clean for larger teams. |
| **`Settings`** | Storage for Header and Footer images for printing templates. | Show customization and church branding. |

---

## 🌟 Feature Breakdown (What They Do & How to Demo)

Here is a breakdown of every operational feature available in the **⏱ Scheduler** menu:

### 1. 👥 Smart Minister Dropdowns (With Vacation Filtering)
*   **What it does:** Ensures that when scheduling a role on a specific day, you only see dropdown options of people who:
    1. Are checked as eligible for that role in the `Ministers` sheet.
    2. Do NOT have a vacation scheduled on that date in the `Vacations` sheet.
*   **How to Demo:** 
    1. Go to the `Schedule` sheet and open the dropdown for **Worship Leader 1** on **May 8**. Point out the available names (e.g., *David, Jewell*).
    2. Go to the `Vacations` sheet and add a vacation for **David** covering **May 8** (e.g., Start: `5/1/2026`, End: `5/10/2026`).
    3. Click `⏱ Scheduler ▸ 👥 Update Minister Dropdowns` (or let the auto-trigger run).
    4. Go back to the `Schedule` sheet: David is now automatically filtered out of the May 8 dropdown!
*   **Key Talking Point:** *"Say goodbye to checking external calendars. The sheet dynamically recalculates who is actually available on any given date."*

### 2. ⚠️ Visual Conflict Detection & Fatigue Warnings
*   **What it does:** Uses a high-visibility, automated color-coding system to flag scheduling issues:
    *   🟩 **Green Background:** Assigned normally (no conflicts).
    *   🟨 **Yellow Background:** Role is empty/unfilled.
    *   🟧 **Orange Background:** Double-booked (assigned to 2 roles on the same day).
    *   🟥 **Red Background:** Triple+ booked (assigned to 3 or more roles on the same day).
    *   🔴 **Red Text:** Fatigue warning! The person is scheduled in the same ministry group for 3 or more consecutive weeks.
*   **How to Demo:** 
    1. Assign the same person (e.g., **David**) to two different roles on the same Friday column. Watch the background instantly turn **Orange**!
    2. Assign them to a third role on that same day. Watch it turn **Red**!
    3. Assign someone to **Worship Leader 1** for three consecutive Fridays. Show the text color turning **Red** as a fatigue warning.
    4. Hover over cell **A2** (the top-left date cell) to show the built-in **Color Legend Note**.
*   **Key Talking Point:** *"The scheduler alerts you visually when you double-book someone or overwork a volunteer across consecutive weeks."*

### 3. 📂 Dynamic Group Filtering
*   **What it does:** Allows users to filter the main Schedule sheet to show only specific ministries (e.g., only "Music Ministry" or "Ushering Ministry" rows), hiding the rest.
*   **How to Demo:** 
    1. Go to `⏱ Scheduler ▸ 📂 Filter by Group ▸ Music Ministry`. Show how all other rows are neatly hidden, leaving only the music roles.
    2. Choose `⏱ Scheduler ▸ 📂 Filter by Group ▸ 👁 Show All Roles` to bring them all back.
*   **Key Talking Point:** *"Large schedules can be overwhelming. Group filtering allows ministry heads to focus solely on their own teams without cluttering the screen."*

### 4. 🔄 Intelligent Role Syncing
*   **What it does:** When you add a new role or delete one in the `Roles` sheet, running this sync propagates it to the `Ministers` sheet (adding/deleting columns) and the `Schedule` sheet (adding/deleting rows) without wiping any of your existing schedule assignments.
*   **How to Demo:** 
    1. Go to the `Roles` sheet and add a new role at the bottom (e.g., `Technical Director`).
    2. Click `⏱ Scheduler ▸ 🔄 Sync New/Removed Roles`.
    3. Show the user that `Technical Director` now has a new column in `Ministers` and a new row in `Schedule`, ready for assignments!
*   **Key Talking Point:** *"Your team structures change. This sync tool updates your database structure seamlessly without risking your historical schedule data."*

### 5. ➕ Add More Weeks
*   **What it does:** Extends the schedule by adding 4 new Friday columns, automatically copying the formatting, monthly theme structure, date calculations, and smart dropdown validations.
*   **How to Demo:** 
    1. Scroll to the far right of the `Schedule` sheet.
    2. Click `⏱ Scheduler ▸ ➕ Add 4 More Weeks`.
    3. Show how the calendar extends by a month with everything pre-configured.
*   **Key Talking Point:** *"Never start a sheet from scratch. Seamlessly grow your scheduling cycle 4 weeks at a time."*

### 6. 🖨 Custom Export for Printing
*   **What it does:** Copies the current view (respecting any group filters and date ranges) to a clean, isolated sheet, applying a beautiful header and footer template defined in the Settings sheet.
*   **How to Demo:** 
    1. Filter the schedule to a specific group (e.g., `Music Ministry`).
    2. Click `⏱ Scheduler ▸ 🖨 Export for Printing`. Select the 2-month range from the custom dialog.
    3. A new sheet `Print - Selection` appears with a branded header/footer, ready to print or save as PDF!
*   **Key Talking Point:** *"No need to copy-paste or clean up formatting. Generate a publication-ready schedule in seconds."*

---

## 🎬 Presenter's 5-Minute Demo Flow

Follow this sequence to deliver a high-impact, smooth demonstration on a pre-configured sheet:

```
  [1. Intro & Problem] 
         │
         ▼
  [2. Schedule & Vacations] (Show how David is filtered out when on vacation)
         │
         ▼
  [3. Create a Conflict] (David assigned to 2 roles -> Orange BG. 3 weeks -> Red Text)
         │
         ▼
  [4. Filter by Group] (Show only Music Ministry roles to declutter)
         │
         ▼
  [5. Export to PDF/Print] (Generate beautiful print layout sheet with 1 click)
         │
         ▼
  [6. Add a New Role & Sync] (Add 'Video Editor' in Roles -> Sync to Ministers/Schedule)
```

1.  **Introduce the Problem:** *"Scheduling volunteers is a manual nightmare of conflicting dates, vacations, and layout updates. Here is a pre-configured central sheet shared by our team."*
2.  **Assign someone & Show Vacations:** Assign a minister on the `Schedule` sheet. Point out that the dropdown only lists eligible volunteers. Add a vacation for David on the `Vacations` sheet, run `Update Minister Dropdowns`, and show David filtered out of the schedule dropdown on his vacation date.
3.  **Create a Conflict & Fatigue:** Assign "David" to two roles on the same day. Point out the Orange warning. Assign him 3 weeks in a row to show the Red Text fatigue warning.
4.  **Filter by Group:** Filter the screen to "Music Ministry". Point out how we instantly zoom in on one team's view.
5.  **Export for Printing:** Click `🖨 Export for Printing`. Select the date range from the dialog and show the gorgeous printed sheet with headers/footers.
6.  **Add a Role & Sync:** Add a new role "Video Editor" in the `Roles` sheet, run `🔄 Sync New/Removed Roles`, and show it appearing instantly as a new column in `Ministers` and a new row in `Schedule` without losing any of the existing assignments.
