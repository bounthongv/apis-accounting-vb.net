# Phase 2c: Global ADODB Eradication (Global Modules)

## 🎯 Objective
This phase focuses on the **global eradication** of `ADODB` references from shared modules. Unlike Phase 2b (which was form-specific), this phase targets the application's "nervous system."

---

## 🚨 SCENARIO A: The Connection Module (`MDSQLConnection.vb`)
**Status:** ⚠️ **CRITICAL CAUTION**
**Why:** This module hosts `Public CNN As ADODB.Connection`. Hundreds of unmigrated forms depend on this. **DO NOT DELETE IT YET.**

### 🛠️ Execution Prompt for `MDSQLConnection.vb`
```
You are executing Phase 2c for the Core Connection Module.

OBJECTIVE:
Refactor `MDSQLConnection.vb` to support a "Hybrid" state: keep legacy ADODB support alive while adding modern ADO.NET support.

STRICT RULES:
1.  **Imports:** Add `Imports System.Data.SqlClient` at the top.
2.  **Legacy Preservation:** 
    *   KEEP `Public CNN As New ADODB.Connection`. 
    *   Add a comment: `' LEGACY SUPPORT - DO NOT REMOVE UNTIL PHASE 4`.
3.  **Modern Addition:** 
    *   ADD `Public sqlCNN As New SqlConnection`.
4.  **Hybrid Connection Logic:**
    *   Refactor `Sub ConnectSQL()` to open **BOTH** `CNN` (Legacy) and `sqlCNN` (Modern).
    *   Ensure `sqlCNN` uses the connection string format compatible with `System.Data.SqlClient` (refer to `DbHelper.vb` connection string logic if needed).
    *   Use `Try...Catch` for the new `sqlCNN` logic.
5.  **Clean Up:**
    *   Remove unused `ADODB` variables (like `Public Comm`, `Public RSC`) *only if* you are sure they are not used globally. If in doubt, keep them.
    *   Replace `On Error GoTo` with `Try...Catch` where possible.

GOAL:
The application must compile and run exactly as before, but `sqlCNN` must be available for modernized modules to use.
```

---

## 🧹 SCENARIO B: Standard Modules & Helpers
**Target:** `Conection_To_Servee.vb`, `Module1.vb`, etc.
**Status:** ✅ Safe to Eradicate

### 🛠️ Execution Prompt for Standard Modules
```
You are executing Phase 2c for a Standard Module.

OBJECTIVE:
Refactor the specified Global Module to completely remove ADODB and replace it with ADO.NET (`System.Data.SqlClient`).

STRICT RULES:
1.  **Remove Imports:** Remove `Imports ADODB`. Add `Imports System.Data.SqlClient`.
2.  **Connection Objects:** Replace `ADODB.Connection` with `SqlConnection`.
3.  **Recordsets:** Replace `ADODB.Recordset` with `DataTable` (data retrieval) or `SqlCommand` (action).
    *   *Pattern:* `rs.Open sql, conn` → `Dim dt As DataTable = DbHelper.GetDataTable(sql)`
4.  **Transaction Handling:** Replace ADODB transactions with `SqlTransaction`.
5.  **Error Handling:** Wrap database operations in `Try...Catch...Finally` blocks.
6.  **Output:** Return the fully refactored module code.

CHECKLIST VALIDATION:
- [ ] No `ADODB` references remain in this file.
- [ ] `System.Data.SqlClient` is imported.
- [ ] Code compiles.
```

---

## ⚠️ Critical Dependencies
*   **Stop!** Do not refactor `Public conn As ADODB.Connection` in a global module until **ALL** forms using that global variable have been refactored in Phase 2b.
*   **Strategy:** Use Scenario A (Hybrid) for the main connection file. Use Scenario B (Eradicate) for everything else.

## 📝 Done Criteria
- `MDSQLConnection.vb` initializes both `CNN` and `sqlCNN`.
- All other modules in `Module\` folder have 0 references to `ADODB`.