# Phase 1 Report: Platform Stabilization

## Completed Actions
1.  **Project Retargeting:**
    *   Updated `ApBank10.vbproj` to target **.NET Framework 4.8** (was 3.5).
    *   Updated `ToolsVersion` to 12.0 to support modern MSBuild features.
2.  **Configuration Update:**
    *   Updated `app.config` to reflect the new .NET runtime versions (System 4.0.0.0, Microsoft.VisualBasic 10.0.0.0).
3.  **Dependency Management:**
    *   Created a local `lib` directory.
    *   Copied critical dependencies (`MySql.Data.dll`, `Newtonsoft.Json.dll`, `AxInterop.VSFlex8U.dll`, `Interop.VSFlex8U.dll`, `Interop.MSDATASRC.dll`) to `lib` to ensure build stability.
    *   Updated `ApBank10.vbproj` with `HintPath`s pointing to the `lib` folder for these assemblies.

## Required User Actions (Manual Steps)
The following steps require software installation on the development machine and could not be automated:

1.  **Install Crystal Reports for Visual Studio:**
    *   Download and install the latest **SAP Crystal Reports for Visual Studio (SP30 or later)**.
    *   The current project still references version `10.5.3700.0`. Upon opening the solution in Visual Studio 2022, you may be prompted to upgrade these references to version `13.0.x.x`. **Allow this upgrade.**
2.  **Verify Build:**
    *   Open `ApBank10.sln` in Visual Studio 2022.
    *   Rebuild the solution.
    *   Check for any missing reference errors (specifically related to Crystal Reports).

## Next Steps (Phase 2)
Once the application builds successfully in .NET 4.8:
1.  **Replace VSFlexGrid:** Begin replacing the ActiveX `AxVSFlexGrid` control with a .NET native grid (DataGridView or 3rd party).
2.  **Replace ADODB:** Refactor the data access layer to use `System.Data.SqlClient`.
