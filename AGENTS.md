# AGENTS.md - Development Guidelines for APBank10 Accounting System

## Project Overview
This is a legacy VB.NET Windows Forms accounting application targeting .NET Framework 4.8. The system handles multi-database operations (SQL Server, MySQL, Access) with extensive Crystal Reports integration.

## Build Commands

### Primary Build Commands
```bash
# Build Debug Configuration
msbuild ApBank10.vbproj /p:Configuration=Debug /p:Platform="AnyCPU"

# Build Release Configuration  
msbuild ApBank10.vbproj /p:Configuration=Release /p:Platform="AnyCPU"

# Clean and Rebuild
msbuild ApBank10.vbproj /t:Clean /p:Configuration=Debug
msbuild ApBank10.vbproj /t:Rebuild /p:Configuration=Debug
```

### Visual Studio Build
- Use Visual Studio 2022 for development
- Build Configuration: Debug (default) or Release
- Platform Target: x86
- Output Path: `bin\Debug\` or `bin\Release\`

## Testing Approach
**No Automated Testing Framework** - This legacy application relies on manual testing only.

### Manual Testing Guidelines
- Test database operations with both SQL Server and MySQL
- Verify Crystal Reports generate correctly
- Test form workflows and data validation
- Use temporary tables (TEST_ABC, TEST_MM) for report calculation testing
- Validate multi-language support (Lao/English)

### Test Database Setup
- SQL Server: Primary testing database
- MySQL: Secondary database compatibility testing  
- Access: Local configuration database (Connection.mdb)

## Code Style Guidelines

### VB.NET Compiler Settings
```xml
<OptionExplicit>On</OptionExplicit>
<OptionStrict>Off</OptionStrict>  <!-- Required for legacy compatibility -->
<OptionInfer>On</OptionInfer>
```

### Naming Conventions
```vb
' Hungarian notation for variables
Dim s As String = 0
Dim w As Integer = 0  
Dim k As String
Dim sql As String

' Module-level variables with prefixes
Public MDSeriel As String
Public MDStarDate As String
Public MuGen_jn As String

' Form naming: Fm + Description
Public Class FmMain
Public Class FmAmtStatus

' Report naming: Cry + Description
Public Class CryRpt_BLS
Public Class CryInvoiceTAX
```

### Import Organization
```vb
' Standard imports (order as shown)
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Linq
Imports System.Xml.Linq

' Third-party imports
Imports AxVSFlex8U
Imports ADODB

' Project imports (custom modules)
Imports ApPBank10.Module
```

### Database Access Patterns
```vb
' Standard connection pattern
Public Sub ConnectSQL()
    With CNN
        On Error GoTo hang
        If .State = ConnectionState.Open Then .Close()
hang:
        If Err.Number = 0 Then
            ' Success logic
        Else
            ' Error handling
        End If
    End With
End Sub

' Data loading pattern
Public Sub LoadSqlData(ByVal StrSql As String, ByVal Rs As ADODB.Recordset)
    With Rs
        If .State = ConnectionState.Open Then .Close()
        .ActiveConnection = CNN
        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
        .CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        .LockType = ADODB.LockTypeEnum.adLockReadOnly
        .Open(StrSql)
        .Requery()
    End With
End Sub

' Recordset processing pattern
LoadSqlData("SELECT * FROM Table1 WHERE Id = '" & id & "'", RSC)
With RSC
    Do Until .EOF = True
        fieldName.Text = .Fields("FieldName").Value
        .MoveNext()
    Loop
End With
```

### Form Structure Patterns
```vb
' Standard form load pattern
Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Call LoadSQL()
    Call LoadFirstRecord()
    ' Additional initialization
End Sub

' Error handling with GoTo (legacy compatibility)
Public Sub SomeMethod()
    On Error GoTo ErrorHandler
    ' Method logic
    Exit Sub
ErrorHandler:
    ' Error handling logic
End Sub
```

### File Organization
```
AAA/         - Legacy forms and reports (51 files)
ACC_NEW/     - New accounting modules (96 files)
AP_NEW/      - Asset management modules (79 files)
Asset/       - Asset-related Crystal Reports (39 files)
Sonexay/     - Main application modules (74 forms, 39 reports)
Module/      - Shared modules and utilities (17 files)
RPT_NEW/     - New reporting modules (37 files)
RPT2022/     - 2022 reporting updates (9 files)
```

## Error Handling Standards
- Use `On Error GoTo` for backward compatibility
- Implement comprehensive error handling in database operations
- Log errors appropriately for debugging
- Use Try-Catch for newer .NET operations where possible

## Database Standards
- **Primary**: SQL Server via ADODB (SQLOLEDB provider)
- **Secondary**: MySQL via MySql.Data.dll
- **Configuration**: Access database (Connection.mdb)
- Use parameterized queries where possible
- Implement proper connection string management
- Handle multi-database compatibility

## UI Framework Guidelines
- **Primary**: Windows Forms
- **Grid Controls**: Migrating from VSFlexGrid (ActiveX) to DataGridView
- **Reports**: Crystal Reports v10.5 (100+ .rpt files)
- **Images**: Store in Resources/ folder, manage via SQL

## Migration Priorities
1. **ActiveX Removal**: Replace VSFlexGrid with DataGridView (10/53 forms completed)
2. **Framework Updates**: Maintain .NET Framework 4.8 compatibility
3. **Database Modernization**: Prepare for future ADO.NET migration
4. **Web Application**: Flutter + FastAPI replacement in progress

## Security Considerations
- Implement proper user authentication and role-based access
- Secure database connection strings
- Validate all user inputs
- Use proper SQL injection prevention
- Handle sensitive data appropriately

## Internationalization
- **Primary Language**: Lao (MuLng = "L")
- **Secondary**: English support
- Use MuLangauge.vb for localization
- Dynamic text loading capability
- Multi-language form labels and messages

## Development Tools
- **IDE**: Visual Studio 2022
- **Source Control**: Git (check for existing .gitignore)
- **Dependencies**: Managed via lib/ folder
- **Build System**: MSBuild

## Key Dependencies
```xml
<!-- Core Dependencies -->
<Reference Include="adodb" />
<Reference Include="CrystalDecisions.CrystalReports.Engine" />
<Reference Include="System.Windows.Forms" />

<!-- Legacy Components -->
<Reference Include="AxInterop.VSFlex8U" />
<Reference Include="Interop.VSFlex8U" />

<!-- Database Support -->
<Reference Include="MySql.Data" />

<!-- Modern Components -->
<Reference Include="Newtonsoft.Json" />
<Reference Include="Microsoft.ReportViewer.Common" />
```

## Performance Guidelines
- Optimize database queries with proper indexing
- Use client-side cursors for read-only operations
- Implement proper form loading sequences
- Manage Crystal Reports memory usage
- Optimize grid control performance

## Code Quality Standards
- Maintain Option Strict Off for legacy compatibility
- Use comprehensive error handling
- Follow established naming conventions
- Document complex business logic
- Test database operations thoroughly

## Important Notes
- **Never enable Option Strict** - would break legacy code compatibility
- **Maintain Hungarian notation** - consistent with existing codebase
- **Preserve ADODB patterns** - database layer depends on this architecture
- **Test multi-database scenarios** - SQL Server and MySQL compatibility required
- **Respect Crystal Reports structure** - extensive reporting depends on existing .rpt files

## Byterover MCP Integration
When using Byterover MCP tools:
- Store knowledge about architectural decisions and patterns
- Retrieve relevant context before making changes
- Document error solutions and debugging techniques
- Share reusable code patterns and utility functions