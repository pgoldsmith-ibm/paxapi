# Universal Report API functions

Universal Report functions can be used to create and interact with Universal Report worksheets. Universal Reports are accessible through the `Reporting.UniversalReports` manager object, which returns individual `UniversalReport` objects.

The Universal Report manager functions that are exposed through the IBM® Cognos® automation objects are:

## Create

> Example

```vb
Public Sub CreateUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Create("http://myserver.ibm.com", "Planning Sample", "plan_BudgetPlan", "Goal Input")
End Sub
```

Create generates a Dynamic Universal Report based on the host system URL, server name, cube name, and view name. The report is inserted on a new sheet in the active workbook.

### Syntax

The following string is the syntax for the Create method.

`Reporting.UniversalReports.Create "<host system URL>", "<server name>", "<cube name>", "<view name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Universal Report is to be created from. | Alphanumeric string
server name | Name of the server which the Universal Report is to be created from. | Alphanumeric string
cube name | Name of the cube which the Universal Report is to be created from. | Alphanumeric string
view name | Name of the view which the Universal Report is to be created from. | Alphanumeric string

### Return value

Data type: UniversalReport object

### Errors

If any argument is invalid, a runtime error is raised. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad host system URL | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad server name | Runtime error `-2146233832`: *Not Found*
Bad cube name | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad view name | Runtime error `-2146233832`: *'\<name\>' can not be found in collection of type 'View'.*

## CreateStatic

> Example

```vb
Public Sub CreateStaticUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.CreateStatic("http://myserver.ibm.com", "Planning Sample", "plan_BudgetPlan", "Goal Input")
End Sub
```

CreateStatic generates a Static Universal Report based on the host system URL, server name, cube name, and view name. The report is inserted on a new sheet in the active workbook.

### Syntax

The following string is the syntax for the CreateStatic method.

`Reporting.UniversalReports.CreateStatic "<host system URL>", "<server name>", "<cube name>", "<view name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Universal Report is to be created from. | Alphanumeric string
server name | Name of the server which the Universal Report is to be created from. | Alphanumeric string
cube name | Name of the cube which the Universal Report is to be created from. | Alphanumeric string
view name | Name of the view which the Universal Report is to be created from. | Alphanumeric string

### Return value

Data type: UniversalReport object

### Errors

If any argument is invalid, a runtime error is raised. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad host system URL | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad server name | Runtime error `-2146233832`: *Not Found*
Bad cube name | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad view name | Runtime error `-2146233832`: *'\<name\>' can not be found in collection of type 'View'.*

## CreateFromMDX

> Example

```vb
Public Sub CreateFromMDXWithContext()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.CreateFromMDX( _
        "http://myserver.ibm.com", _
        "Planning Sample", _
        "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].[Revenue]} ON 1, " & _
        "{[plan_time].[plan_time].[2004]} ON 0 FROM [plan_BudgetPlan] " & _
        "WHERE ([plan_department].[plan_department].[Total Organization], " & _
        "[plan_business_unit].[plan_business_unit].[Total Business Unit]," & _
        "[plan_exchange_rates].[plan_exchange_rates].[actual]," & _
        "[plan_version].[plan_version].[FY 2004 Budget])")
End Sub
```

CreateFromMDX generates a Dynamic Universal Report based on a host system URL, server name, and MDX statement. The report is inserted on a new sheet in the active workbook.

<aside class="notice">
Unlike other reporting modes, Universal Reports created from MDX do not automatically generate slicer context. Any dimensions you want to appear as slicers must be explicitly specified in the <code>WHERE</code> clause of the MDX query.
</aside>

### Syntax

The following string is the syntax for the CreateFromMDX method.

`Reporting.UniversalReports.CreateFromMDX "<host system URL>", "<server name>", "<MDX statement>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Universal Report is to be created from. | Alphanumeric string
server name | Name of the server which the Universal Report is to be created from. | Alphanumeric string
MDX statement | MDX statement which the Universal Report is to be created from. | Alphanumeric string

### Return value

Data type: UniversalReport object

### Errors

If any argument is invalid, a runtime error is raised. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad host system URL | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad server name | Runtime error `-2146233832`: *Not Found*
Bad MDX statement | Runtime error `-2147024809`: *Unreducible state* (MDX parse error — check the MDX statement for syntax errors)

## CreateStaticFromMDX

> Example

```vb
Public Sub CreateStaticFromMDXWithContext()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.CreateStaticFromMDX( _
        "http://myserver.ibm.com", _
        "Planning Sample", _
        "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].[Revenue]} ON 1, " & _
        "{[plan_time].[plan_time].[2004]} ON 0 FROM [plan_BudgetPlan] " & _
        "WHERE ([plan_department].[plan_department].[Total Organization], " & _
        "[plan_business_unit].[plan_business_unit].[Total Business Unit]," & _
        "[plan_exchange_rates].[plan_exchange_rates].[actual]," & _
        "[plan_version].[plan_version].[FY 2004 Budget])")
End Sub
```

CreateStaticFromMDX generates a Static Universal Report based on a host system URL, server name, and MDX statement. The report is inserted on a new sheet in the active workbook.

<aside class="notice">
Unlike other reporting modes, Universal Reports created from MDX do not automatically generate slicer context. Any dimensions you want to appear as slicers must be explicitly specified in the <code>WHERE</code> clause of the MDX query.
</aside>

### Syntax

The following string is the syntax for the CreateStaticFromMDX method.

`Reporting.UniversalReports.CreateStaticFromMDX "<host system URL>", "<server name>", "<MDX statement>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Universal Report is to be created from. | Alphanumeric string
server name | Name of the server which the Universal Report is to be created from. | Alphanumeric string
MDX statement | MDX statement which the Universal Report is to be created from. | Alphanumeric string

### Return value

Data type: UniversalReport object

### Errors

If any argument is invalid, a runtime error is raised. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad host system URL | Runtime error `-2147467261`: *Value cannot be null. Parameter name: datasource*
Bad server name | Runtime error `-2146233832`: *Not Found*
Bad MDX statement | Runtime error `-2147024809`: *Unreducible state* (MDX parse error — check the MDX statement for syntax errors)

## Get

> Example

```vb
Public Sub GetUR()
    On Error GoTo Handler:
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    MsgBox "Found report on cube: " & oReport.Cube
    Exit Sub
Handler:
    CognosOfficeAutomationObject.TraceError "UniversalReports.Get failed: " & Err.Description
End Sub
```

Get retrieves an existing Universal Report object from a specified workbook and sheet by its report ID. A runtime error is raised if any argument is invalid or the report cannot be found.

### Syntax

The following string is the syntax for the Get method.

`Reporting.UniversalReports.Get "<report ID>", "<book name>", "<sheet name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
report ID | The ID of the Universal Report to retrieve. | String
book name | The name of the workbook that contains the Universal Report. | String
sheet name | The name of the worksheet that contains the Universal Report. | String

### Return value

Data type: UniversalReport object.

### Errors

A runtime error is raised for all invalid inputs. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad book name | Runtime error `-2147024809`: *No open workbook named \<name\>.*
Bad sheet name | Runtime error `-2147024809`: *No sheet named \<name\> in workbook \<book\>.*
Bad report ID | Runtime error `-2147024809`: *No Universal Report with ID '\<id\>' found on sheet '\<sheet\>' in workbook '\<book\>'.*

## GetReportsFromSheet

> Example

```vb
Public Sub ListSheetReports()
    Dim col As Object
    Set col = Reporting.UniversalReports.GetReportsFromSheet(ThisWorkbook.Name, ActiveSheet.Name)
    Dim i As Integer
    For i = 0 To col.Count - 1
        MsgBox col.Item(i).Cube
    Next i
End Sub
```

GetReportsFromSheet returns a collection of all Universal Report objects found on the specified sheet. Returns an empty collection (not an error) if the sheet exists but contains no Universal Reports.

### Syntax

`Reporting.UniversalReports.GetReportsFromSheet "<book name>", "<sheet name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
book name | The name of the workbook to search. | String
sheet name | The name of the worksheet to search. | String

### Return value

Data type: Collection of UniversalReport objects.

### Errors

A runtime error is raised if the book or sheet cannot be resolved. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad book name | Runtime error `-2147024809`: *No open workbook named \<name\>.*
Bad sheet name | Runtime error `-2147024809`: *No sheet named \<name\> in workbook \<book\>.*

## GetReportsFromBook

> Example

```vb
Public Sub ListBookReports()
    Dim col As Object
    Set col = Reporting.UniversalReports.GetReportsFromBook(ThisWorkbook.Name)
    Dim report As Variant
    For Each report In col.items()
        MsgBox report.Sheet & ": " & report.Cube
    Next report
End Sub
```

GetReportsFromBook returns a collection of all Universal Report objects found across all sheets in the specified workbook.

### Syntax

`Reporting.UniversalReports.GetReportsFromBook "<book name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
book name | The name of the workbook to search. | String

### Return value

Data type: Collection of UniversalReport objects.

### Errors

A runtime error is raised if the book cannot be resolved. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad book name | Runtime error `-2147024809`: *No open workbook named \<name\>.*

---

## Universal Report object

> Example

```vb
Public Sub GetURProperties()
    On Error GoTo Handler:
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    MsgBox "Book: " & oReport.Book & vbNewLine & _
           "Sheet: " & oReport.Sheet & vbNewLine & _
           "Id: " & oReport.Id & vbNewLine & _
           "Cube: " & oReport.Cube & vbNewLine & _
           "DataSource: " & oReport.DataSource
    Exit Sub
Handler:
    CognosOfficeAutomationObject.TraceError "UniversalReports.Get failed: " & Err.Description
End Sub
```

The following functions and properties are available on a `UniversalReport` object returned by the manager methods above.

### Properties

Property | Description | Data type
--------- | ------- | -----------
Book | The name of the workbook that contains the Universal Report. | String
Sheet | The name of the worksheet that contains the Universal Report. | String
Id | The unique identifier of the Universal Report. | String
Cube | The name of the cube that the Universal Report is based on. | String
DataSource | The name of the data source (server) that the Universal Report is connected to. | String

## Commit (Universal Report)

> Example

```vb
Public Sub CommitUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    oReport.Commit
End Sub
```

Commit writes back any pending data changes made in the Universal Report to the TM1 server.

### Syntax

The following string is the syntax for the Commit method.

`<UniversalReport>.Commit`

### Errors

Commit never raises a runtime error. If the report has not been refreshed there is nothing to commit and the call is a no-op. If the sheet contains invalid values, the report silently restores those cells to their correct state (as of the last refresh) rather than writing the invalid values to the server.

## SetSlicer (Universal Report)

> Example

```vb
Public Sub SetSlicerUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    oReport.SetSlicer "plan_version", "Actual"
End Sub
```

SetSlicer updates the slicer value for a given dimension in the Universal Report and triggers a refresh. The member name is resolved to its unique name (MUN) against the connected server before the slicer formula is updated.

### Syntax

The following string is the syntax for the SetSlicer method.

`<UniversalReport>.SetSlicer "<dimension name>", "<member name>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
dimension name | The name of the cube dimension whose slicer value is to be updated. | String
member name | The display name of the member to set as the slicer value. | String

### Errors

A runtime error is raised if either argument is invalid. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Bad dimension name | Runtime error `-2147024809`: *Slicer dimension '\<name\>' not found.*
Bad member name | Runtime error `-2146233832`: *'\<name\>' can not be found in collection of type 'Member'.*

## GetCellAddressFromMUN

> Example

```vb
Public Sub FindMUNCell()
    On Error GoTo ErrHandler
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    Dim addr As String
    addr = oReport.GetCellAddressFromMUN("[plan_version].[plan_version].[Actual]", True)
    If addr <> "" Then
        MsgBox "Member cell is at: " & addr
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub
```

GetCellAddressFromMUN returns the A1-style Excel cell address of the header cell whose unique member name (MUN) matches the supplied value on the specified axis. Both member ID-based MUNs (e.g. `[Region].[Region].[North America]`) and alias-based MUNs (e.g. `[Region].[Region].[Amérique du Nord]`) are supported. Returns an empty string if no matching member is found.

### Syntax

The following string is the syntax for the GetCellAddressFromMUN method.

`<UniversalReport>.GetCellAddressFromMUN "<MUN>", <isRow>`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
MUN | The unique member name to locate in the report header cache. Accepts both member ID-based and alias-based MUNs. | String
isRow | `True` to search the row axis; `False` to search the column axis. | Boolean

### Return value

Data type: String. The A1-style cell address, or an empty string if not found.

### Errors

A runtime error is raised if the report has not been refreshed. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details. For all other failure cases, the method returns an empty string — always check the return value before using it.

Case | Result
-----|-------
MUN not found in report | Returns an empty string. No error is raised.
Mismatched `isRow` boolean (searching the wrong axis) | Returns an empty string. No error is raised.
Report not yet refreshed | Runtime error `-2146233079`: *The Universal Report must be fully refreshed before this operation can be performed. Wait for the report to finish loading and try again.*

## InsertUserRow

> Example

```vb
Public Sub InsertUserRowUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    ' Pass the header cell of the anchor row as a Range
    ' Pass "" for header or expression to use defaults
    oReport.InsertUserRow Range("E53"), 2, "Revenue", "[plan_chart_of_accounts].[plan_chart_of_accounts].[Revenue]"
End Sub
```

InsertUserRow inserts a custom user-defined row into a Dynamic Universal Report at a position relative to the row that contains the supplied Range. Accepts a header label and an MDX expression to populate the row data cells — pass an empty string for either to use the default (no visible header, cells display as blank). The report must be fully refreshed before calling this method.

<aside class="notice">
InsertUserRow is only supported for Dynamic Universal Reports. Calling this method on a Static Universal Report will raise an error.
</aside>

### Syntax

The following string is the syntax for the InsertUserRow method.

`<UniversalReport>.InsertUserRow <range>, <position>, "<header>", "<expression>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
range | An Excel Range whose first cell falls within the row axis of the report. The anchor member is resolved from this cell using the report's header cache. | Range
position | The position at which to insert the row. Accepted values: `1` (Before selection), `2` (After selection), `3` (Start of hierarchy), `4` (End of hierarchy). | Integer
header | The header label for the new row. Pass an empty string to insert the row without a visible header. | String
expression | An MDX value expression evaluated by the TM1 server to populate each data cell in the new row. Any valid MDX value expression is accepted — for example, a member reference such as `[Measures].[Revenue]`, or a calculated expression such as `[Measures].[Revenue] / [Measures].[Units]`. Pass an empty string and the cells will display as blank. | String

### Errors

A runtime error is raised for all invalid inputs. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Range outside report bounds | Runtime error `-2147024809`: *The specified range does not correspond to a header cell within the report. Verify that the cell is inside the row or column axis of the universal report.*
Invalid `position` value | Runtime error `-2147024809`: *Invalid position '\<value\>'. Expected 1 (before), 2 (after), 3 (start of hierarchy), or 4 (end of hierarchy).*
Called on a Static Universal Report | Runtime error `-2146233079`: *Insert user content is only supported for Dynamic Universal Reports.*
Report not fully refreshed | Runtime error `-2146233079`: *The Universal Report must be fully refreshed before inserting user content. Wait for the report to finish loading and try again.*
Duplicate header (same as an existing header on the axis) | Runtime error `-2147024809`: *The header '\<name\>' already exists on the row axis. Each header on an axis must be unique.*
Bad expression | Runtime error `-2147024809`: *The expression for the user row is invalid and could not be evaluated by the server. Verify that the expression is a valid MDX value expression and try again. The row was not added to the universal report.*

## InsertUserCol

> Example

```vb
Public Sub InsertUserColUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    ' Pass the header cell of the anchor column as a Range
    ' Pass "" for header or expression to use defaults
    oReport.InsertUserCol Range("F51"), 1, "Q2 - Q1", "[plan_time].[plan_time].[Q2-2004] - [plan_time].[plan_time].[Q1-2004]"
End Sub
```

InsertUserCol inserts a custom user-defined column into a Dynamic Universal Report at a position relative to the column that contains the supplied Range. Accepts a header label and an MDX expression to populate the column data cells — pass an empty string for either to use the default (no visible header, cells display as blank). The report must be fully refreshed before calling this method.

<aside class="notice">
InsertUserCol is only supported for Dynamic Universal Reports. Calling this method on a Static Universal Report will raise an error.
</aside>

### Syntax

The following string is the syntax for the InsertUserCol method.

`<UniversalReport>.InsertUserCol <range>, <position>, "<header>", "<expression>"`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
range | An Excel Range whose first cell falls within the column axis of the report. The anchor member is resolved from this cell using the report's header cache. | Range
position | The position at which to insert the column. Accepted values: `1` (Before selection), `2` (After selection), `3` (Start of hierarchy), `4` (End of hierarchy). | Integer
header | The header label for the new column. Pass an empty string to insert the column without a visible header. | String
expression | An MDX value expression evaluated by the TM1 server to populate each data cell in the new column. Any valid MDX value expression is accepted — for example, a member reference such as `[Measures].[Revenue]`, or a calculated expression such as `[Measures].[Revenue] / [Measures].[Units]`. Pass an empty string and the cells will display as blank. | String

### Errors

A runtime error is raised for all invalid inputs. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Range outside report bounds | Runtime error `-2147024809`: *The specified range does not correspond to a header cell within the report. Verify that the cell is inside the row or column axis of the universal report.*
Invalid `position` value | Runtime error `-2147024809`: *Invalid position '\<value\>'. Expected 1 (before), 2 (after), 3 (start of hierarchy), or 4 (end of hierarchy).*
Called on a Static Universal Report | Runtime error `-2146233079`: *Insert user content is only supported for Dynamic Universal Reports.*
Report not fully refreshed | Runtime error `-2146233079`: *The Universal Report must be fully refreshed before inserting user content. Wait for the report to finish loading and try again.*
Duplicate header (same as an existing header on the axis) | Runtime error `-2147024809`: *The header '\<name\>' already exists on the column axis. Each header on an axis must be unique.*
Bad expression | Runtime error `-2147024809`: *The expression for the user column is invalid and could not be evaluated by the server. Verify that the expression is a valid MDX value expression and try again. The column was not added to the universal report.*

## InsertSpacerRow

> Example

```vb
Public Sub InsertSpacerRowUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    ' Pass any cell in the target row
    oReport.InsertSpacerRow Range("E22"), 2
End Sub
```

InsertSpacerRow inserts a blank spacer row into a Static Universal Report at a position relative to the row that contains the supplied Range.

<aside class="notice">
InsertSpacerRow is only supported for Static Universal Reports. Calling this method on a Dynamic Universal Report will raise an error.
</aside>

### Syntax

The following string is the syntax for the InsertSpacerRow method.

`<UniversalReport>.InsertSpacerRow <range>, <position>`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
range | An Excel Range whose row number is used to determine the insertion point. | Range
position | The position relative to the selected row. Accepted values: `1` (Before selection), `2` (After selection). | Integer

### Errors

A runtime error is raised for all invalid inputs. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Invalid range (non-Range value passed) | Runtime error `-2147024809`: *The range argument is not a valid Excel Range. Pass a Range object — for example, a cell reference such as Range("A1").*
Range outside report bounds | Runtime error `-2146233080`: *Index was outside the bounds of the array.*
Invalid `position` value | Runtime error `-2147024809`: *Invalid spacer position '\<value\>'. Expected 1 (before) or 2 (after).*
Called on a Dynamic Universal Report | Runtime error `-2146233079`: *Insert spacer is only supported for Static Universal Reports.*
Report not fully refreshed | Runtime error `-2146233079`: *The Universal Report must be fully refreshed before inserting user content. Wait for the report to finish loading and try again.*
Insert failure (e.g. sheet is protected or column limit reached) | Runtime error `-2146233079`: *Unable to add the spacer to the universal report.*

## InsertSpacerCol

> Example

```vb
Public Sub InsertSpacerColUR()
    Dim oReport As Object
    Set oReport = Reporting.UniversalReports.Get("0", ThisWorkbook.Name, ActiveSheet.Name)
    ' Pass any cell in the target column
    oReport.InsertSpacerCol Range("F51"), 1
End Sub
```

InsertSpacerCol inserts a blank spacer column into a Static Universal Report at a position relative to the column that contains the supplied Range.

<aside class="notice">
InsertSpacerCol is only supported for Static Universal Reports. Calling this method on a Dynamic Universal Report will raise an error.
</aside>

### Syntax

The following string is the syntax for the InsertSpacerCol method.

`<UniversalReport>.InsertSpacerCol <range>, <position>`

### Arguments

Argument | Description | Data type
--------- | ------- | -----------
range | An Excel Range whose column number is used to determine the insertion point. | Range
position | The position relative to the selected column. Accepted values: `1` (Before selection), `2` (After selection). | Integer

### Errors

A runtime error is raised for all invalid inputs. Use an `On Error GoTo` handler to catch errors; call `CognosOfficeAutomationObject.TraceError` inside the handler to log details.

Case | Result
-----|-------
Invalid range (non-Range value passed) | Runtime error `-2147024809`: *The range argument is not a valid Excel Range. Pass a Range object — for example, a cell reference such as Range("A1").*
Range outside report bounds | Runtime error `-2146233080`: *Index was outside the bounds of the array.*
Invalid `position` value | Runtime error `-2147024809`: *Invalid spacer position '\<value\>'. Expected 1 (before) or 2 (after).*
Called on a Dynamic Universal Report | Runtime error `-2146233079`: *Insert spacer is only supported for Static Universal Reports.*
Report not fully refreshed | Runtime error `-2146233079`: *The Universal Report must be fully refreshed before inserting user content. Wait for the report to finish loading and try again.*
Insert failure (e.g. sheet is protected or column limit reached) | Runtime error `-2146233079`: *Unable to add the spacer to the universal report.*
