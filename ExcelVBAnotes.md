<!-- markdownlint-disable MD025 -->
<!-- markdownlint-disable MD029 -->
<!-- markdownlint-disable MD033 -->

# <centre> # **Section2: Your First Macro**

---

- When saving a macro shortcut, it is preferable to have capital letter e.g. ctrl + Shift + "Letter"

- alt + F11 (View VBE)

Code below is recorded via Find and Select option in Excel. Home > Find and Select

```vb
Sub SpecialCells()
'
' SpecialCells Macro
'
' Keyboard Shortcut: Ctrl+Shift+N
'
' Task 1 Write N to empty Cells


    Selection.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "N"


End Sub
Sub HighlightFormulas()

' Task2 1 Highlight formulas

Selection.CurrentRegion.Select
Selection.SpecialCells(xlCellTypeFormulas, 23).Select
ActiveWindow.SmallScroll Down:=-6
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


End Sub

```

---

```vb

Sub AbsoluteMacro()
'
' AbsoluteMacro Macro
'

'
    Range("A3").Select
    Selection.End(xlDown).Select
    Range("A9").Select
End Sub

Sub RelativeMacro()
'
' RelativeMacro Macro
'
    Range("A3").Select ' like absolute macro
   ' ActiveCell.Offset(-5, 0).Range("A1").Select // This will appear on relative macro.
   ' It is good practice to copy the start like absolute macro
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

```

# 11. 7 Ways to Run Macros / vb code (incl. creative & modern buttons)

- alt + F8 = Opens Macro window
- Click macro in Dev tab
- In the view tab, Macros > Run Macro
- From click-access toolbar (Macros Must be turned on in the quick access)
- Ribbon (Right click on the ribbon)
- Insert a shape (Illustrations > Shapes), to add hover text > Put hyperlink behind the button/image (Put any letter in address)
- Normal button (Not ActiveX Button)

---

# <centre> **Section 3: The Object Model**

---

## VB Guidelines & Color procedures

Sub my_Macro() ==== Most used vb Procedure is thr Sub Procedure.
End Sub ==== This consists of a set of commands the code should execute

Function my_Formula() ==== Function Procedures are commands that create formulas
End Function === The return one value or array.

Application.CutCopyMode = False ==== vb assigns color to keywords and capitalizes code references

Very useful features >

- Auto Syntax Check (Checks the syntax errors) in options
- Require Variable Decleration (It puts Option Explicit, helps with vb efficiency) more on section 5 in options
- Auto List Members always on (Code Snippets) in options

- Ctrl + Space (Enables code snippets in specific line)
- F5 to run the project
- F8 to step into code
- F9 Toggle breakpoint
- Ctrl + Shift + F9 Clear all breakpoints

---

# 21. **How to Find the Object, Property & Method**

---

- Record the macro
- Use the Object Library (F2)
- F1 to the Microsoft website
- IntelliSense (Code Snippets)
- Ctrl + Space
- Use the Immediate Window
  In order to work in immediate window (Like Quokka.js), put ? before the code line
  in order to run it, remove question mark

---

# 22. **Summary**

---

1. You refer to an object through its position in the object hierarchy. This dot is used as a separator. If you do not specify the parent, Excel assumes it's the active object.
2. You don't need to select object to manipulate them.
3. Objects have specific porperties & methods.
4. Properties can return a reference to another object.
5. Macro and vb code is kept inside Sub Procedures.

---

# <centre> **Section 4: Referencing Ranges, Worksheets & Workbooks with vb**

---

# 24. Referring to Ranges & Writing to Cells in vb

---

```vb
Option Explicit

Sub ReferToCells()
Cells.Clear
Range("A1").Value = "1st" 'Cells(1, 1) = "1st"
Range("A2:C2").Value = "Second"
Range("A3:C3,E3:F3").Value = "Third"
Range("A4,C4") = "4th"
Range("A5", "C5") = "5th"
Range("A" & 6, "C" & 6) = "6th"
' Range(Cells(6, 1), Cells(6, 3)).Value = "6th"
Range("A4:C7").Cells(4, 2).Value = "7th"
Range("A1").Offset(7, 2).Value = "8th"
Range("A1:B1").Offset(8, 1).Value = "9th"
Range("LastOne").Value = "10th, LastOne"

Rows("12:14").RowHeight = 30
Range("16:16, 18:18,20:20").RowHeight = 30
Range("H:H,J:J").ColumnWidth = 10
Cells.Columns.AutoFit

End Sub


```

For work and simplicity, these are best:

**Range("A1").Value = "Hello"** – direct and intuitive

**Cells(1, 1)** – great for row-column dynamic loops. Also refers to cell A1, but using row and column numbers.

**Range("A1").Offset(2, 3).Value = "Moved 2 down, 3 right"**– perfect for flexible positioning. Moves from a known reference point by a number of rows and columns.

**Range("LastOne").Value = "Final value"**– clean and scalable with named ranges. Refers to a named range in your Excel sheet.

**Range("A" & i, "C" & i)** – easy loop integration. Creates a horizontal range from column A to C on a given row (e.g., A3:C3 if i = 3).

---

# 25. **Most Useful Range Properties & Methods**

---

| **Code Execution**           | **Description**                                                                                                                   | **Type**     |
| ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------- | ------------ |
| `Value`                      | Shows the underlying value in a cell. This is the default property of the range object.                                           | Read / Write |
| `Cells`                      | Returns a cell or range of cells within a range object.                                                                           | Read / Write |
| `End`                        | Returns the last cell of the range. Similar to Ctrl + ↓ or ↑ or → or ←                                                            | Read-only    |
| `Offset`                     | Returns a new range based on offset row & column arguments.                                                                       | Read / Write |
| `Count`                      | Returns the number of cells in a range.                                                                                           | Read-only    |
| `Column` / `Row`             | Returns the column or row number of a range. If you select more than one cell, returns the first occurrence in the range.         | Read-only    |
| `CurrentRegion`              | Used with other properties such as `.Address`; returns the range of data.                                                         | Read-only    |
| `EntireColumn` / `EntireRow` | Returns a range object that represents the entire row or column.                                                                  | Read-only    |
| `Resize`                     | Changes the size of the range by defining the rows & columns for resizing.                                                        | Read / Write |
| `Address`                    | Shows the range address including the `$` signs.                                                                                  | Read-only    |
| `Font`                       | Returns a font object that has other properties (e.g., bold).                                                                     | Read / Write |
| `Interior`                   | Used with other properties such as `.Color` to set colors.                                                                        | Read / Write |
| `Formula`                    | Places a formula in a cell. Use English syntax for compatibility. Use `FormulaLocal` if using a different Excel language version. | Read / Write |
| `NumberFormat`               | Defines number format (uses English version).                                                                                     | Read / Write |
| `Text`                       | Returns the data as a string & includes formatting.                                                                               | Read-only    |
| `HasFormula`                 | Returns `True`, `False`, or `Null` if the range has a mix.                                                                        | Read-only    |

| **Code Execution** | **Description**                                                                                                                                                                                                    | **Type** |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------- |
| `Copy`             | This is a practical method because it has paste destination as its argument. This means you just need one line of code.                                                                                            |          |
| `PasteSpecial`     | Allows usage of Excel’s Paste Special options. To use more than one option, repeat the line of code with the new option.                                                                                           |          |
| `Clear`            | Deletes contents and cell formatting in a specified range.                                                                                                                                                         |          |
| `Delete`           | Deletes the cells and shifts the cell around the area to fill up the deleted range. The delete method uses an argument to define how to shift the cells. Select `xlToLeft` or `xlUp`.                              |          |
| `SpecialCells`     | Returns a range that matches the specified cell types. This method has 2 arguments. `xlCellType` is required (e.g. cells with formulas or comments) and an optional argument defines more detail (e.g. constants). |          |
| `Sort`             | Sorts a range of values.                                                                                                                                                                                           |          |
| `PrintOut`         | Also a method of the worksheet object.                                                                                                                                                                             |          |
| `Select`           | Used by the macro recorder to select a cell, but when writing vb, it is not necessary to select objects. Code is faster without selecting.                                                                         |          |

---

# 26. **4 Methods to Find the Last Row of your Range**

---

1- Use the End Property of the Range Object

- Range("K6").Value = Cells(Rows.Count,1).End(xlUp.Row)
- Range("K6").Value = Range("A4").End(xlDown).Row
- Range("K8").Value = Cells(4,Columns.Count).End(xlToLeft).Column

2- Use the CurrentRegion Property of the Range Object

- Range("K10").Value = Range("A4").CurrentRegion.Rows.Count

3- Use the SpecialCells Method of the Range Object

- Range("K11").Value = Cells.SpecialCells(xlTypeLastCell).Row
  4- Use the UsedRange Proprty of the Worksheet Object

---

# 27. **Copying & resizing a variably sized range**

---

1. Copy Method

```vb
Range("A4:E10").Copy.Range("J4")
' Copies a fixed-size range (A4:E10) and pastes it starting at J4.
Range("A4").CurrentRegion.Copy.Range("J4")
' Dynamically copies the entire block of contiguous data starting from A4, including all directions until empty cells are hit.
```

2. Paste Special

```vb
Range("A4").CurrentRegion.Copy
'First line copies the dynamic range.
Range("J20").PasteSpecial xlPasteValuesAndNumberFormats
'Second line pastes only values and formatting (no formulas or links).
Range("J20").PasteSpecial xlPasteComments
'Third line pastes just comments.
```

3. Resize Property with Copy Method

```vb
Range("A4").CurrentRegion
```

```vb
Application.CutCopyMode = False 'Cancels the copy "marching ants" and clipboard state.
```

---

# 28. **Properly Referencing Worksheets**

---

```vb

ActiveSheet ' Refers to the sheet where the macro is currently running
Worksheets(6).Select ' Refers to the 6th worksheet in the workbook
Sheets(6).Select ' Same as above but includes chart sheets too
Sheet6.Range("A3").Value = "" ' Refers to a sheet by its code name
```

Better examples >

```vb
Worksheets("SalesData").Range("A1").Value = "Loaded"
'By Name (less safe than code name, but readable)
ThisWorkbook.Worksheets("Summary").Range("A1").Value = "Updated"
'ThisWorkbook ensures it works on the workbook containing the code, not just any active workbook.
SalesSheet.Range("B2").Value = "Final"
'Where SalesSheet is the code name you set in the vb editor (left pane, not sheet tab).
```

---

# 29. **Properly Referencing Workbooks**

---

```vb

ActiveWorkbook ' Refers to the workbook that is currently active (on top).
' ⚠️ Use with caution — it can change if the user clicks another workbook.

' --------------------------

' Referring by index number (e.g., the 1st workbook opened in this session)
' ⚠️ Not recommended — fragile and not readable
Workbooks(1) ' Refers to the first opened workbook

' --------------------------

' Referring to a workbook by its name (must be open!)
Workbooks("Deskbook.xlsx").Sheets(1).Range("A3").Value = "I will copy data here"
' ✅ Safer method — clearly identifies which workbook and sheet to use
' ⚠️ The workbook name must match exactly, including extension (e.g., .xlsm, .xlsx)

' --------------------------

ThisWorkbook ' Refers to the workbook **where this vb code is written**
' ✅ Very reliable — doesn't change even if another workbook is active
' Best used when your macro always runs from a specific workbook

' --------------------------

' Opening another workbook from a file path
Application.Workbooks.Open("C:\Users\YourName\Documents\Data.xlsx")
' ✅ Use this to load external files

' --------------------------

' Closing the currently active workbook and saving changes
ActiveWorkbook.Close SaveChanges:=True
' ⚠️ Only use if you're sure which workbook is active — safer to reference by name or object
```

- Always use ThisWorkbook if your macro is tied to your own workbook (like a tool or template).

- Avoid using ActiveWorkbook unless you're handling user-driven tasks (like dragging files in).

- For automation, it's best to assign opened workbooks to variables:

---

# <centre> Section 5: Working with Variables

---

# 35. **Declaring Variables, Arrays & Constants (Role of Option Explicit)**

---

```vb
Option Explicit  ' Forces explicit declaration of all variables — same idea as "use strict" in JavaScript
                 ' Prevents bugs from typos or undeclared variables

Public Sub DefiningVariables()

    ' Declare two Long integers to hold row numbers (like let lastRow, firstRow in JS)
    Dim lastRow As Long, FirstRow As Long

    ' Assign the total number of rows in the worksheet to lastRow (1,048,576 in Excel 365)
    Let lastRow = Rows.Count
    Debug.Print lastRow  ' Print to the Immediate Window (like console.log)

    ' Declare a fixed-size array of 12 elements to hold month names
    Dim MyMonth(1 To 12) As String
    MyMonth(1) = "Jan"
    MyMonth(2) = "Feb"
    MyMonth(12) = "Dec"
    ' This is similar to: const myMonth = []; myMonth[0] = "Jan";

    ' Declare a 2D array (12 rows × 3 columns), type Variant allows mixed data types
    Dim MonthSales(1 To 12, 1 To 3) As Variant
    ' Similar to: let monthSales = Array(12).fill().map(() => Array(3));

    ' Declare a constant — its value cannot be changed later
    Const myScenario As String = "Actual"
    ' Like: const myScenario = "Actual";

End Sub
```

---

# 36. **Using Object Variables (Set statement)**

---

Variables can also hold objects. Common objects are:

```vb
Dim NewBook as WorkBook 'Workbook Object
Dim NewSheet As WorkSheet 'Worksheet Object
Dim NewRange As Range 'Range Object
```

To Assign variables to objects, you need to use the SET statement

```vb
Set NewBook = Workbooks.Add

'Example

Option Explicit  ' Enforces variable declaration to avoid bugs from typos

Public Sub DefiningVariables()

    ' Declare a Workbook object to store the new workbook
    Dim NewBook As Workbook

    ' Declare a Worksheet object to refer to the first sheet in that new workbook
    Dim NewSheet As Worksheet

    ' Create a new workbook and assign it to the NewBook variable
    Set NewBook = Workbooks.Add

    ' Get the first worksheet from the new workbook and assign it to NewSheet
    Set NewSheet = NewBook.Sheets(1)

    ' Write the value "New One" into cell A1 of the new worksheet
    NewSheet.Range("A1").Value = "New One"

    ' (Optional) Rename the worksheet to make it clearer
    NewSheet.Name = "Summary"

    ' (Optional) Autofit column A to match content width
    NewSheet.Columns("A").AutoFit

    ' (Optional) Save the new workbook to a specified path
    NewBook.SaveAs Filename:="C:\Users\YourName\Documents\NewFile.xlsx"

    ' (Optional) Close the workbook after saving
    NewBook.Close SaveChanges:=False

End Sub


```

---

# 37. **Scoping Variables**

---

1. Procedure: Variables exists only when the procedures run. Dim is inside the procedure. Memory is released afted procedure ends.

```vb
Sub Defining_Variables()
    Dim LastRow As Long, FirstRow as Long
    '---code
End Sub
```

2. Module: Variable exists for All Procedures within the module. Dim is outside any procedure. Typically right below **Option explicit**. Value is kept in memory after procedure completes.

```vb
Option Explicit
Dim LastRow As Long, FirstRow as Long

Sub Defining_Variables()
    '---code
End Sub
```

3. All modules & procedures: Variables exists for All Modules and all Procedures. USe the keyword "Public" to declare these variables. Can be declared in any Module before the first procedure. Value is kept in memory after procedure completes.

```vb
Option Explicit
Public LastRow As Long, FirstRow as Long

Sub Defining_Variables()
    '---code
End Sub
```

---

# <centre> Section 6: **Looping through Collections & Making Decisions**

---

# 41. **With & End With for Easier Code Writing**

---

```vb
Option Explicit

Sub With_Change_Font()

Dim MyRange As Range
    Set MyRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    Debug.Print MyRange.Address 'Shows the exact cells we are working with
    With MyRange.Font 'The With ... End With block in vb is used to simplify and optimize code when you're repeatedly accessing properties or methods of the same object.
        .Name = "Arial"
        .Size = 12
        .Bold = True

    End With

End Sub

Sub With_Reset_Font()

Dim MyRange As Range
    Set MyRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    Debug.Print MyRange.Address 'Shows the exact cells we are working with
    With MyRange.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False

    End With

End Sub
```

---

# 42. **For Each to Loop Through Collections (sheets, ranges etc.) in one go**

---

```vb
Sub Protect_All_Sheets()
Dim Sh As Worksheet
For Each Sh In ThisWorkbook.Worksheets
    Sh.Protect '"test"
    Debug.Print Sh.Name

Next Sh
End Sub
```

---

# 43. **IF Then (Else, ElseIF) for Conditional Outcomes**

---

```vb
If Range("B3").Value <> "" Then 'This means not equal to, same as ! in JS, otherwise it is =
    Range("C3").Value = Range("B3").Value

End If
```

```vb
Sub Protect_Special_Sheets()

    Dim Sh As Worksheet

    For Each Sh In ThisWorkbook.Worksheets

        If Sh.Name = "Purpose" Then
            ' Protect "Purpose" sheet — don't allow formatting (default protection)
            Sh.Protect

        ElseIf Sh.CodeName = "shWith" Then
            ' Protect sheet using CodeName
            Sh.Protect

        Else
            ' Other sheets – nothing happens
            ' Add other code here if needed
        End If

    Next Sh

End Sub


```

# 44. **Select Case as Alternative for Many IF Statements**

```vb
Sub SimpleCase()

    Select Case Range("B3").Value
        Case 1 To 200
            Range("C3").Value = "Good"
        Case 0
            Range("C3").Value = ""
        Case Is > 200
            Range("C3").Value = "Excellent"
        Case Else
            Range("C3").Value = "Bad"

    End Select

End Sub
```

---

# 45. **Goto Statement to Change Program Flow**

---

- Mainly for error handling
- Execute a different part of code depending on a condition

```vb
Sub Simple_GoTo()

    Range("D3").Value = "" ' Clear previous message

    If IsError(Range("B3").Value) Then GoTo GetOut

    ' No error: copy value from B3 to C3
    Range("C3").Value = Range("B3").Value
    Exit Sub

GetOut:
    Range("D3").Value = "You have an error in the cell"

End Sub
```

---

# 48. **Bonus: Unhide All Sheets in One Go**

```vb
Sub Unhide_All()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
    sh.Visible = True
    Next sh
End Sub

```

We can also use Personal Macro Workbook to apply it everywhere, even for non macro worksheets. But first, do dummy record as personal one, and do this >

```vb

Sub Unhide_All()
    Dim sh As Worksheet
    For Each sh In Worksheets
    sh.Visible = True
    Next sh
End Sub
```

---

# <centre> Section 7: **Useful Built-in Functions**

---

# 50.**Overview: vb versus Worksheet Functions**

---

```vb
Option Explicit

Sub vb_Excel_Functions()
    With Sheet1

    .Range("B3").Value = vb.DateTime.Date 'Get todays date (Date works too)
    .Range("B6").Value = vb.UCase(.Range("A6").Value) ' B6 will be A6 but uppercase
    .Range("B7").Value = vb.LCase(.Range("A7").Value)
    .Range("B8").Value = vb.StrConv(.Range("A8"), vbProperCase)
    .Range("B9").Value = Excel.Application.WorksheetFunction.Proper(.Range("A9").Value)


    Dim myRange As Range 'define variable
    Set myRange = .Range("A13").CurrentRegion 'Give variable the amount

    .Range("B11").Value = Excel.WorksheetFunction.Max(myRange) 'Find max in the variable

    End With

End Sub

```

---

# 51. **Most Useful vb Functions**

---

```vb
Sub vb_Functions()

Range("B3").Value = vb.Month(vb.Date) 'Numeric
Range("B4").Value = vb.MonthName(vb.Month(vb.Date)) 'String
Range("B9").Value = vb.IsEmpty(Range("A9")) 'If there is a value in B9, print the outcome in A9

End Sub

```

---

# 52. **Message Box (also with Yes, No buttons)**

---

```vb
Sub Simple_MsgBox()

vb.Interaction.MsgBox "Hello"

End Sub

```

```vb
Sub Simple_MsgBox()


MsgBox prompt:="Hello " & Excel.Application.UserName & "." & vb.Constants.vbNewLine & "Thanks for stopping by.", Title:="Welcome!"

End Sub

```

```vb
Sub Yes_No_MsgBox()
Dim Answer As VbMsgBoxResult
Answer = MsgBox("Are you sure?", vbYesNo + vbQuestion + vbDefaultButton2, "Clear Cells") 'Default button is the one chosen straight away, in this instance button no

If Answer = vbYes Then
    Range("A7:B9").Clear
Else
        Exit Sub
End If
End Sub
```

---

# 53. **Input Box (vb InputBox Function)**

---

```vb
Sub vb_Simple_InputBox()

    Dim myInp   As String
    myInp = vb.InputBox("Please input subtitle", "subtitle please")
    If myInp = "" Then Exit Sub
    Range("A2").Value = Excel.WorksheetFunction.Proper(myInp)

End Sub
```

---

# <centre> Section 8: **Debugging, Error Handling & Procedure Scope**

---

# 60. **Error Handling: Different Methods**

---

- 🔹 1. Basic Input Validation (If Statement)

```vb
  If WorksheetFunction.Count(Range("A1:A10")) >= 3 Then
  ' Safe to proceed
  End If
```

🔹 2. On Error GoTo [Label]
Redirects to error handling code when an error occurs.

```vb
On Error GoTo ErrorHandler

' Your main code here

Exit Sub

ErrorHandler:
MsgBox "An error has occurred."
```

🔹 3. On Error Resume Next

```vb

On Error Resume Next
Range("A1").Value = 1 / 0 ' This will error, but be ignored
On Error GoTo 0 ' Resume normal error handling
```

🔹 4. On Error GoTo Label + Error Number Check

```vb
On Error GoTo ErrorHandler

' Your code here that might trigger error 424

Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 424 ' Object required (e.g., user clicked Cancel)
        Exit Sub
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
End Select
```

🔹 5. Prevent Error Message on Successful Completion

```vb
On Error GoTo ErrorHandler

' Code that may error

Exit Sub ' Skip the error handler if everything was OK

ErrorHandler:
MsgBox "Something went wrong."
```

---

# 61. **Faster Code by Suppressing Pop-ups & Flickering Screen**

---

```vb
Sub Slower_code()
'for tab Faster
    Dim ShNew As Worksheet
    Dim cell As Range
    Application.StatusBar = "Wait"
    Set ShNew = Worksheets.Add
    For Each cell In ShNew.Range("A1:A100000")
        cell.Value = 10
    Next cell

    ShNew.Delete
    Sheet8.Select
    Application.StatusBar = ""

End Sub
```

```vb
Sub Faster_Code()
'for tab Faster
    Dim ShNew As Worksheet
    Dim cell As Range

    'turn off some features
    With Application
        .StatusBar = "Short wait"
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With

    Set ShNew = Worksheets.Add

    For Each cell In ShNew.Range("A1:A100000")
        cell.Value = 10
    Next cell

    ShNew.Delete
    Sheet8.Select

    'turn features on again
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .StatusBar = ""
        'in case you have used the copy and pastespecial methods, you could have a lot of data on the clipboard.
        .Application.CutCopyMode = False
    End With

End Sub
```

---

# 62. **Procedure Scope & Running one Sub from Another**

VBA sub procedures are public by default and can be called from other subs using Call SubName or just SubName. Use Private to limit access within the same module and hide subs from the macro list. Breaking code into smaller subs (e.g., Entry_Point, Exit_Point) improves readability, reusability, and helps manage repeated setup/cleanup tasks. It’s cleaner, safer, and easier to maintain.
