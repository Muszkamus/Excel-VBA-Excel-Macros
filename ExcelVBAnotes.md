<!-- markdownlint-disable MD025 -->
<!-- markdownlint-disable MD033 -->

# <centre> # **Section2: Your First Macro**

---

- When saving a macro shortcut, it is preferable to have capital letter e.g. ctrl + Shift + "Letter"

- alt + F11 (View VBE)

Code below is recorded via Find and Select option in Excel. Home > Find and Select

```vba
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

```vba

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

# 11. 7 Ways to Run Macros / VBA code (incl. creative & modern buttons)

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

Sub my_Macro() ==== Most used VBA Procedure is thr Sub Procedure.
End Sub ==== This consists of a set of commands the code should execute

Function my_Formula() ==== Function Procedures are commands that create formulas
End Function === The return one value or array.

Application.CutCopyMode = False ==== VBA assigns color to keywords and capitalizes code references

Very useful features >

- Auto Syntax Check (Checks the syntax errors) in options
- Require Variable Decleration (It puts Option Explicit, helps with VBA efficiency) more on section 5 in options
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
5. Macro and VBA code is kept inside Sub Procedures.

---

# <centre> **Section 4: Referencing Ranges, Worksheets & Workbooks with VBA**

---

# 24. Referring to Ranges & Writing to Cells in VBA

---

```vba
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
| `Select`           | Used by the macro recorder to select a cell, but when writing VBA, it is not necessary to select objects. Code is faster without selecting.                                                                        |          |

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
