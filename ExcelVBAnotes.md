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
