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
