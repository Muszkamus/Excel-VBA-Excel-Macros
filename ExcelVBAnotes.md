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
