Attribute VB_Name = "Alerts_Holds_Format_Standalone"
Sub Alerts_Holds_Formatting()
Attribute Alerts_Holds_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"

'Turn off screen updating to enable faster running of macro
Application.ScreenUpdating = False

'Select everything, autofit all cols and rows
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
'Add the six (6) columns used for this report
    Columns("O:O").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("O:P").Select
    Range("P1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Margin Holds"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Export Holds"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Manual Holds"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Agile/SWB T&R Mismatch"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Line Holds"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Misc Alerts/Notes"
    Range("O1:S1").Select
    
'Color the first five (5) cols with black fill
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'Color the first five (5) cols with white text
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
'Color last col with yellow fill
    Range("T1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'Adjust column width of several cols to reduce wasted space
    Columns("C:C").ColumnWidth = 8.5
    Columns("E:E").ColumnWidth = 3.2
    Columns("G:G").ColumnWidth = 10.3
    Columns("M:M").ColumnWidth = 8.1
    Columns("O:T").Select
    Selection.ColumnWidth = 15

'Turn screen updating back on
Application.ScreenUpdating = True
End Sub

