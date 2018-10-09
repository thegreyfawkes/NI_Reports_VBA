Attribute VB_Name = "Bklg_Add_CONCAT_COL_Standalone"
Sub Add_CONCAT_to_Bklg()
Attribute Add_CONCAT_to_Bklg.VB_ProcData.VB_Invoke_Func = " \n14"

'--DECLARATIONS------------------------------------
Lastrow = Range("A" & Rows.Count).End(xlUp).Row                     'Determines last row count
'--------------------------------------------------
' Add_CONCAT_to_Bklg Macro

'
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "CONCAT"
    Range("B2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[2]&""-""&RC[3]"
    Range("B2").Select
    Selection.Copy
    'Range("B2:B2400").Select
    ActiveSheet.Range("B2:B" & Lastrow).Select
    ActiveSheet.Paste
    'Columns("B:B").Select
    ActiveSheet.Range("B2:B" & Lastrow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace what:=" ", Replacement:="", lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
