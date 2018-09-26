Attribute VB_Name = "Module4"
Sub two_weeks_out_report_formatting()
Attribute two_weeks_out_report_formatting.VB_ProcData.VB_Invoke_Func = " \n14"

'--DECLARATIONS------------------------------------
LastRow = Range("A" & Rows.Count).End(xlUp).Row                     'Determines last row count
Dim StartDate As Long, EndDate As Long                              'Date range declarations
    StartDate = DateSerial(Year(Date), Month(Date), Day(Date) + 0)
    EndDate = DateSerial(Year(Date), Month(Date), Day(Date) + 14)

'--------------------------------------------------

Application.ScreenUpdating = False  'turn off screen updating to speed up macro

'Del "Ship To Customer Number" and "Cat ID" columns
    Range("C:C,I:I").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del all columns to the right of "Line Item Status" column
    Range("W:BA").Select
    Selection.Delete Shift:=xlToLeft
'Turn on autofilter "Line Item Status" column to display all BUT "Closed"
    ActiveSheet.Range("A1:W" & LastRow).AutoFilter Field:=22, Criteria1:=Array( _
        "Awaiting Receipt", "Awaiting Shipping", "Booked"), Operator:=xlFilterValues
'Filter for only zeros in the "Reservation Qty" column
    ActiveSheet.Range("A1:W" & LastRow).AutoFilter Field:=18, Criteria1:="0"
'Color "OPD" column BLUE - color details extracted from Record Macro function
    Range("O1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
'Insert column and name col header "Changes Made" and expland col width to 40
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 40
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Changes Made"
'Place thick line down the left side of "OPD" Col
    Range("N1:N" & LastRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -6974059
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -6974059
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'Place thick line down the left side of "Reservation Qty" Col
    Range("R1:R" & LastRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -6974059
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -6974059
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
'Filter "OPD" to display dates from TODAY (using vba DATE) to 14 days in the future
    ActiveSheet.Range("A1:W" & LastRow).AutoFilter Field:=15, _
                                Criteria1:=">=" & StartDate, _
                                Operator:=xlAnd, _
                                Criteria2:="<=" & EndDate
      
Application.ScreenUpdating = True   'turn screen updating back on

End Sub
