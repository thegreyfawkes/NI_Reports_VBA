Attribute VB_Name = "Bklg_ERNI_Processing"
Sub ERNI_Backlog_Processing()


'--PART 1:-----------------------------------------
'Delete (a) first TWO rows, and hide unneeded COLS
'--------------------------------------------------

'Delete first 2 rows
Rows(1).EntireRow.Delete
Rows(1).EntireRow.Delete

'--DECLARATIONS-----------------------------------
Lastrow = Range("A" & Rows.Count).End(xlUp).Row
'-------------------------------------------------

'Hide Cols A:E, H:J, L, N, R:Y, and AB
Columns("A:E").EntireColumn.Hidden = True
Columns("H:J").EntireColumn.Hidden = True
Columns("L:L").EntireColumn.Hidden = True
Columns("N:N").EntireColumn.Hidden = True
Columns("R:Y").EntireColumn.Hidden = True
Columns("AB:AB").EntireColumn.Hidden = True

'--PART 2:-------------------------------------------------
'Insert the five (5) columns used for the ERNI report, and color the needed cells
'----------------------------------------------------------

'Insert five (5) Columns to the left of Column R (18, "New Dock Date")
    Columns("AC:AG").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormatFromLeftOrAbove 'or xlFormatFromRightOrBelow

'Name the five (5) column headers, and color all seven of the total used headers as dark green (RGB(118, 147, 60))
Range("AC1") = "MFR PO Qty"
Range("AD1") = "MFR PO"
Range("AE1") = "MFR PO Line #"
Range("AF1") = "MFR PO Need By Date"
Range("AG1") = "MFR PO Promise Date"
Range("AC1:AG1").Interior.Color = RGB(118, 147, 60)
Range("M1").Interior.Color = RGB(118, 147, 60)
Range("O1").Interior.Color = RGB(118, 147, 60)

'Color the cell ranges used in the colored headers as light green (RGB(235, 241, 222))
Range("AC2:AG" & Lastrow).Interior.Color = RGB(235, 241, 222)
Range("M2:M" & Lastrow).Interior.Color = RGB(235, 241, 222)
Range("O2:O" & Lastrow).Interior.Color = RGB(235, 241, 222)



End Sub


