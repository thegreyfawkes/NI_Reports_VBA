Attribute VB_Name = "Module6"
Sub SureShip_Processing()

'--DECLARATIONS------------------------------------
LastRow = Range("A" & Rows.Count).End(xlUp).Row         'Determines last row count
'''
Dim YesOrNo As Integer                                  'Msg box to proceed or cancel the macro
'''
'Dim DbExtract, DuplicateRecords As Worksheet
'Set DbExtract = ThisWorkbook.Sheets("Sheet1")
'Set DuplicateRecords = ThisWorkbook.Sheets("SureShip")
'--------------------------------------------------

YesOrNo = MsgBox("Use SAVE AS macro first, as this will delete columns needed for the backlog! Continue?", vbOKCancel)
If YesOrNo = 2 Then Exit Sub

Application.ScreenUpdating = False  'turn off screen updating to speed up macro

'Delete first 2 rows of title and empty space (data starts on Row 3)
Rows(1).EntireRow.Delete
Rows(1).EntireRow.Delete


'Del Cols BR (70) thru DA (105)
    Range("BR:DA").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AY (51) thru BP (68)
    Range("AY:BP").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AM (39) thru AW (49)
    Range("AM:AW").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AH (34) thru AK (37)
    Range("AH:AK").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AB (28) thru AD (30)
    Range("AB:AD").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols P (16) thru Y (25)
    Range("P:Y").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols I (9) and J (10)
    Range("I:J").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Col G (7)
    Range("G:G").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols A (1) thru E (5)
    Range("A:E").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

'Filter for Warehouse Status of "Ready to Release"
    ActiveSheet.Range("A1:O" & LastRow).AutoFilter Field:=13, Criteria1:=Array( _
        "Released to Warehouse", "Staged/Pick Confirmed"), Operator:=xlFilterValues

'Filter for Warehouse Status of "Ready to Release"
    ActiveSheet.Range("A1:O" & LastRow).AutoFilter Field:=14, Criteria1:="*Air*"

'Copy visible cells to new sheet "SureShip"
    ActiveSheet.Range("A1:O" & LastRow).Select
    Selection.Copy
    'Sheets.Add After:=ActiveSheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SureShip"
    Rows("1:1").Select
    ActiveSheet.Paste

'Delete "Sheet1"
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.Delete
    
Application.ScreenUpdating = True   'turn screen updating back on
    
End Sub

Sub AllInOne_Morning_Reports()

'--DECLARATIONS------------------------------------
LastRow = Range("A" & Rows.Count).End(xlUp).Row         'Determines last row count
'''
Dim YesOrNo As Integer                                  'Msg box to proceed or cancel the macro
'''
'Dim DbExtract, DuplicateRecords As Worksheet
'Set DbExtract = ThisWorkbook.Sheets("Sheet1")
'Set DuplicateRecords = ThisWorkbook.Sheets("SureShip")
'--------------------------------------------------

YesOrNo = MsgBox("This will make permanent changes to this workbook that cannot be undone. Continue?", vbOKCancel)
If YesOrNo = 2 Then Exit Sub

Application.ScreenUpdating = False  'turn off screen updating to speed up macro

'_______________________________________
'      SURESHIP REPORT
'_______________________________________

'Copy ALL cells from Sheet1 to new sheet "SureShip"
    ActiveSheet.Range("A1:DA" & LastRow).Select
    Selection.Copy
    'Sheets.Add After:=ActiveSheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SureShip"
    Rows("1:1").Select
    ActiveSheet.Paste


'Delete first 2 rows of title and empty space (data starts on Row 3)
Rows(1).EntireRow.Delete
Rows(1).EntireRow.Delete


'Del Cols BR (70) thru DA (105)
    Range("BR:DA").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AY (51) thru BP (68)
    Range("AY:BP").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AM (39) thru AW (49)
    Range("AM:AW").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AH (34) thru AK (37)
    Range("AH:AK").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols AB (28) thru AD (30)
    Range("AB:AD").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols P (16) thru Y (25)
    Range("P:Y").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols I (9) and J (10)
    Range("I:J").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Col G (7)
    Range("G:G").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
'Del Cols A (1) thru E (5)
    Range("A:E").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

'Filter for Warehouse Status of "Ready to Release"
    ActiveSheet.Range("A1:O" & LastRow).AutoFilter Field:=13, Criteria1:=Array( _
        "Released to Warehouse", "Staged/Pick Confirmed*"), Operator:=xlFilterValues

'Filter for Warehouse Status of "Ready to Release"
    ActiveSheet.Range("A1:O" & LastRow).AutoFilter Field:=14, Criteria1:="*Air*"

'_______________________________________
'      INTERNAL BACKLOG
'_______________________________________
'Back to Sheet1, Copy all cells, and paste to Backlog INT
    Sheets("Sheet1").Select
    ActiveSheet.Range("A1:DA" & LastRow).Select
    Selection.Copy
    'Sheets.Add After:=ActiveSheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Backlog_INT"
    Rows("1:1").Select
    ActiveSheet.Paste
    
    'Delete first 2 rows
    Rows(1).EntireRow.Delete
    Rows(1).EntireRow.Delete

'Delete cols in reverse numerical order
Columns(97).EntireColumn.Delete 'S&D Auth No
Columns(95).EntireColumn.Delete 'SPH Code [new COL on 27-Aug-2018 report]
Columns(94).EntireColumn.Delete 'Special Handling
Columns(93).EntireColumn.Delete 'Line Item Add after Order Entry
Columns(92).EntireColumn.Delete 'Line Item Creation Date
Columns(89).EntireColumn.Delete 'Govt Contract
Columns(88).EntireColumn.Delete 'Govt Rating
Columns(85).EntireColumn.Delete 'Sold To Contract Name
Columns(84).EntireColumn.Delete 'Purchasing SPQ
Columns(83).EntireColumn.Delete 'BSA Line Num
Columns(82).EntireColumn.Delete 'Var Data 12
Columns(81).EntireColumn.Delete 'Var data 11
Columns(80).EntireColumn.Delete 'var data 10
Columns(79).EntireColumn.Delete 'var data 9
Columns(78).EntireColumn.Delete 'var data 8
Columns(77).EntireColumn.Delete 'var data 7
Columns(76).EntireColumn.Delete 'var data 6
Columns(75).EntireColumn.Delete 'var data 5
Columns(74).EntireColumn.Delete 'var data 4
Columns(73).EntireColumn.Delete 'var data 3
Columns(72).EntireColumn.Delete 'var data 2
Columns(67).EntireColumn.Delete 'Price List
Columns(66).EntireColumn.Delete 'SQR Number
Columns(65).EntireColumn.Delete 'AUN
Columns(63).EntireColumn.Delete 'Order Source
Columns(62).EntireColumn.Delete 'DW Eligible
Columns(61).EntireColumn.Delete 'Ship To Customer Early Days
Columns(60).EntireColumn.Delete 'Bill To Customer Early Days
Columns(57).EntireColumn.Delete 'Item Status
Columns(56).EntireColumn.Delete 'End Customer Name
Columns(55).EntireColumn.Delete 'Override ATP
Columns(54).EntireColumn.Delete 'Line Category
Columns(53).EntireColumn.Delete 'Line Type
Columns(46).EntireColumn.Delete 'Owned BY
Columns(45).EntireColumn.Delete 'End Customer ISR Name
Columns(44).EntireColumn.Delete 'End Customer FSR Name
Columns(43).EntireColumn.Delete 'FSR Name
Columns(42).EntireColumn.Delete 'ISR Name
Columns(40).EntireColumn.Delete 'Exchange Rate
Columns(39).EntireColumn.Delete 'Buy Currency Code
Columns(38).EntireColumn.Delete 'Warehouse Status
Columns(35).EntireColumn.Delete 'TOH Qty
Columns(29).EntireColumn.Delete 'Order Date
Columns(25).EntireColumn.Delete 'HiREL
Columns(24).EntireColumn.Delete 'BPB Code
Columns(23).EntireColumn.Delete 'Margim Amt
Columns(22).EntireColumn.Delete 'GM%
Columns(21).EntireColumn.Delete 'Extended Resale USD (A)
Columns(19).EntireColumn.Delete 'Cost
Columns(10).EntireColumn.Delete 'Buyer Name
Columns(9).EntireColumn.Delete 'Bill To No
Columns(7).EntireColumn.Delete 'Account Number
Columns(5).EntireColumn.Delete 'EB Name
Columns(4).EntireColumn.Delete 'Branch
Columns(3).EntireColumn.Delete 'Region/Program
Columns(2).EntireColumn.Delete 'Location Name
Columns(1).EntireColumn.Delete 'OU Name

'-----------------------------------------------------
'PART 2: Color the headers
'-----------------------------------------------------

'Color INTERNAL ONLY header cells ORANGE
Range("B1").Interior.Color = RGB(255, 192, 0)       'Ship To Customer Number
Range("H1").Interior.Color = RGB(255, 192, 0)       'Cat ID
Range("Q1:R1").Interior.Color = RGB(255, 192, 0)    'SSD and SAD
Range("AB1").Interior.Color = RGB(255, 192, 0)      'Order Type
Range("AG1").Interior.Color = RGB(255, 192, 0)      'Delivery ID
Range("AJ1:AL1").Interior.Color = RGB(255, 192, 0)  'Sales Order Line Alert Active Cnt, Sales Order Line Hold Active Cnt, Source of Supply
Range("AO1").Interior.Color = RGB(255, 192, 0)      'Internal Comments

'Color the "Line Item Status" column GREEN
Range("W1").Interior.Color = RGB(146, 208, 80)      'Line Item Status

'Color all other EXTERNAL header cells GREY
Range("A1").Interior.Color = RGB(213, 217, 226)
Range("C1:G1").Interior.Color = RGB(213, 217, 226)
Range("I1:P1").Interior.Color = RGB(213, 217, 226)
Range("S1:V1").Interior.Color = RGB(213, 217, 226)
Range("X1:AA1").Interior.Color = RGB(213, 217, 226)
Range("AC1:AF1").Interior.Color = RGB(213, 217, 226)
Range("AH1:AI1").Interior.Color = RGB(213, 217, 226)
Range("AM1:AN1").Interior.Color = RGB(213, 217, 226)

'-----------------------------------------------------
'PART 3: Autofilter VAS line items, delete rows, then unfilter
'-----------------------------------------------------
ActiveSheet.Range("M1").AutoFilter Field:=13, Criteria1:="*.*.*"
Range("M2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.EntireRow.Delete
ActiveSheet.AutoFilterMode = False

'Resize the rows and cols
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("AH:AH").Activate
    Selection.ColumnWidth = 160.71
    Range("AO:AO").Activate
    Selection.ColumnWidth = 160.71
    Cells.Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

'_______________________________________
'      EXTERNAL BACKLOG
'_______________________________________

'Copy INT Backlog to new sheet "Backlog_EXT"
    ActiveSheet.Range("A1:DA" & LastRow).Select
    Selection.Copy
    'Sheets.Add After:=ActiveSheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Backlog_EXT"
    Rows("1:1").Select
    ActiveSheet.Paste
    
' Delete all internal Cols
    Range("B:B,H:H,Q:Q,R:R").Select
    Range("R1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("B:B,H:H,Q:Q,R:R,AB:AB,AG:AG").Select
    Range("AG1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("B:B,H:H,Q:Q,R:R,AB:AB,AG:AG,AJ:AJ,AK:AK,AL:AL,AO:AO").Select
    Range("AO1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.LargeScroll ToRight:=-1
    
'Resize the rows and cols
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("AB:AB").Activate
    Selection.ColumnWidth = 160.71
    Cells.Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

'_______________________________________
'      OTX Report
'_______________________________________
'Back to Sheet1, Copy all cells, and paste to "OTX" sheet
    Sheets("Backlog_INT").Select
    ActiveSheet.Range("A1:DA" & LastRow).Select
    Selection.Copy
    'Sheets.Add After:=ActiveSheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "OTX"
    Rows("1:1").Select
    ActiveSheet.Paste

'-----------------------------------------------------
'PART 1: Delete columns not needed for OTX Report
'-----------------------------------------------------
'Delete cols in reverse numerical order
Columns(41).EntireColumn.Delete 'Internal Comments
Columns(40).EntireColumn.Delete 'Ship Method
Columns(39).EntireColumn.Delete 'Supplier Allocation
Columns(38).EntireColumn.Delete 'Source of Supply
Columns(37).EntireColumn.Delete 'Sales Order...Hold
Columns(36).EntireColumn.Delete 'Sales Order...Alert
Columns(33).EntireColumn.Delete 'Delivery ID
Columns(30).EntireColumn.Delete 'Arrow Lead Time
Columns(29).EntireColumn.Delete 'Supplier Lead Time
Columns(28).EntireColumn.Delete 'Order Type
Columns(23).EntireColumn.Delete 'Line Item Status
Columns(22).EntireColumn.Delete 'Order Status
Columns(20).EntireColumn.Delete 'FOH Qty
Columns(19).EntireColumn.Delete 'Reservation Qty
Columns(18).EntireColumn.Delete 'SAD
Columns(17).EntireColumn.Delete 'SSD
Columns(15).EntireColumn.Delete 'Request date
Columns(11).EntireColumn.Delete 'Original Qty
Columns(10).EntireColumn.Delete 'Resale
Columns(8).EntireColumn.Delete 'Cat ID
Columns(2).EntireColumn.Delete 'Ship To Customer

'-----------------------------------------------------
'PART 2: Filter & delete rows that have a "0" or are BLANK in the TRACKING NUMBER col (Col P, or #16)
'-----------------------------------------------------
'old range P1
ActiveSheet.Range("P1").AutoFilter Field:=16, Criteria1:="=", Operator:=xlOr, _
Criteria2:="0"
''Select the first cell of data in the first column, then the rest of the filtered data
Range("P2").Select
Range(Selection, Selection.End(xlDown)).Select
''Delete the filtered data
Selection.EntireRow.Delete
''Remove the autofilter
ActiveSheet.AutoFilterMode = False

'-----------------------------------------------------
'PART 3: Filter & delete rows that are BLANK in the FREIGHT FORD CODE col (Col T, or #20)
'-----------------------------------------------------
'old range R1
ActiveSheet.Range("R1").AutoFilter Field:=18, Criteria1:="="
''Select the first cell of data in the first column, then the rest of the filtered data
Range("R2").Select
Range(Selection, Selection.End(xlDown)).Select
''Delete the filtered data
Selection.EntireRow.Delete
''Remove the autofilter
ActiveSheet.AutoFilterMode = False

'Resize the rows and cols
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("O:O").Activate
    Selection.ColumnWidth = 30
    Range("P:P").Activate
    Selection.ColumnWidth = 80

'_______old stuff at the end of everything_______
'Copy visible cells to new sheet "SureShip"
    'ActiveSheet.Range("A1:O" & LastRow).Select
    'Selection.Copy
    'Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SureShip"
    'Rows("1:1").Select
    'ActiveSheet.Paste

'Delete "Sheet1"
    'Sheets("Sheet1").Select
    'ActiveWindow.SelectedSheets.Delete
    
Application.ScreenUpdating = True   'turn screen updating back on

End Sub

