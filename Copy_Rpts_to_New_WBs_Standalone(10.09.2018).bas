Attribute VB_Name = "Copy_Rpts_to_New_WBs_Standalone"
Sub AllInOne_Copy_to_New_WBs()
Attribute AllInOne_Copy_to_New_WBs.VB_ProcData.VB_Invoke_Func = " \n14"
'=== Declarations=========
Dim SureShipName As String
Dim Backlog_INTName As String
Dim Backlog_EXTName As String
Dim OTXName As String
DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
SureShipName = "SureShip_" & DateString
Backlog_INTName = "Daily_Backlog_ARROW_" & DateString
Backlog_EXTName = "NI_OTB_" & DateString
OTXName = "OTX_Report_" & DateString

ActiveWindow.Caption = "BaseFile"
'==========================
'=== Copy SureShip to new workbook
Workbooks.Add
ActiveWindow.Caption = "SureShip"
ActiveWorkbook.SaveAs Filename:=SureShipName
Windows("BaseFile").Activate
Worksheets("SureShip").Select
Cells.Select
Selection.Copy
Windows("SureShip").Activate
Sheets("Sheet1").Select
    Rows("1:1").Select
    ActiveSheet.Paste

'=== Copy Internal Backlog to new workbook
Workbooks.Add
ActiveWindow.Caption = "Backlog_INT"
ActiveWorkbook.SaveAs Filename:=Backlog_INTName
Windows("BaseFile").Activate
Worksheets("Backlog_INT").Select
Cells.Select
Selection.Copy
Windows("Backlog_INT").Activate
Sheets("Sheet1").Select
    Rows("1:1").Select
    ActiveSheet.Paste

'=== Copy External Backlog to new workbook
Workbooks.Add
ActiveWindow.Caption = "Backlog_EXT"
ActiveWorkbook.SaveAs Filename:=Backlog_EXTName
Windows("BaseFile").Activate
Worksheets("Backlog_EXT").Select
Cells.Select
Selection.Copy
Windows("Backlog_EXT").Activate
Sheets("Sheet1").Select
    Rows("1:1").Select
    ActiveSheet.Paste

'=== Copy OTX report to new workbook
Workbooks.Add
ActiveWindow.Caption = "OTX"
ActiveWorkbook.SaveAs Filename:=OTXName
Windows("BaseFile").Activate
Worksheets("OTX").Select
Cells.Select
Selection.Copy
Windows("OTX").Activate
Sheets("Sheet1").Select
    Rows("1:1").Select
    ActiveSheet.Paste

    
End Sub
