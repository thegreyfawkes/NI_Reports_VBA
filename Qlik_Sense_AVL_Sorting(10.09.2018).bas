Attribute VB_Name = "Qlik_Sense_AVL_Sorting"
Sub Qlik_Sense_AVL_Sorting()

'--DECLARATIONS------------------------------------
Lastrow = Range("A" & Rows.Count).End(xlUp).Row                     'Determines last row count
'--------------------------------------------------

'Turn on autofilter for "Site" column to display "HU" (DEB) entries
    ActiveSheet.Range("A1:Z" & Lastrow).AutoFilter Field:=1, Criteria1:="HU", Operator:=xlFilterValues
'Copy filtered results to new tab named "DEB"
    ActiveSheet.Range("A1:Z" & Lastrow).Select
    Selection.Copy
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "DEB"
    Rows("1:1").Select
    ActiveSheet.Paste
'Select Row 1:1 and freeze the header row
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
'Go back to main sheet ("Sheet1"), and unfilter old selection
    Sheets("Sheet1").Select
    ActiveSheet.AutoFilterMode = False
'Turn on autofilter for "Site" column to display "MY" (PEN) entries
    ActiveSheet.Range("A1:Z" & Lastrow).AutoFilter Field:=1, Criteria1:="MY", Operator:=xlFilterValues
    ActiveSheet.Range("A1:Z" & Lastrow).Select
    Selection.Copy
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "PEN"
    Rows("1:1").Select
    ActiveSheet.Paste
'Select Row 1:1 and freeze the header row
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
'Go back to main sheet ("Sheet1"), and unfilter old selection
    Sheets("Sheet1").Select
    ActiveSheet.AutoFilterMode = False
'Select Row 1:1 and freeze the header row
    Rows("1:1").Select
        With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    
    
    
End Sub
