Attribute VB_Name = "Bklg_DEL_INT_COLS_Standalone"
Sub DEL_INT_Cols()
Attribute DEL_INT_Cols.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DEL_INT_Cols Macro

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
End Sub

