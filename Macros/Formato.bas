Attribute VB_Name = "Module3"
Sub Formato()
Attribute Formato.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Formato Macro
'

'
    Columns("B:D").Select
    Selection.Cut Destination:=Columns("A:C")
    Columns("G:P").Select
    Selection.Cut Destination:=Columns("E:N")
    Columns("J:M").Select
    Selection.ColumnWidth = 8.57
    Columns("I:I").ColumnWidth = 8.57
    Selection.Cut Destination:=Columns("I:L")
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Part Num."
    Range("A3").Select
    Selection.Copy
    Range("E3").Select
    ActiveSheet.Paste
    Range("I3").Select
    ActiveSheet.Paste
    Range("C2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "P-pack per LU"
    Columns("A:L").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$L"), , xlYes).Name = _
        "Table1"
    Columns("D:D").Select
    Selection.ListObject.ListColumns.Add Position:=4
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "P-pack per level"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-9
    Columns("B:C").Select
    Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("E2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=[@[LU_Qty]]/[@[Primary Pck Qty]]"
    Range("E3").Select
    ActiveWindow.SmallScroll Down:=-39
    Range("D2").Select
    ActiveWindow.SmallScroll Down:=-15
    Columns("D:E").Select
    Range("Table1[[#Headers],[P-pack per LU]]").Activate
    Selection.NumberFormat = "0"
    ActiveWindow.SmallScroll Down:=-18
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=([@[LU_Length_(mm)]]/[@[P-pack L]])*([@[LU_Width_(mm)]]/[@[P-pack W]])"
    Range("D3").Select
    ActiveWindow.SmallScroll Down:=3
    Range("G31").Select
    ActiveWindow.SmallScroll Down:=-27
End Sub
