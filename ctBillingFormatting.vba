Sub CT_Billing()
'
' CT_Billing Macro
'

'
    ' Add a filter to the raw data to get rid of rejected service notes
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Range("G3").Select
    ActiveSheet.Range("$A$3:$N$45").AutoFilter Field:=7, Criteria1:=Array( _
        "Approved", "In Review", "Update Required", "Draft"), Operator:=xlFilterValues
    
    ' Copy, paste and then reorganize the PTP names
    Columns("K:K").Select
    Range("K3").Activate
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Paste
    Range("B4").Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-1]), """", LEFT(RC[-1], FIND(""@"", SUBSTITUTE(RC[-1], "" "", ""@"", LEN(RC[-1]) - LEN(SUBSTITUTE(RC[-1], "" "", """"))))-1))"
    Range("B4").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Range("A4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Delete
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Participant Name"
    
    ' Copy, paste and then reorganize the service date
    Sheets("Code2").Select
    Columns("D:D").Select
    Range("D3").Activate
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("B1").Select
    ActiveSheet.Paste
    Range("C3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Date of Service"
    Range("C4").Select
    Selection.FormulaR1C1 = "=IF(ISBLANK(RC[-1]), """", LEFT(RC[-1], 10))"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, operation:=xlNone, _
        skipblanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Delete
    
    ' Copy and paste procedure codes
    Sheets("Code2").Select
    Columns("L:L").Select
    Range("L3").Activate
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("C3").Select
    Selection.Font.Bold = False
    ActiveCell.FormulaR1C1 = "Procedure Code"
    
    ' Get the duration and create "HOUR" and "Hours" colums from it
    Sheets("Code2").Select
    Columns("H:H").Select
    Range("H3").Activate
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("D1").Select
    ActiveSheet.Paste
    Range("D3").Select
    Selection.Font.Bold = False
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "HOUR"
    Range("E4").Select
    Selection.FormulaR1C1 = "=IF(ISBLANK(RC[-1]), """", RC[-1]/(24*60))"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, operation:=xlNone, _
        skipblanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "hh:mm;@"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-2]), """", INT(RC[-2]/60))"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "=IF(ISBLANK(RC[-3]), """", MOD(RC[-3], 60))"
    Range("H4").Select
    Selection.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-4]), """", RC[-2]+IF(RC[-1]<8, 0, IF(RC[-1]< 23, 0.25, IF(RC[-1]<38, 0.5, IF(RC[-1]<53, 0.75, 1)))))"
    Range("F4:H4").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, operation:=xlNone, _
        skipblanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Range("F4:G4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Delete
    
    ' Add "Hours", "Rate" and "Amount" columns' headers
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Hours"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "Rate"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "Amount"
    
    ' Bring DSP names and format properly
    Sheets("Code2").Select
    Columns("C:C").Select
    Range("C3").Activate
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("I1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("J4").Select
    Selection.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-1]), """", LEFT(RC[-1], FIND(""@"", SUBSTITUTE(RC[-1], "" "", ""@"", LEN(RC[-1]) - LEN(SUBSTITUTE(RC[-1], "" "", """"))))-1))"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, operation:=xlNone, _
        skipblanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, skipblanks _
        :=False, Transpose:=False
    Range("I4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Delete
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "DSP"
    
    ' Add "Payer" column's header
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Payer"
    
    ' Bring "Service Status" and "EVV Match Status" columns
    Sheets("Code2").Select
    Range("G:G,I:I").Select
    Range("I1").Activate
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("K1").Select
    ActiveSheet.Paste
    Range("K3:L3").Select
    Application.CutCopyMode = False
    Selection.Font.Bold = False
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("L1:L2").Select
    Range("L2").Activate
    Selection.EntireRow.Delete
    Range("L1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Font.Bold = True
    Range("A:L").Select
    Range("L1").Activate
    Selection.Columns.AutoFit
    Range("A:L").Select
    Range("K2").Activate
    
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Range("A1").Select
    
End Sub
