Sub CT_Billing()
    'It converts SN data into desired format
    
    'Determine the required variables
    Dim i, lRow As Integer
    Dim sn, ts As Worksheet
    
    'Set the worksheets
    Set sn = ActiveWorkbook.Sheets(1)
    sn.Copy after:=sn
    Set ts = ActiveWorkbook.Sheets(2)
    
    With ts
        lRow = .Cells.Find("*", searchdirection:=xlPrevious, searchorder:=xlByRows).Row
    
        'Delete unnecessary columns and rearrange the rest
        .Range("A:B, E:F, J:J, M:N").Delete
        .Columns("F").Cut
        .Columns("A").Insert shift:=xlToRight
        .Columns("C").Cut
        .Columns("B").Insert shift:=xlToRight
        .Columns("G").Cut
        .Columns("C").Insert shift:=xlToRight
        .Columns("F").Cut
        .Columns("D").Insert shift:=xlToRight
        .Columns("E:H").Insert shift:=xlToRight
        .Columns("J").Insert shift:=xlToRight
        Application.CutCopyMode = False
        .Range("A1:L1").Value = Array("PTP", "Date of Service", "Proc. Code", "Duration", "Hours", _
            "Billing Hours", "Rate", "Amount", "DSP", "Payer", "SN Status", "EVV Match Status")
        For i = lRow To 2 Step -1
            On Error Resume Next
            If Trim(.Cells(i, "K").Value) = "Rejected" Then
                .Rows(i).EntireRow.Delete
            Else
                .Cells(i, "A").Value = Left(.Cells(i, "A").Value, InStr(.Cells(i, "A").Value, "(") - 2)
                .Cells(i, "I").Value = Left(.Cells(i, "I").Value, InStrRev(.Cells(i, "I").Value, " ") - 1)
                .Cells(i, "E").Value = .Cells(i, "D").Value / (24 * 60)
                .Cells(i, "E").NumberFormat = "hh:mm"
                If .Cells(i, "D").Value Mod 60 < 8 Then
                    .Cells(i, "F").Value = Int(.Cells(i, "D").Value / 60)
                ElseIf .Cells(i, "D").Value Mod 60 < 23 Then
                    .Cells(i, "F").Value = Int(.Cells(i, "D").Value / 60) + 0.25
                ElseIf .Cells(i, "D").Value Mod 60 < 38 Then
                    .Cells(i, "F").Value = Int(.Cells(i, "D").Value / 60) + 0.5
                ElseIf .Cells(i, "D").Value Mod 60 < 53 Then
                    .Cells(i, "F").Value = Int(.Cells(i, "D").Value / 60) + 0.75
                Else
                    .Cells(i, "F").Value = Int(.Cells(i, "D").Value / 60) + 1
                End If
            End If
        Next i
    End With
End Sub
