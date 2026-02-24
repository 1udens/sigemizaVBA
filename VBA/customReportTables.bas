Sub CreateAnalysisSheet()
    Dim wsAna As Worksheet
    Dim anaSheetName As String: anaSheetName = "Analysis"

    On Error Resume Next
    Set wsAna = ThisWorkbook.Worksheets(anaSheetName)
    On Error GoTo 0

    If wsAna Is Nothing Then
        Set wsAna = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAna.Name = anaSheetName
    Else
        wsAna.Cells.Delete
    End If

'Start Table 1------------------------------------------------------------------
    With wsAna.Range("B2:K8")
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .Borders.Color = RGB(244, 123, 61) 'Ta'Amay Orange
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With

    'Table title
    With wsAna.Range("B2:K2")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(244, 123, 61)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .Value = "Data courses in 4 Ta'Amay Centres"
    End With

    'Table subtitle (Timeframe)
    With wsAna.Range("B3:K3")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(244, 123, 61)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .Value = "(START - END)"
    End With

    'Data headers
    With wsAna.Range("B4:K5")
    .Font.Bold = True
    End With

    'Data column: Course type
    With wsAna.Range("B4:B5")
    .Merge
    .Value = "Course type"
    End With
    wsAna.Range("B6").Value = "Short courses"
    wsAna.Range("B7").Value = "Long courses"
    wsAna.Range("B8").Value = "All courses"

    'Data column: Course count
    With wsAna.Range("C4:C5")
    .Merge
    .Value = "Number of Courses Delivered"
    End With
    wsAna.Range("C6").Formula = "=0"
    wsAna.Range("C7").Formula = "=0"
    wsAna.Range("C8").Formula = "=COUNTA(PQ_Table12[codigo_curso])"

    'Data column: Participants by sex
    With wsAna.Range("D4:F4")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "Participants"
    End With
    wsAna.Range("D5").Value = "Total"
    wsAna.Range("D6").Formula = "=0"
    wsAna.Range("D7").Formula = "=0"
    wsAna.Range("D8").Formula = "=COUNTA(PQ_Table13[Sexo])"
    wsAna.Range("E5").Value = "Female"
    wsAna.Range("E6").Formula = "=0"
    wsAna.Range("E7").Formula = "=0"
    wsAna.Range("E8").Formula = "=COUNTIF(PQ_Table13[Sexo],""F"")"
    wsAna.Range("F5").Value = "Male"
    wsAna.Range("F6").Formula = "=0"
    wsAna.Range("F7").Formula = "=0"
    wsAna.Range("F8").Formula = "=COUNTIF(PQ_Table13[Sexo],""M"")"

    'Data column: Participants by age
    With wsAna.Range("G4:K4")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "Age ranges"
    End With
    wsAna.Range("G5").Value = "Under 14"
    wsAna.Range("H5").Value = "14-18"
    wsAna.Range("I5").Value = "18-30"
    wsAna.Range("J5").Value = "30-50"
    wsAna.Range("K5").Value = "Over 50"

    wsAna.Columns("B:K").AutoFit
'End Table 1------------------------------------------------------------------
End Sub