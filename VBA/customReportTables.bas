Attribute VB_Name = "CustomReportTables"
Sub ReportTables()
    Dim wsAna As Worksheet
    Dim anaSheetName As String: anaSheetName = "Tables"

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
    

'Start Table 2------------------------------------------------------------------
    With wsAna.Range("B17:K23")
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .Borders.Color = RGB(244, 123, 61) 'Ta'Amay Orange
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With

    'Table title
    With wsAna.Range("B17:K17")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(244, 123, 61)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .Value = "Data courses in 4 Ta'Amay Centres"
    End With

    'Table subtitle (Timeframe)
    With wsAna.Range("B18:K18")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(244, 123, 61)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .Value = "(START - END)"
    End With

    'Data headers
    With wsAna.Range("B19:K20")
    .Font.Bold = True
    End With

    'Data column: Course type
    With wsAna.Range("B19:B20")
    .Merge
    .Value = "Course type"
    End With
    wsAna.Range("B21").Value = "Short courses"
    wsAna.Range("B22").Value = "Long courses"
    wsAna.Range("B23").Value = "All courses"

    'Data column: Course count
    With wsAna.Range("C19:C20")
    .Merge
    .Value = "Number of Courses Delivered"
    End With
    wsAna.Range("C21").Formula = "=COUNTIFS(PQ_Table12[duracion],""SHORT"")"
    wsAna.Range("C22").Formula = "=COUNTIFS(PQ_Table12[duracion],""LONG"")"
    wsAna.Range("C23").Formula = "=COUNTIFS(PQ_Table12[duracion],""LONG"") + COUNTIFS(PQ_Table12[duracion],""SHORT"")"

    'Data column: Participants by sex
    With wsAna.Range("D19:F19")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "Participants"
    End With
    wsAna.Range("D20").Value = "Total"
    wsAna.Range("D21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("D22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("D23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"") + COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("E20").Value = "Female"
    wsAna.Range("E21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("E22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("E23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"") + COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("F20").Value = "Male"
    wsAna.Range("F21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("F22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("F23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"") + COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"")"

    'Data column: Participants by age
    With wsAna.Range("G19:K19")
    .Merge
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = "Age ranges"
    End With
    wsAna.Range("G20").Value = "Under 14"
    wsAna.Range("G21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"")"
    wsAna.Range("G22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"")"
    wsAna.Range("G23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"") + COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"")"
    wsAna.Range("H20").Value = "14-18"
    wsAna.Range("H21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"")"
    wsAna.Range("H22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"")"
    wsAna.Range("H23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"") + COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"")"
    wsAna.Range("I20").Value = "18-30"
    wsAna.Range("I21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"")"
    wsAna.Range("I22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"")"
    wsAna.Range("I23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"") + COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"")"
    wsAna.Range("J20").Value = "30-50"
    wsAna.Range("J21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"")"
    wsAna.Range("J22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"")"
    wsAna.Range("J23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"") + COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"")"
    wsAna.Range("K20").Value = "Over 50"
    wsAna.Range("K21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"")"
    wsAna.Range("K22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"")"
    wsAna.Range("K23").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"") + COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"")"

'End Table 2------------------------------------------------------------------
End Sub
