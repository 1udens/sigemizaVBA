Attribute VB_Name = "CustomReportTables"
Sub ReportTables()
    Dim wsAna As Worksheet
    Dim anaSheetName As String: anaSheetName = "Tables"

    ' 1. SETUP SHEET
    On Error Resume Next
    Set wsAna = ThisWorkbook.Worksheets(anaSheetName)
    On Error GoTo 0

    If wsAna Is Nothing Then
        Set wsAna = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAna.Name = anaSheetName
    Else
        wsAna.Cells.Delete
    End If

    ' START TABLE 1: TIMEFRAME SUMMARY ---------------------------------------
    
    ' Table Title
    With wsAna.Range("B2:I2")
        .Merge
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(244, 123, 61)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Value = "Timeframe Summary (All Centres)"
    End With
    
    ' Inscriptions Section Styles
    With wsAna.Range("B3:E10")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(244, 123, 61)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Inscriptions Headers
    With wsAna.Range("B3:E3")
        .Interior.Color = RGB(252, 228, 214)
        .Font.Bold = True
        .Value = Array("Course Status", "Female", "Male", "Total")
    End With
    
    ' Inscriptions Row Labels
    With wsAna.Range("B4:B10")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = Application.Transpose(Array("Total inscriptions", _
                                             "Certified", _
                                             "Not certified", _
                                             "In course", _
                                             "Withdrew", _
                                             "Inscribed only", _
                                             "Desertion %"))
    End With

    ' Formulas: Table 1 Rows
    wsAna.Range("C4:E4").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""<> "",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""<> "",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[sexo],""<> "")")
                                         
    wsAna.Range("C5:E5").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""1"",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""1"",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""1"")")
                                         
    wsAna.Range("C6:E6").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""2"",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""2"",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""2"")")
                                         
    wsAna.Range("C7:E7").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""3"",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""3"",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""3"")")
                                         
    wsAna.Range("C8:E8").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""4"",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""4"",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""4"")")
                                         
    wsAna.Range("C9:E9").Formula = Array("=COUNTIFS(PQ_Table13[txt_finalizo],""5"",PQ_Table13[sexo],""F"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""5"",PQ_Table13[sexo],""M"")", _
                                         "=COUNTIFS(PQ_Table13[txt_finalizo],""5"")")

    With wsAna.Range("C10:E10")
        .NumberFormat = "0%"
        .Formula = Array("=(C8+C9)/C4", "=(D8+D9)/D4", "=(E8+E9)/E4")
    End With
    
    ' Beneficiaries Section Styles
    With wsAna.Range("B11:E15")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(244, 123, 61)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Beneficiaries Headers
    With wsAna.Range("B11:E11")
        .Interior.Color = RGB(252, 228, 214)
        .Font.Bold = True
        .Value = Array("Beneficiaries", "Female", "Male", "Total")
    End With
    
    ' Beneficiaries Row Labels
    With wsAna.Range("B12:B15")
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Value = Application.Transpose(Array("Total beneficiaries", "Guatemalan", "Belizean", "Other"))
    End With

    ' Formulas: Beneficiaries
    wsAna.Range("C12:E12").Formula = Array("=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""F"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""M"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""<> "")")
                                           
    wsAna.Range("C13:E13").Formula = Array("=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""F"",PQ_TABLE13_UNIQUE[nacionalidad],""Guatemalteca"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""M"",PQ_TABLE13_UNIQUE[nacionalidad],""Guatemalteca"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""<> "",PQ_TABLE13_UNIQUE[nacionalidad],""Guatemalteca"")")
                                           
    wsAna.Range("C14:E14").Formula = Array("=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""F"",PQ_TABLE13_UNIQUE[nacionalidad],""Beliceña"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""M"",PQ_TABLE13_UNIQUE[nacionalidad],""Beliceña"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""<> "",PQ_TABLE13_UNIQUE[nacionalidad],""Beliceña"")")
                                           
    wsAna.Range("C15:E15").Formula = Array("=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""F"",PQ_TABLE13_UNIQUE[nacionalidad],""<>Beliceña"",PQ_TABLE13_UNIQUE[nacionalidad],""<>Guatemalteca"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""M"",PQ_TABLE13_UNIQUE[nacionalidad],""<>Beliceña"",PQ_TABLE13_UNIQUE[nacionalidad],""<>Guatemalteca"")", _
                                           "=COUNTIFS(PQ_TABLE13_UNIQUE[SEXO],""<> "",PQ_TABLE13_UNIQUE[nacionalidad],""<>Beliceña"",PQ_TABLE13_UNIQUE[nacionalidad],""<>Guatemalteca"")")
    
    ' Counts by Courses Taken Styles
    With wsAna.Range("G3:I8")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(244, 123, 61)
        .HorizontalAlignment = xlCenter
    End With
    
    With wsAna.Range("G3:I3")
        .Interior.Color = RGB(252, 228, 214)
        .Font.Bold = True
        .Value = Array("Courses", "Count", "%")
    End With
    
    wsAna.Range("G4:G8").Value = Application.Transpose(Array("1", "2", "3", "4", "+5"))
    wsAna.Range("G4:G8").Font.Bold = True

    ' Formulas: Course Counts and %
    wsAna.Range("H4:H8").Formula = Application.Transpose(Array("=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[cursos_totales],""1"")", _
                                                               "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[cursos_totales],""2"")", _
                                                               "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[cursos_totales],""3"")", _
                                                               "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[cursos_totales],""4"")", _
                                                               "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[cursos_totales],"">=5"")"))
    
    With wsAna.Range("I4:I8")
        .NumberFormat = "0%"
        .Formula = Application.Transpose(Array("=H4/E12", "=H5/E12", "=H6/E12", "=H7/E12", "=H8/E12"))
    End With

    ' Counts by Ages Styles
    With wsAna.Range("G10:I15")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(244, 123, 61)
        .HorizontalAlignment = xlCenter
    End With
    
    With wsAna.Range("G10:I10")
        .Interior.Color = RGB(252, 228, 214)
        .Font.Bold = True
        .Value = Array("Age", "Count", "%")
    End With
    
    wsAna.Range("G11:G15").Value = Application.Transpose(Array("Below 14", "14-18", "18-30", "30-50", "+50"))
    wsAna.Range("G11:G15").Font.Bold = True

    ' Formulas: Age Counts and %
    wsAna.Range("H11:H15").Formula = Application.Transpose(Array("=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[edad],""<14"")", _
                                                                 "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[edad],"">=14"",PQ_Table13_UNIQUE[edad],""<18"")", _
                                                                 "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[edad],"">=18"",PQ_Table13_UNIQUE[edad],""<30"")", _
                                                                 "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[edad],"">=30"",PQ_Table13_UNIQUE[edad],""<50"")", _
                                                                 "=COUNTIFS(PQ_Table13_UNIQUE[txt_finalizo],""<>5"",PQ_Table13_UNIQUE[edad],"">=50"")"))
    
    With wsAna.Range("I11:I15")
        .NumberFormat = "0%"
        .Formula = Application.Transpose(Array("=H11/E12", "=H12/E12", "=H13/E12", "=H14/E12", "=H15/E12"))
    End With

    ' Outer Border for Table 1
    wsAna.Range("B2:I15").BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=RGB(244, 123, 61)


    ' START TABLE 2: DATA COURSES ---------------------------------------------
    
    ' Table Layout Styles
    With wsAna.Range("B17:K23")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(244, 123, 61)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Titles and Subtitles
    With wsAna.Range("B17:K17")
        .Merge
        .Interior.Color = RGB(244, 123, 61)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Value = "Data courses in 4 Ta'Amay Centres"
    End With

    With wsAna.Range("B18:K18")
        .Merge
        .Interior.Color = RGB(244, 123, 61)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Value = "(START - END)"
    End With

    ' Header Merging
    wsAna.Range("B19:B20").Merge: wsAna.Range("B19:B20").Value = "Course type"
    wsAna.Range("C19:C20").Merge: wsAna.Range("C19:C20").Value = "Number of Courses Delivered"
    
    With wsAna.Range("D19:F19")
        .Merge
        .Value = "Participants"
    End With
    
    With wsAna.Range("G19:K19")
        .Merge
        .Value = "Age ranges"
    End With
    
    wsAna.Range("B19:K20").Font.Bold = True

    ' Row Labels and Column Sub-headers
    wsAna.Range("B21:B23").Value = Application.Transpose(Array("Short courses", "Long courses", "All courses"))
    wsAna.Range("D20:K20").Value = Array("Total", "Female", "Male", "Under 14", "14-18", "18-30", "30-50", "Over 50")

    ' Formulas: Course Counts
    wsAna.Range("C21").Formula = "=COUNTIFS(PQ_Table12[duracion],""SHORT"")"
    wsAna.Range("C22").Formula = "=COUNTIFS(PQ_Table12[duracion],""LONG"")"
    wsAna.Range("C23").Formula = "=C21+C22" ' Efficient summing

    ' Formulas: Participants (Short, Long, Total)
    wsAna.Range("D21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("D22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("D23").Formula = "=D21+D22"

    wsAna.Range("E21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("E22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""F"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("E23").Formula = "=E21+E22"

    wsAna.Range("F21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("F22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[Sexo],""M"",PQ_Table13[txt_finalizo],""<>5"")"
    wsAna.Range("F23").Formula = "=F21+F22"

    ' Formulas: Age Ranges
    wsAna.Range("G21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"")"
    wsAna.Range("G22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],""<14"")"
    wsAna.Range("G23").Formula = "=G21+G22"

    wsAna.Range("H21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"")"
    wsAna.Range("H22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=14"",PQ_Table13[edad],""<18"")"
    wsAna.Range("H23").Formula = "=H21+H22"

    wsAna.Range("I21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"")"
    wsAna.Range("I22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=18"",PQ_Table13[edad],""<30"")"
    wsAna.Range("I23").Formula = "=I21+I22"

    wsAna.Range("J21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"")"
    wsAna.Range("J22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=30"",PQ_Table13[edad],""<50"")"
    wsAna.Range("J23").Formula = "=J21+J22"

    wsAna.Range("K21").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""SHORT"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"")"
    wsAna.Range("K22").Formula = "=COUNTIFS(PQ_Table13[txt_duracion],""LONG"",PQ_Table13[txt_finalizo],""<>5"",PQ_Table13[edad],"">=50"")"
    wsAna.Range("K23").Formula = "=K21+K22"

    ' Auto-fit for professional look
    'wsAna.Columns("B:K").AutoFit

End Sub

