Attribute VB_Name = "CustomReportTables"
Sub CreateAnalysisSheet()
    Dim wsAna As Worksheet
    Dim anaSheetName As String: anaSheetName = "Analysis"
    
    ' 1. Check if the Analysis sheet exists; create if it doesn't
    On Error Resume Next
    Set wsAna = ThisWorkbook.Worksheets(anaSheetName)
    On Error GoTo 0
    
    If wsAna Is Nothing Then
        Set wsAna = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAna.Name = anaSheetName
    Else
        ' We use .Delete here instead of .Clear to completely kill
        ' any existing ListObjects/Tables so we can reuse the names.
        wsAna.Cells.Delete
    End If
    
    ' 2. Create Header for a "Metrics" Table
    ' Note: After a .Delete, we start fresh at cell B2
    wsAna.Range("B2").Value = "Metric Name"
    wsAna.Range("C2").Value = "Value"
    
    ' 3. Add Labels and Formulas
    wsAna.Range("B3").Value = "Total Registered"
    wsAna.Range("C3").Formula = "=COUNTA(PQ_Table13[nacionalidad])"
    
    wsAna.Range("B4").Value = "Average Age"
    wsAna.Range("C4").Formula = "=AVERAGE(PQ_Table13[edad])"
    
    wsAna.Range("B5").Value = "Total Cursos"
    wsAna.Range("C5").Formula = "=SUM(PQ_Table13[cursos_totales])"
    
    ' 4. Turn the range into a formal Excel Table
    Dim lo As ListObject
    Set lo = wsAna.ListObjects.Add(xlSrcRange, wsAna.Range("B2:C5"), , xlYes)
    lo.Name = "SummaryTable"
    lo.TableStyle = "TableStyleMedium2"
    
    ' 5. Auto-fit for cleanliness
    wsAna.Columns("B:C").AutoFit
End Sub
