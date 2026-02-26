Attribute VB_Name = "FormatImport"
Sub Formatting()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tblName As String
    Dim colFinalizo As Range
    
    Dim wsUnique As Worksheet
    Dim loUnique As ListObject
    
    'Formatting--------------------------------------------------------------------------------------------
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 7) = "Import_" Then
            
            If ws.ListObjects.Count > 0 Then
                Set lo = ws.ListObjects(1)
                tblName = Mid(ws.Name, 8)
                
                lo.TableStyle = "TableStyleMedium2"
                lo.Range.EntireColumn.AutoFit
                
                On Error Resume Next
                Select Case tblName
                                                
                Case "Table12"
                    lo.ListColumns("codigo_curso").DataBodyRange.NumberFormatLocal = "@"
                    lo.ListColumns("jornada").DataBodyRange.NumberFormatLocal = "@"
                    lo.ListColumns("cupo").DataBodyRange.NumberFormatLocal = "0"
                    lo.ListColumns("lugar").DataBodyRange.NumberFormatLocal = "@"
                    lo.ListColumns("observaciones").DataBodyRange.NumberFormatLocal = "@"

                Case "Table13"
                    Set colFinalizo = lo.ListColumns("txt_finalizo").DataBodyRange
                    If Not colFinalizo Is Nothing Then
                        colFinalizo.Replace "Si finalizó + Certificado", "1", xlWhole
                        colFinalizo.Replace "Sí finalizó + Certificado", "1", xlWhole
                        colFinalizo.Replace "Sí finalizó", "2", xlWhole
                        colFinalizo.Replace "Si finalizó", "2", xlWhole
                        colFinalizo.Replace "En curso", "3", xlWhole
                        colFinalizo.Replace "No finalizó", "4", xlWhole
                        colFinalizo.Replace "Sólo se inscribió", "5", xlWhole
                        colFinalizo.NumberFormat = "0"
                    End If

                    lo.ListColumns("txt_alumno").DataBodyRange.NumberFormatLocal = "@"
                    lo.ListColumns("sexo").DataBodyRange.NumberFormat = "@"
                    lo.ListColumns("edad").DataBodyRange.NumberFormat = "0"
                    lo.ListColumns("nacionalidad").DataBodyRange.NumberFormat = "@"
                    lo.ListColumns("cursos_totales").DataBodyRange.NumberFormat = "0"
                
                End Select
                On Error GoTo 0
            End If
        End If
    Next ws
    
        'Unique Table Refresh-------------------------------------------------------------------------------
    On Error Resume Next
    Set wsUnique = ThisWorkbook.Worksheets("PQ_Table13_Unique")
    Set loUnique = wsUnique.ListObjects("PQ_Table13_Unique")
    On Error GoTo 0

    If Not loUnique Is Nothing Then
        loUnique.QueryTable.Refresh BackgroundQuery:=False

        With loUnique.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=loUnique.ListColumns("txt_finalizo").Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With

        loUnique.Range.RemoveDuplicates Columns:=loUnique.ListColumns("txt_alumno").Index, Header:=xlYes
        
        lo.ListColumns("txt_alumno").DataBodyRange.NumberFormatLocal = "@"
        lo.ListColumns("sexo").DataBodyRange.NumberFormat = "@"
        lo.ListColumns("edad").DataBodyRange.NumberFormat = "0"
        lo.ListColumns("nacionalidad").DataBodyRange.NumberFormat = "@"
        lo.ListColumns("cursos_totales").DataBodyRange.NumberFormat = "0"
        
        wsUnique.Columns.AutoFit
    End If
    
    Call ReportTables
    
End Sub

