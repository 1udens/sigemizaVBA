Attribute VB_Name = "FormatImport"
Sub Formatting()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tblName As String
    Dim colFinalizo As Range
    
    Dim wsUnique As Worksheet
    Dim loUnique As ListObject
    
    ' 1. MAIN FORMATTING LOOP FOR IMPORT SHEETS ---------------------------------------
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 7) = "Import_" Then
            
            If ws.ListObjects.Count > 0 Then
                Set lo = ws.ListObjects(1)
                tblName = Mid(ws.Name, 8)
                
                ' Apply General Table Styles
                lo.TableStyle = "TableStyleMedium2"
                lo.Range.EntireColumn.AutoFit
                
                On Error Resume Next
                Select Case tblName
                    
                    Case "Table12"
                        With lo
                            .ListColumns("codigo_curso").DataBodyRange.NumberFormat = "@"
                            .ListColumns("jornada").DataBodyRange.NumberFormat = "@"
                            .ListColumns("cupo").DataBodyRange.NumberFormat = "0"
                            .ListColumns("lugar").DataBodyRange.NumberFormat = "@"
                            .ListColumns("observaciones").DataBodyRange.NumberFormat = "@"
                        End With

                    Case "Table13"
                        ' Clean up and Numeric-code the "txt_finalizo" status
                        Set colFinalizo = lo.ListColumns("txt_finalizo").DataBodyRange
                        
                        If Not colFinalizo Is Nothing Then
                            With colFinalizo
                                .Replace "Si finalizó + Certificado", "1", xlWhole
                                .Replace "Sí finalizó + Certificado", "1", xlWhole
                                .Replace "Sí finalizó", "2", xlWhole
                                .Replace "Si finalizó", "2", xlWhole
                                .Replace "En curso", "3", xlWhole
                                .Replace "No finalizó", "4", xlWhole
                                .Replace "Sólo se inscribió", "5", xlWhole
                                .NumberFormat = "0"
                            End With
                        End If

                        ' Format other columns in Table13
                        With lo
                            .ListColumns("txt_alumno").DataBodyRange.NumberFormat = "@"
                            .ListColumns("sexo").DataBodyRange.NumberFormat = "@"
                            .ListColumns("edad").DataBodyRange.NumberFormat = "0"
                            .ListColumns("nacionalidad").DataBodyRange.NumberFormat = "@"
                            .ListColumns("cursos_totales").DataBodyRange.NumberFormat = "0"
                        End With
                
                End Select
                On Error GoTo 0
            End If
        End If
    Next ws
    
    ' 2. UNIQUE TABLE REFRESH AND DEDUPLICATION --------------------------------------
    On Error Resume Next
    Set wsUnique = ThisWorkbook.Worksheets("PQ_Table13_Unique")
    Set loUnique = wsUnique.ListObjects("PQ_Table13_Unique")
    On Error GoTo 0

    If Not loUnique Is Nothing Then
        ' Refresh from Power Query
        loUnique.QueryTable.Refresh BackgroundQuery:=False

        ' Sort by status so "Certified" (1) stays during RemoveDuplicates
        With loUnique.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=loUnique.ListColumns("txt_finalizo").Range, _
                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With

        ' Deduplicate based on Student Name
        loUnique.Range.RemoveDuplicates Columns:=loUnique.ListColumns("txt_alumno").Index, Header:=xlYes
        
        ' Format Unique Table Columns
        With loUnique
            .ListColumns("txt_alumno").DataBodyRange.NumberFormat = "@"
            .ListColumns("sexo").DataBodyRange.NumberFormat = "@"
            .ListColumns("edad").DataBodyRange.NumberFormat = "0"
            .ListColumns("nacionalidad").DataBodyRange.NumberFormat = "@"
            .ListColumns("cursos_totales").DataBodyRange.NumberFormat = "0"
        End With
        
        wsUnique.Columns.AutoFit
    End If
    
    ' 3. GENERATE SUMMARY REPORT
    Call ReportTables
    
End Sub

