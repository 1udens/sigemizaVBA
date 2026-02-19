Attribute VB_Name = "FormattingModule"
Sub ApplyAllFormats()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tblName As String

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 7) = "Import_" Then
            
            If ws.ListObjects.Count > 0 Then
                Set lo = ws.ListObjects(1)
                tblName = Mid(ws.Name, 8)
                
                lo.TableStyle = "TableStyleMedium2"
                lo.Range.EntireColumn.AutoFit
                
                On Error Resume Next
                Select Case tblName
                
                Case "Table13"
                lo.ListColumns("txt_alumno").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("vigencia_inicio").DataBodyRange.NumberFormat = "dd/mm/yyyy"
                lo.ListColumns("vigencia_final").DataBodyRange.NumberFormat = "dd/mm/yyyy"
                lo.ListColumns("fecha_de_inscripcion").DataBodyRange.NumberFormat = "dd/mm/yyyy"
                lo.ListColumns("sexo").DataBodyRange.NumberFormat = "@"
                lo.ListColumns("edad").DataBodyRange.NumberFormat = "0"
                lo.ListColumns("nacionalidad").DataBodyRange.NumberFormat = "@"
                lo.ListColumns("cursos_totales").DataBodyRange.NumberFormat = "0"
                
                Case "Table12"
                lo.ListColumns("codigo_curso").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("jornada").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("fecha_de_inicio").DataBodyRange.NumberFormatLocal = "dd/mm/yyyy"
                lo.ListColumns("fecha_de_finalizacion").DataBodyRange.NumberFormatLocal = "dd/mm/yyyy"
                lo.ListColumns("cupo").DataBodyRange.NumberFormatLocal = "0"
                lo.ListColumns("lugar").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("observaciones").DataBodyRange.NumberFormatLocal = "@"
                
                Case "Table11"
                lo.ListColumns("nombre").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("nacionalidad").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("sexo").DataBodyRange.NumberFormatLocal = "@"
                lo.ListColumns("fecha_nacimiento").DataBodyRange.NumberFormatLocal = "dd/mm/yyyy"
                lo.ListColumns("edad").DataBodyRange.NumberFormatLocal = "0"
                lo.ListColumns("cursos").DataBodyRange.NumberFormatLocal = "0"
                lo.ListColumns("nombre").DataBodyRange.NumberFormatLocal = "@"
                
                End Select
                On Error GoTo 0
            End If
        End If
    Next ws
End Sub
