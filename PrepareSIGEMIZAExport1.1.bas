Sub ConfigureSheets()
    Application.ScreenUpdating = False

    Dim wsAlumnos As Worksheet, wsCursos As Worksheet, wsInscripciones As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateParts() As String
    Dim rawText As String

    Set wsAlumnos = ThisWorkbook.Worksheets("Alumnos")
    Set wsCursos = ThisWorkbook.Worksheets("Cursos")
    Set wsInscripciones = ThisWorkbook.Worksheets("Inscripciones")

    ' ==========================================
    ' SHEET: ALUMNOS
    ' ==========================================
    With wsAlumnos
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("K:K").Resize(ColumnSize:=2).Insert Shift:=xlToRight
        .Range("K4:L4").Value = Array("edad", "cursos")
        
        .Range("K5:K" & lastRow).Formula = "=YEARFRAC([@[fecha_nacimiento]], TODAY())"
        .Range("L5:L" & lastRow).Formula = "=COUNTIF(Table13[txt_alumno], A5)"

        .Range("F:F").NumberFormatLocal = "@"
        .Range("H:H").NumberFormatLocal = "@"
        .Range("J:J").NumberFormatLocal = "d/mm/yyyy"
        .Range("K:L").NumberFormatLocal = "0"
    End With

    ' ==========================================
    ' SHEET: CURSOS
    ' ==========================================
    With wsCursos
        .Range("C:C").Resize(ColumnSize:=1).Insert Shift:=xlToRight
        .Range("C4").Value = "codigo_curso"

        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("C5:C" & lastRow).Formula = "=[@codigo] & "" - "" & [@curso]"

        .Range("M:N").NumberFormatLocal = "d/mm/yyyy"
        .Range("C:C").NumberFormatLocal = "@"
        .Range("K:K").NumberFormatLocal = "@"
        .Range("O:O").NumberFormatLocal = "@"
    End With

    ' ==========================================
    ' SHEET: INSCRIPCIONES
    ' ==========================================
    With wsInscripciones
        .Range("C:D").Insert Shift:=xlToRight        
        .Range("F:F").Resize(ColumnSize:=4).Insert Shift:=xlToRight        
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row

        .Range("C4:D4").Value = Array("vigencia_inicio", "vigencia_final")
        .Range("F4:I4").Value = Array("sexo", "edad", "nacionalidad", "cursos")

       For i = 5 To lastRow
            rawText = .Range("B" & i).Value
            If InStr(rawText, " al ") > 0 Then
                dateParts = Split(rawText, " ")

                Dim strDate1 As String
                strDate1 = Trim(dateParts(0))
                If IsDate(strDate1) Then
                    .Range("C" & i).Value = CDate(strDate1)
                End If

                Dim strDate2 As String
                strDate2 = Trim(Replace(dateParts(2), ".", ""))
                If IsDate(strDate2) Then
                    .Range("D" & i).Value = CDate(strDate2)
                End If
            End If
        Next i

        .Range("F5:F" & lastRow).Formula = "=XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[sexo])"
        .Range("G5:G" & lastRow).Formula = "=XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[edad])"
        .Range("H5:H" & lastRow).Formula = "=XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[nacionalidad])"
        .Range("I5:I" & lastRow).Formula = "=XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[cursos])"

        .Range("C:D").NumberFormatLocal = "d/mm/yyyy"
        .Range("F:F").NumberFormatLocal = "@"
        .Range("H:H").NumberFormatLocal = "@"
        .Range("G:G").NumberFormatLocal = "0"
        .Range("I:I").NumberFormatLocal = "0"
        .Columns("C:I").AutoFit
    End With

    Application.ScreenUpdating = True
    MsgBox "Export Prepared.", vbInformation
End Sub

