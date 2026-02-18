' ────────────────────────────────────────────────────────────
' ConfigureSheets
' ────────────────────────────────────────────────────────────
Sub ConfigureSheets()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsAlumnos As Worksheet, wsCursos As Worksheet, wsInscripciones As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateParts() As String, rawText As String
    Dim strDate1 As String, strDate2 As String

    If Not SheetExists("Alumnos") Or Not SheetExists("Cursos") Or Not SheetExists("Inscripciones") Then
        MsgBox "No se encontraron las hojas requeridas: Alumnos, Cursos, Inscripciones." & _
               vbCrLf & "Asegúrese de importar primero los datos desde la base de datos.", vbCritical
        GoTo Cleanup
    End If

    Set wsAlumnos       = ThisWorkbook.Worksheets("Alumnos")
    Set wsCursos        = ThisWorkbook.Worksheets("Cursos")
    Set wsInscripciones = ThisWorkbook.Worksheets("Inscripciones")

    ' ── ALUMNOS ──────────────────────────────────────────────
    With wsAlumnos
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5 Then GoTo SkipAlumnos

        If .Cells(4, 11).Value <> "edad" Then
            .Range("K:K").Resize(ColumnSize:=2).Insert Shift:=xlToRight
            .Range("K4:L4").Value = Array("edad", "cursos")
        End If

        .Range("K5:K" & lastRow).Formula = "=IFERROR(INT(YEARFRAC([@[fecha_nacimiento]],TODAY())),"""")"
        .Range("L5:L" & lastRow).Formula = "=IFERROR(COUNTIF(Inscripciones!$C:$C,[@nombre]),0)"

        .Range("F:F").NumberFormatLocal = "@"
        .Range("H:H").NumberFormatLocal = "@"
        .Range("J:J").NumberFormatLocal = "dd/mm/yyyy"
        .Range("K:L").NumberFormatLocal = "0"
    End With
SkipAlumnos:

    ' ── CURSOS ───────────────────────────────────────────────
    With wsCursos
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5 Then GoTo SkipCursos

        If .Cells(4, 3).Value <> "codigo_curso" Then
            .Range("C:C").Resize(ColumnSize:=1).Insert Shift:=xlToRight
            .Range("C4").Value = "codigo_curso"
        End If

        .Range("C5:C" & lastRow).Formula = "=[@codigo] & "" - "" & [@curso]"

        .Range("M:N").NumberFormatLocal = "dd/mm/yyyy"
        .Range("C:C").NumberFormatLocal = "@"
        .Range("K:K").NumberFormatLocal = "@"
        .Range("O:O").NumberFormatLocal = "@"
    End With
SkipCursos:

    ' ── INSCRIPCIONES ────────────────────────────────────────
    With wsInscripciones
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        If lastRow < 5 Then GoTo SkipInsc

        If .Cells(4, 3).Value <> "vigencia_inicio" Then
            .Range("C:D").Insert Shift:=xlToRight
            .Range("C4:D4").Value = Array("vigencia_inicio", "vigencia_final")
        End If

        If .Cells(4, 7).Value <> "sexo" Then
            .Range("G:J").Insert Shift:=xlToRight
            .Range("G4:J4").Value = Array("sexo", "edad", "nacionalidad", "cursos_totales")
        End If

        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\d{2}/\d{2}/\d{4}) al (\d{2}/\d{2}/\d{4})"

        For i = 5 To lastRow
            rawText = CStr(.Cells(i, 2).Value)
            If re.Test(rawText) Then
                Dim m As Object
                Set m = re.Execute(rawText)
                strDate1 = m(0).SubMatches(0)
                strDate2 = m(0).SubMatches(1)
                If IsDate(strDate1) Then .Cells(i, 3).Value = CDate(strDate1)
                If IsDate(strDate2) Then .Cells(i, 4).Value = CDate(strDate2)
            End If
        Next i

        .Range("G5:G" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Alumnos!$A:$A,Alumnos!$H:$H),"""")"
        .Range("H5:H" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Alumnos!$A:$A,Alumnos!$K:$K),0)"
        .Range("I5:I" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Alumnos!$A:$A,Alumnos!$F:$F),"""")"
        .Range("J5:J" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Alumnos!$A:$A,Alumnos!$L:$L),0)"

        ' Formatos
        .Range("C:D").NumberFormatLocal = "dd/mm/yyyy"
        .Range("G:G").NumberFormatLocal = "@"
        .Range("I:I").NumberFormatLocal = "@"
        .Range("H:J").NumberFormatLocal = "0"
        .Columns("C:J").AutoFit
    End With
SkipInsc:

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Configuration complete", vbInformation
End Sub