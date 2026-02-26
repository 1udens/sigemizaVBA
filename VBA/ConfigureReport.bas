Attribute VB_Name = "ConfigureReport"
Sub ConfigureReport()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsAlumnos As Worksheet, wsCursos As Worksheet, wsInscripciones As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateParts() As String, rawText As String
    Dim strDate1 As String, strDate2 As String

    Set wsAlumnos = ThisWorkbook.Worksheets("Alumnos")
    Set wsCursos = ThisWorkbook.Worksheets("Cursos")
    Set wsInscripciones = ThisWorkbook.Worksheets("Inscripciones")

'Alumnos configuration-----------------------------------------------------------------------------------------------------
    With wsAlumnos
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5 Then GoTo SkipAlumnos

        If .Cells(4, 11).Value <> "edad" Then
            .Range("K:K").Resize(ColumnSize:=2).Insert Shift:=xlToRight
            .Range("K4:L4").Value = Array("edad", "cursos")
        End If

        .Range("K5:K" & lastRow).Formula = "=IFERROR(INT(YEARFRAC([@[fecha_nacimiento]],TODAY())),"""")"
        .Range("L5:L" & lastRow).Formula = "=IFERROR(COUNTIF(Inscripciones!$C:$C,[@nombre]),0)"

        '.Range("F:F").NumberFormatLocal = "@"
        '.Range("H:H").NumberFormatLocal = "@"
        '.Range("J:J").NumberFormatLocal = "dd/mm/yyyy"
        '.Range("K:L").NumberFormatLocal = "0"
    End With
SkipAlumnos:

'Cursos configuration-----------------------------------------------------------------------------------------------------
With wsCursos
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5 Then GoTo SkipCursos

        If .Cells(4, 3).Value <> "codigo_curso" Then
            .Range("C:C").Insert Shift:=xlToRight
            .Range("C4").Value = "codigo_curso"
        End If

        .Range("C5:C" & lastRow).Formula = "=[@codigo] & "" - "" & [@curso]"

        If .Cells(4, 18).Value <> "duracion" Then
            .Columns("R").Insert Shift:=xlToRight
            .Cells(4, 17).Value = "financiador"
            .Cells(4, 18).Value = "duracion"
        End If

        .Range("Q5:Q" & lastRow).TextToColumns _
            Destination:=.Range("Q5"), _
            DataType:=xlDelimited, _
            Semicolon:=True, _
            Comma:=False, _
            Space:=False, _
            Other:=False
            
        .Columns("R").Replace What:=" ", Replacement:="", LookAt:=xlPart
        .Range("R5:R" & lastRow).Value = Application.Trim(.Range("R5:R" & lastRow).Value)

        '.Range("M:N").NumberFormatLocal = "dd/mm/yyyy"
        '.Range("C:C").NumberFormatLocal = "@"
        '.Range("K:K").NumberFormatLocal = "@"
        '.Range("O:O").NumberFormatLocal = "@"
        '.Range("Q:Q").NumberFormatLocal = "@"
        '.Range("R:R").NumberFormatLocal = "@"
    End With
SkipCursos:

'Inscripciones configuration-----------------------------------------------------------------------------------------------------
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
        
        If .Cells(4, 15).Value <> "txt_financiador" Then
            .Range("O4:P4").Value = Array("txt_financiador", "txt_duracion")
        End If
        
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "(\d{1,2}/\d{1,2}/\d{4}) al (\d{1,2}/\d{1,2}/\d{4})"

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
            "=IFERROR(XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[sexo]),"""")"
        .Range("H5:H" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[edad]),0)"
        .Range("I5:I" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[nacionalidad]),"""")"
        .Range("J5:J" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_alumno]],Table11[nombre],Table11[cursos]),0)"
        .Range("O5:O" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_curso]],Table12[codigo_curso],Table12[financiador]),"""")"
        .Range("P5:P" & lastRow).Formula = _
            "=IFERROR(XLOOKUP([@[txt_curso]],Table12[codigo_curso],Table12[duracion]),"""")"

        '.Range("C:D").NumberFormatLocal = "dd/mm/yyyy"
        '.Range("G:G").NumberFormatLocal = "@"
        '.Range("I:I").NumberFormatLocal = "@"
        '.Range("H:J").NumberFormatLocal = "0"
        '.Columns("C:J").AutoFit
    End With
SkipInsc:

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

