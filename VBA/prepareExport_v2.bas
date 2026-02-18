'Attribute VB_Name = "PrepareExport_v2"
' ============================================================
' PrepareExport_v2.bas
' Módulo de gestión de reportes — Sistema de Cursos
' ============================================================
' Submacros disponibles:
'   1. ConfigureSheets       — Prepara columnas calculadas (original mejorado)
'   2. CreatePivotTables     — Crea tablas dinámicas desde los datos limpios
'   3. RefreshAllPivots      — Actualiza todas las tablas dinámicas del libro
'   4. FilterByDate          — Aplica filtros de fecha en Inscripciones_Data
'   5. ExportFilteredReport  — Exporta resumen filtrado a nueva hoja
'   6. CleanupHelperColumns  — Revierte columnas calculadas (útil al re-exportar)
' ============================================================

Option Explicit

' ────────────────────────────────────────────────────────────
' 1. ConfigureSheets (versión mejorada del original)
' ────────────────────────────────────────────────────────────
Sub ConfigureSheets()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsAlumnos As Worksheet, wsCursos As Worksheet, wsInscripciones As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateParts() As String, rawText As String
    Dim strDate1 As String, strDate2 As String

    ' Verificar que las hojas existan
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

        ' Insertar columnas edad y cursos si no existen
        If .Cells(4, 11).Value <> "edad" Then
            .Range("K:K").Resize(ColumnSize:=2).Insert Shift:=xlToRight
            .Range("K4:L4").Value = Array("edad", "cursos")
        End If

        .Range("K5:K" & lastRow).Formula = "=IFERROR(INT(YEARFRAC([@[fecha_nacimiento]],TODAY())),"""")"
        .Range("L5:L" & lastRow).Formula = "=IFERROR(COUNTIF(Inscripciones!$C:$C,[@nombre]),0)"

        ' Formatear columnas
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

        ' Columnas vigencia
        If .Cells(4, 3).Value <> "vigencia_inicio" Then
            .Range("C:D").Insert Shift:=xlToRight
            .Range("C4:D4").Value = Array("vigencia_inicio", "vigencia_final")
        End If

        ' Columnas demográficas
        If .Cells(4, 7).Value <> "sexo" Then
            .Range("G:J").Insert Shift:=xlToRight
            .Range("G4:J4").Value = Array("sexo", "edad", "nacionalidad", "cursos_totales")
        End If

        ' Parsear fechas de vigencia desde texto "dd/mm/yyyy al dd/mm/yyyy."
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

        ' XLOOKUP de datos del alumno (requiere Excel 365/2019+)
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
    MsgBox "Configuración completada.", vbInformation
End Sub

' ────────────────────────────────────────────────────────────
' 2. CreatePivotTables — Crea/Reemplaza tablas dinámicas
' ────────────────────────────────────────────────────────────
Sub CreatePivotTables()
    Application.ScreenUpdating = False

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim dataRange As Range

    ' ── Pivot 1: Inscripciones por Sede y Estado ──────────────
    If Not SheetExists("Inscripciones_Data") Then
        MsgBox "Hoja 'Inscripciones_Data' no encontrada. Ejecute primero ConfigureSheets.", vbExclamation
        GoTo Cleanup
    End If

    Set wsData = ThisWorkbook.Worksheets("Inscripciones_Data")
    Set dataRange = wsData.Range("A2").CurrentRegion

    ' Eliminar pivot anterior si existe
    DeleteSheetIfExists "PT_Por_Sede"
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "PT_Por_Sede"
    wsPivot.Tab.Color = RGB(46, 117, 182)

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PivotPorSede")

    With pt
        .PivotFields("txt_lugar").Orientation = xlRowField
        .PivotFields("txt_finalizo").Orientation = xlColumnField
        .AddDataField .PivotFields("txt_alumno"), "Inscripciones", xlCount
        .NullString = "0"
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium9"
    End With

    ' Agregar slicer de fecha si está disponible (Excel 2013+)
    On Error Resume Next
    Dim slFecha As SlicerCache
    Set slFecha = ThisWorkbook.SlicerCaches.Add2(pt, "fecha_de_inscripcion", , xlTimeline)
    slFecha.TimelineState.SetFilterDateRange _
        CDate("01/01/2022"), CDate("31/12/2026")
    On Error GoTo 0

    wsPivot.Range("A1").Value = "INSCRIPCIONES POR SEDE Y ESTADO DE FINALIZACIÓN"
    wsPivot.Range("A1").Font.Bold = True
    wsPivot.Range("A1").Font.Size = 13
    wsPivot.Range("A2").Value = "Filtro: use la Escala de tiempo (slicer) para filtrar por fecha de inscripción"
    wsPivot.Range("A2").Font.Italic = True
    wsPivot.Range("A2").Font.Color = RGB(80, 80, 80)

    ' ── Pivot 2: Inscripciones por Jornada y Sede ─────────────
    DeleteSheetIfExists "PT_Por_Jornada"
    Dim wsPivot2 As Worksheet
    Set wsPivot2 = ThisWorkbook.Worksheets.Add
    wsPivot2.Name = "PT_Por_Jornada"
    wsPivot2.Tab.Color = RGB(23, 165, 137)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot2.Range("A3"), _
        TableName:="PivotPorJornada")

    With pt
        .PivotFields("txt_jornada").Orientation = xlRowField
        .PivotFields("txt_lugar").Orientation = xlColumnField
        .AddDataField .PivotFields("txt_alumno"), "Inscripciones", xlCount
        .NullString = "0"
        .TableStyle2 = "PivotStyleMedium2"
    End With

    wsPivot2.Range("A1").Value = "INSCRIPCIONES POR JORNADA Y SEDE"
    wsPivot2.Range("A1").Font.Bold = True
    wsPivot2.Range("A1").Font.Size = 13

    ' ── Pivot 3: Asistencia mensual ───────────────────────────
    If SheetExists("Asistencia_Data") Then
        Set wsData = ThisWorkbook.Worksheets("Asistencia_Data")
        Set dataRange = wsData.Range("A2").CurrentRegion

        DeleteSheetIfExists "PT_Asistencia"
        Dim wsPivot3 As Worksheet
        Set wsPivot3 = ThisWorkbook.Worksheets.Add
        wsPivot3.Name = "PT_Asistencia"
        wsPivot3.Tab.Color = RGB(231, 76, 60)

        Set pc = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=dataRange, _
            Version:=xlPivotTableVersion15)

        Set pt = pc.CreatePivotTable( _
            TableDestination:=wsPivot3.Range("A3"), _
            TableName:="PivotAsistencia")

        With pt
            Dim pfFecha As PivotField
            Set pfFecha = .PivotFields("fecha")
            pfFecha.Orientation = xlRowField
            pfFecha.AutoGroup   ' Agrupa automáticamente por mes/año

            .PivotFields("lugar_donde_se_imparte").Orientation = xlColumnField
            .AddDataField .PivotFields("porcentaje_de_asistencia"), "% Asistencia Prom.", xlAverage
            .DataFields("% Asistencia Prom.").NumberFormat = "0.0%"
            .TableStyle2 = "PivotStyleMedium5"
        End With

        wsPivot3.Range("A1").Value = "PORCENTAJE DE ASISTENCIA POR MES Y SEDE"
        wsPivot3.Range("A1").Font.Bold = True
        wsPivot3.Range("A1").Font.Size = 13
    End If

Cleanup:
    Application.ScreenUpdating = True
    MsgBox "Tablas dinámicas creadas." & vbCrLf & _
           "  • PT_Por_Sede     — Inscripciones por sede y estado" & vbCrLf & _
           "  • PT_Por_Jornada  — Inscripciones por jornada y sede" & vbCrLf & _
           "  • PT_Asistencia   — % asistencia por mes y sede", vbInformation
End Sub

' ────────────────────────────────────────────────────────────
' 3. RefreshAllPivots
' ────────────────────────────────────────────────────────────
Sub RefreshAllPivots()
    Application.ScreenUpdating = False
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        On Error Resume Next
        pc.Refresh
        On Error GoTo 0
    Next pc
    Application.ScreenUpdating = True
    MsgBox "Todas las tablas dinámicas actualizadas.", vbInformation
End Sub

' ────────────────────────────────────────────────────────────
' 4. FilterByDate — Aplica AutoFilter en Inscripciones_Data
'    Lee fechas desde la hoja PANEL (C5=inicio, C6=fin)
' ────────────────────────────────────────────────────────────
Sub FilterByDate()
    Dim wsPANEL As Worksheet
    Dim wsInsc  As Worksheet
    Dim fechaIni As Date, fechaFin As Date
    Dim dateCol  As Long

    ' Leer filtros del PANEL
    If Not SheetExists("PANEL") Then
        MsgBox "Hoja PANEL no encontrada.", vbExclamation
        Exit Sub
    End If
    Set wsPANEL = ThisWorkbook.Worksheets("PANEL")

    If Not IsDate(wsPANEL.Range("C5").Value) Or Not IsDate(wsPANEL.Range("C6").Value) Then
        MsgBox "Las celdas C5 y C6 del PANEL deben contener fechas válidas.", vbExclamation
        Exit Sub
    End If

    fechaIni = CDate(wsPANEL.Range("C5").Value)
    fechaFin = CDate(wsPANEL.Range("C6").Value)

    If fechaIni > fechaFin Then
        MsgBox "La Fecha de Inicio no puede ser posterior a la Fecha de Fin.", vbExclamation
        Exit Sub
    End If

    If Not SheetExists("Inscripciones_Data") Then
        MsgBox "Hoja 'Inscripciones_Data' no encontrada.", vbExclamation
        Exit Sub
    End If

    Set wsInsc = ThisWorkbook.Worksheets("Inscripciones_Data")

    ' Encontrar columna "fecha_de_inscripcion" en la fila 2 (header)
    dateCol = 0
    Dim c As Range
    For Each c In wsInsc.Range("A2:Z2")
        If LCase(Trim(c.Value)) = "fecha_de_inscripcion" Then
            dateCol = c.Column
            Exit For
        End If
    Next c

    If dateCol = 0 Then
        MsgBox "No se encontró la columna 'fecha_de_inscripcion' en Inscripciones_Data.", vbExclamation
        Exit Sub
    End If

    ' Aplicar AutoFilter
    wsInsc.AutoFilterMode = False
    wsInsc.Range("A2").CurrentRegion.AutoFilter _
        Field:=dateCol, _
        Criteria1:=">=" & CDbl(fechaIni), _
        Operator:=xlAnd, _
        Criteria2:="<=" & CDbl(fechaFin)

    wsInsc.Activate
    MsgBox "Filtro aplicado en Inscripciones_Data:" & vbCrLf & _
           "  Desde: " & Format(fechaIni, "dd/mm/yyyy") & vbCrLf & _
           "  Hasta: " & Format(fechaFin, "dd/mm/yyyy"), vbInformation
End Sub

' ────────────────────────────────────────────────────────────
' 5. ExportFilteredReport
'    Copia los datos filtrados de Inscripciones_Data a una
'    nueva hoja con nombre basado en el rango de fechas.
' ────────────────────────────────────────────────────────────
Sub ExportFilteredReport()
    Dim wsInsc  As Worksheet
    Dim wsExport As Worksheet
    Dim wsPANEL As Worksheet
    Dim sheetName As String
    Dim fechaIni As Date, fechaFin As Date

    If Not SheetExists("Inscripciones_Data") Then
        MsgBox "Hoja 'Inscripciones_Data' no encontrada.", vbExclamation
        Exit Sub
    End If
    Set wsInsc = ThisWorkbook.Worksheets("Inscripciones_Data")

    If Not wsInsc.AutoFilterMode Then
        MsgBox "Aplique primero un filtro de fecha usando 'FilterByDate'.", vbExclamation
        Exit Sub
    End If

    ' Leer fechas del PANEL para el nombre
    If SheetExists("PANEL") Then
        Set wsPANEL = ThisWorkbook.Worksheets("PANEL")
        If IsDate(wsPANEL.Range("C5").Value) Then fechaIni = CDate(wsPANEL.Range("C5").Value)
        If IsDate(wsPANEL.Range("C6").Value) Then fechaFin = CDate(wsPANEL.Range("C6").Value)
    End If

    sheetName = "Reporte_" & Format(fechaIni, "ddmmyy") & "_" & Format(fechaFin, "ddmmyy")

    ' Limitar nombre a 31 caracteres (límite Excel)
    If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)

    DeleteSheetIfExists sheetName

    Set wsExport = ThisWorkbook.Worksheets.Add
    wsExport.Name = sheetName
    wsExport.Tab.Color = RGB(255, 165, 0)

    ' Copiar solo filas visibles
    Dim srcRange As Range
    Set srcRange = wsInsc.Range("A2").CurrentRegion.SpecialCells(xlCellTypeVisible)
    srcRange.Copy wsExport.Range("A1")

    ' Agregar título
    wsExport.Rows(1).Insert Shift:=xlDown
    wsExport.Range("A1").Value = "REPORTE DE INSCRIPCIONES — " & _
        Format(fechaIni, "dd/mm/yyyy") & " AL " & Format(fechaFin, "dd/mm/yyyy")
    wsExport.Range("A1").Font.Bold = True
    wsExport.Range("A1").Font.Size = 13
    wsExport.Rows(1).RowHeight = 22

    wsExport.Columns.AutoFit
    wsExport.Activate

    MsgBox "Reporte exportado a la hoja: '" & sheetName & "'", vbInformation
End Sub

' ────────────────────────────────────────────────────────────
' 6. CleanupHelperColumns — Revierte columnas insertadas
' ────────────────────────────────────────────────────────────
Sub CleanupHelperColumns()
    Dim answer As Integer
    answer = MsgBox("¿Desea eliminar las columnas calculadas de Alumnos, Cursos e Inscripciones?" & _
                    vbCrLf & "Esta acción no se puede deshacer.", _
                    vbYesNo + vbQuestion, "Confirmar")
    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    On Error Resume Next
    ' Alumnos: remover edad (K) y cursos (L)
    If SheetExists("Alumnos") Then
        With ThisWorkbook.Worksheets("Alumnos")
            If .Cells(4, 11).Value = "edad" Then .Columns("K:L").Delete
        End With
    End If

    ' Cursos: remover codigo_curso (C)
    If SheetExists("Cursos") Then
        With ThisWorkbook.Worksheets("Cursos")
            If .Cells(4, 3).Value = "codigo_curso" Then .Columns("C:C").Delete
        End With
    End If

    ' Inscripciones: remover vigencia_inicio/fin y demográficas
    If SheetExists("Inscripciones") Then
        With ThisWorkbook.Worksheets("Inscripciones")
            If .Cells(4, 3).Value = "vigencia_inicio" Then .Columns("C:D").Delete
            If .Cells(4, 7).Value = "sexo" Then .Columns("G:J").Delete
        End With
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Columnas calculadas eliminadas.", vbInformation
End Sub

' ════════════════════════════════════════════════════════════
' UTILIDADES PRIVADAS
' ════════════════════════════════════════════════════════════

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

Private Sub DeleteSheetIfExists(sheetName As String)
    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub
