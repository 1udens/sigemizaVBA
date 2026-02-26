Attribute VB_Name = "PowerQueryImport"
Sub PowerQueryImport()
    Dim targetFilePath As Variant
    Dim objExcel As Object, objWB As Object, ws As Object, lo As Object
    Dim frm As New tableSelectorForm
    Dim qryName As String, mCode As String, qry As WorkbookQuery
    Dim tblName As Variant
    Dim newWS As Worksheet
    Dim sheetExists As Worksheet
    Dim targetConn As WorkbookConnection

    ' 1. FILE SELECTION -------------------------------------------------------
    targetFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Source Workbook")
    If targetFilePath = False Then Exit Sub

    ' 2. EXTRACT TABLE NAMES FROM SOURCE WORKBOOK -----------------------------
    Set objExcel = CreateObject("Excel.Application")
    Set objWB = objExcel.Workbooks.Open(targetFilePath, ReadOnly:=True)

    For Each ws In objWB.Sheets
        For Each lo In ws.ListObjects
            frm.lstTables.AddItem lo.Name
        Next lo
    Next ws

    objWB.Close SaveChanges:=False
    objExcel.Quit

    ' 3. USER FORM INTERACTION ------------------------------------------------
    frm.Show
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If

    ' 4. QUERY PROCESSING LOOP ------------------------------------------------
    For Each tblName In frm.SelectedTables
        qryName = "PQ_" & tblName

        ' Build Power Query M-Code
        mCode = "let" & vbCrLf & _
                "    Source = Excel.Workbook(File.Contents(""" & targetFilePath & """), null, true)," & vbCrLf & _
                "    TargetTable = Source{[Item=""" & tblName & """,Kind=""Table""]}[Data]" & vbCrLf & _
                "in" & vbCrLf & _
                "    TargetTable"

        ' Check if Query and Connection already exist
        On Error Resume Next
        Set qry = ActiveWorkbook.Queries(qryName)
        Set targetConn = ActiveWorkbook.Connections("Query - " & qryName)
        On Error GoTo 0

        If Not qry Is Nothing And Not targetConn Is Nothing Then
            ' Update existing query
            qry.Formula = mCode
            targetConn.Refresh
            Debug.Print tblName & " refreshed on 26/02/2026"

        Else
            ' Create New Query logic
            If Not qry Is Nothing Then qry.Delete

            ' Handle Worksheet Creation/Clearing
            On Error Resume Next
            Set sheetExists = ActiveWorkbook.Worksheets("Import_" & tblName)
            On Error GoTo 0

            If sheetExists Is Nothing Then
                Set newWS = ActiveWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
                newWS.Name = "Import_" & tblName
            Else
                Set newWS = sheetExists
                newWS.Cells.Clear
            End If

            ' Add New Query and ListObject Connection
            ActiveWorkbook.Queries.Add Name:=qryName, Formula:=mCode

            With newWS.ListObjects.Add(SourceType:=0, _
                Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qryName, _
                Destination:=newWS.Range("$A$1")).QueryTable
                
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [" & qryName & "]")
                .ListObject.DisplayName = qryName
                .Refresh BackgroundQuery:=False
            End With
        End If
        
        ' Reset variables for next iteration
        Set qry = Nothing
        Set targetConn = Nothing
    Next tblName

    ' 5. CLEANUP --------------------------------------------------------------
    Unload frm
    MsgBox "Import complete, proceed to manually filter dates then run FormatImport", vbInformation
End Sub

