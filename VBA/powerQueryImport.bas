Sub ImportSpecificTables()
    Dim targetFilePath As Variant
    Dim objExcel As Object, objWB As Object, ws As Object, lo As Object
    Dim frm As New TableSelectorForm
    Dim qryName As String, mCode As String, qry As WorkbookQuery
    Dim tblName As Variant
    Dim newWS As Worksheet ' Declared once at the top
    Dim sheetExists As Worksheet

' 1. Select the File
    targetFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Source Workbook")
    If targetFilePath = False Then Exit Sub

    ' 2. Scan for Tables (ListObjects)
    Set objExcel = CreateObject("Excel.Application")
    Set objWB = objExcel.Workbooks.Open(targetFilePath, ReadOnly:=True)

    For Each ws In objWB.Sheets
        For Each lo In ws.ListObjects
            frm.lstTables.AddItem lo.Name
        Next lo
    Next ws

    objWB.Close SaveChanges:=False
    objExcel.Quit

' 3. Show Popup
    frm.Show
    If frm.Cancelled Then Unload frm: Exit Sub

' 4. Process Each Selected Table
    For Each tblName In frm.SelectedTables
        qryName = "PQ_" & tblName

        ' Build M Code
        mCode = "let" & vbCrLf & _
                "    Source = Excel.Workbook(File.Contents(""" & targetFilePath & """), null, true)," & vbCrLf & _
                "    TargetTable = Source{[Item=""" & tblName & """,Kind=""Table""]}[Data]" & vbCrLf & _
                "in" & vbCrLf & _
                "    TargetTable"

    ' 5. Check if Query Exists
        On Error Resume Next
        Set qry = ActiveWorkbook.Queries(qryName)
        On Error GoTo 0

        ' Try to find the connection separately to avoid Error 9
        Dim targetConn As WorkbookConnection
        On Error Resume Next
        Set targetConn = ActiveWorkbook.Connections("Query - " & qryName)
        On Error GoTo 0

        If Not qry Is Nothing And Not targetConn Is Nothing Then
            ' --- REFRESH ONLY ---
            ' Both the Query and the Connection exist
            qry.Formula = mCode
            targetConn.Refresh
            Debug.Print tblName & " refreshed on 19/02/2026"

        Else
            ' --- CREATE NEW (OR RE-LINK BROKEN) ---
            ' Either the query doesn't exist, or the connection is missing

            ' If the query exists but connection is missing, delete old query to start fresh
            If Not qry Is Nothing Then qry.Delete

            ' A. Handle the Worksheet
            On Error Resume Next
            Set sheetExists = ActiveWorkbook.Worksheets("Import_" & tblName)
            On Error GoTo 0

            If sheetExists Is Nothing Then
                Set newWS = ActiveWorkbook.Worksheets.Add
                newWS.Name = "Import_" & tblName
            Else
                Set newWS = sheetExists
                newWS.Cells.Clear
            End If

            ' B. Create the Query
            ActiveWorkbook.Queries.Add Name:=qryName, Formula:=mCode

            ' C. Load to the sheet
            With newWS.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qryName, Destination:=newWS.Range("$A$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [" & qryName & "]")
                .ListObject.DisplayName = qryName
                .Refresh BackgroundQuery:=False
            End With
        End If
    Next tblName
    
    Unload frm

    Call ApplyAllFormats
    
    MsgBox "Import/Refresh and Formatting Complete!", vbInformation
End Sub