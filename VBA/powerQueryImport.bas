Attribute VB_Name = "PowerQueryImportModule"
Sub ImportSpecificTables()
    Dim targetFilePath As Variant
    Dim objExcel As Object, objWB As Object, ws As Object, lo As Object
    Dim frm As New TableSelectorForm
    Dim qryName As String, mCode As String, qry As WorkbookQuery
    Dim tblName As Variant
    Dim newWS As Worksheet
    Dim sheetExists As Worksheet

    targetFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Source Workbook")
    If targetFilePath = False Then Exit Sub

    Set objExcel = CreateObject("Excel.Application")
    Set objWB = objExcel.Workbooks.Open(targetFilePath, ReadOnly:=True)

    For Each ws In objWB.Sheets
        For Each lo In ws.ListObjects
            frm.lstTables.AddItem lo.Name
        Next lo
    Next ws

    objWB.Close SaveChanges:=False
    objExcel.Quit

    frm.Show
    If frm.Cancelled Then Unload frm: Exit Sub

    For Each tblName In frm.SelectedTables
        qryName = "PQ_" & tblName

        mCode = "let" & vbCrLf & _
                "    Source = Excel.Workbook(File.Contents(""" & targetFilePath & """), null, true)," & vbCrLf & _
                "    TargetTable = Source{[Item=""" & tblName & """,Kind=""Table""]}[Data]" & vbCrLf & _
                "in" & vbCrLf & _
                "    TargetTable"

        On Error Resume Next
        Set qry = ActiveWorkbook.Queries(qryName)
        On Error GoTo 0

        Dim targetConn As WorkbookConnection
        On Error Resume Next
        Set targetConn = ActiveWorkbook.Connections("Query - " & qryName)
        On Error GoTo 0

        If Not qry Is Nothing And Not targetConn Is Nothing Then
            qry.Formula = mCode
            targetConn.Refresh
            Debug.Print tblName & " refreshed on 19/02/2026"

        Else
        
            If Not qry Is Nothing Then qry.Delete

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

            ActiveWorkbook.Queries.Add Name:=qryName, Formula:=mCode

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
