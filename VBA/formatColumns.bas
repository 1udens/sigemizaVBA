' --- Place this in a NEW module named Mod_Formatting ---
Sub ApplyAllFormats()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tblName As String

    ' Loop through every sheet in this workbook
    For Each ws In ThisWorkbook.Worksheets
        ' We only care about the sheets we created with our import prefix
        If Left(ws.Name, 7) = "Import_" Then
            
            ' Check if there is a table on the sheet
            If ws.ListObjects.Count > 0 Then
                Set lo = ws.ListObjects(1)
                tblName = Mid(ws.Name, 8) ' Strip "Import_" to get the original Table Name
                
                ' Apply Global Basic Formatting
                lo.TableStyle = "TableStyleMedium2"
                lo.Range.EntireColumn.AutoFit
                
                ' Manual Table-Specific Column Formatting
                On Error Resume Next ' Skip if a column name is not found
                Select Case tblName
                    
                    Case "SalesData"
                        lo.ListColumns("Order Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
                        lo.ListColumns("Revenue").DataBodyRange.NumberFormat = "$#,##0.00"
                        lo.ListColumns("Customer ID").DataBodyRange.NumberFormat = "0000"
                        
                    Case "Inventory"
                        lo.ListColumns("SKU").DataBodyRange.NumberFormat = "@" ' Text format
                        lo.ListColumns("Price").DataBodyRange.NumberFormat = "€#,##0.00"
                        lo.ListColumns("Stock Level").DataBodyRange.NumberFormat = "0"
                        
                    ' Add more cases as needed...
                        
                End Select
                On Error GoTo 0
            End If
        End If
    Next ws
End Sub