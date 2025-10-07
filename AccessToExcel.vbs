Dim accessApp, db, tblDef, fld, rs, xl, wb, ws
Dim tableName, fieldName, recordCount, i, row

' Create Access application
Set accessApp = CreateObject("Access.Application")
accessApp.OpenCurrentDatabase "C:\Scripts\EngineeringTools\JobNoBurnt.accdb"

' Create Excel application
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add()

WScript.Echo "=== ACCESS TO EXCEL MIGRATION ==="
WScript.Echo "Connected to Access database"

' Get database reference
Set db = accessApp.CurrentDb()

' Process each table
Dim sheetIndex
sheetIndex = 1

For Each tblDef In db.TableDefs
    tableName = tblDef.Name
    
    ' Skip system tables
    If Left(tableName, 4) <> "MSys" And Left(tableName, 1) <> "~" Then
        WScript.Echo "Processing table: " & tableName
        
        ' Create or use worksheet
        If sheetIndex = 1 Then
            Set ws = wb.Worksheets(1)
            ws.Name = tableName
        Else
            Set ws = wb.Worksheets.Add()
            ws.Name = tableName
        End If
        
        ' Open recordset
        Set rs = db.OpenRecordset("SELECT * FROM [" & tableName & "]")
        
        ' Write headers
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(1, i + 1).Value = rs.Fields(i).Name
            ws.Cells(1, i + 1).Font.Bold = True
        Next
        
        ' Write data
        row = 2
        recordCount = 0
        Do While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If Not IsNull(rs.Fields(i).Value) Then
                    ws.Cells(row, i + 1).Value = rs.Fields(i).Value
                End If
            Next
            rs.MoveNext
            row = row + 1
            recordCount = recordCount + 1
        Loop
        
        rs.Close
        
        ' Auto-fit columns
        ws.Columns.AutoFit
        
        WScript.Echo "  Exported " & recordCount & " records"
        sheetIndex = sheetIndex + 1
    End If
Next

' Save Excel file
wb.SaveAs "C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"
WScript.Echo "Saved Excel database: C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"

' Clean up
wb.Close
xl.Quit
accessApp.Quit

Set ws = Nothing
Set wb = Nothing
Set xl = Nothing
Set rs = Nothing
Set db = Nothing
Set accessApp = Nothing

WScript.Echo "Migration complete!"