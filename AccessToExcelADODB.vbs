On Error Resume Next

' Create Excel application
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error creating Excel application: " & Err.Description
    WScript.Quit 1
End If

xlApp.Visible = True
xlApp.DisplayAlerts = False

' Create new workbook
Set xlWorkbook = xlApp.Workbooks.Add()
If Err.Number <> 0 Then
    WScript.Echo "Error creating Excel workbook: " & Err.Description
    xlApp.Quit
    WScript.Quit 1
End If

' Database connection string for Access
dbPath = "C:\Scripts\EngineeringTools\JobNoBurnt.accdb"
connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

' Create ADODB connection
Set conn = CreateObject("ADODB.Connection")
If Err.Number <> 0 Then
    WScript.Echo "Error creating ADODB connection: " & Err.Description
    xlWorkbook.Close False
    xlApp.Quit
    WScript.Quit 1
End If

' Open connection
conn.Open connStr
If Err.Number <> 0 Then
    WScript.Echo "Error opening database connection: " & Err.Description
    xlWorkbook.Close False
    xlApp.Quit
    WScript.Quit 1
End If

WScript.Echo "Successfully connected to Access database"

' Get list of tables
Set rs = CreateObject("ADODB.Recordset")
Set rs = conn.OpenSchema(20) ' adSchemaTables

Dim tableNames()
Dim tableCount
tableCount = 0

' Collect table names (exclude system tables)
Do While Not rs.EOF
    If rs("TABLE_TYPE") = "TABLE" And Left(rs("TABLE_NAME"), 4) <> "MSys" Then
        ReDim Preserve tableNames(tableCount)
        tableNames(tableCount) = rs("TABLE_NAME")
        tableCount = tableCount + 1
        WScript.Echo "Found table: " & rs("TABLE_NAME")
    End If
    rs.MoveNext
Loop
rs.Close

' Remove default worksheets except the first one
While xlWorkbook.Worksheets.Count > 1
    xlWorkbook.Worksheets(xlWorkbook.Worksheets.Count).Delete
Wend

' Process each table
For i = 0 To UBound(tableNames)
    tableName = tableNames(i)
    WScript.Echo "Processing table: " & tableName
    
    ' Create or use worksheet
    If i = 0 Then
        Set ws = xlWorkbook.Worksheets(1)
        ws.Name = tableName
    Else
        Set ws = xlWorkbook.Worksheets.Add()
        ws.Name = tableName
    End If
    
    ' Query the table
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & tableName & "]", conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then
        ' Write headers
        For j = 0 To rs.Fields.Count - 1
            ws.Cells(1, j + 1).Value = rs.Fields(j).Name
            ws.Cells(1, j + 1).Font.Bold = True
        Next
        
        ' Write data
        rowNum = 2
        Do While Not rs.EOF
            For j = 0 To rs.Fields.Count - 1
                If Not IsNull(rs.Fields(j).Value) Then
                    ws.Cells(rowNum, j + 1).Value = rs.Fields(j).Value
                End If
            Next
            rowNum = rowNum + 1
            rs.MoveNext
        Loop
        
        ' Auto-fit columns
        ws.Columns.AutoFit
        
        WScript.Echo "Migrated " & (rowNum - 2) & " records from " & tableName
    Else
        WScript.Echo "Table " & tableName & " is empty"
    End If
    
    rs.Close
Next

' Close connection
conn.Close

' Save workbook
saveFile = "C:\Scripts\EngineeringTools\EngineeringDatabase.xlsx"
xlWorkbook.SaveAs saveFile, 51 ' xlOpenXMLWorkbook
WScript.Echo "Database saved as: " & saveFile

' Clean up
xlApp.DisplayAlerts = True
xlWorkbook.Close
xlApp.Quit

WScript.Echo "Migration completed successfully!"