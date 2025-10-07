Attribute VB_Name = "DBandRS"
'create an access database connection
Public Function connectDatabase()
    Set DBCONT = CreateObject("ADODB.Connection")
    
    Dim sConn As String
    sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & strDBPath & ";" & _
            "Jet OLEDB:Engine Type=5;" & _
            "Persist Security Info=False;"
    
    On Error GoTo 2013
    DBCONT.Open sConn
    DBCONT.CursorLocation = 3
    Exit Function
2013:
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & strDBPath & ";" & _
            "Persist Security Info=False;"
                
    On Error GoTo Error
    DBCONT.Open sConn
    DBCONT.CursorLocation = 3
    Exit Function
Error:
    MsgBox "Database connection failed to initialize"
    Call closeDatabase
End Function

'close database connection
Public Function closeDatabase()
    On Error Resume Next
    Main.DBCONT.Close
    Set Main.DBCONT = Nothing
    On Error GoTo 0
End Function

'sub to update an sql query on a particular sheet by connection name
Public Sub connQueryUpdate(connName As String, strSQL As String)
        ActiveWorkbook.Connections(connName).ODBCConnection.CommandText = strSQL
        ActiveWorkbook.Connections(connName).ODBCConnection.BackgroundQuery = False
        ActiveWorkbook.Connections(connName).ODBCConnection.Refresh
End Sub

'create sql query for the burnlist
Public Function burnSQL(strThisSht As String) As String
    
    'declare variables
    Dim strSQL As String
    Dim jobno As String
    Dim release As Integer
    Dim status As String
    Dim procID As String
    Dim bgnStDate As String
    Dim endStDate As String
    Dim bgnFnDate As String
    Dim endFnDate As String
    
    status = Range("G11").Value2
    procID = Range("H11").Value2
    bgnStDate = Range("H6").Value2
    endStDate = Range("I6").Value2
    bgnFnDate = Range("H7").Value2
    endFnDate = Range("I7").Value2
    
    'manual entry error handling
    If IsNull(release) Then
        release = -1
    Else
        If IsNumeric(release) Then
            release = release
        Else
            release = -1
        End If
    End If
    
    If StrComp(status, "") Then
        status = UCase(status)
    Else
        status = "RELEASED"
    End If
    
    If StrComp(procID, "") Then
        procID = UCase(procID)
    Else
        procID = "FLASERS"
    End If
    
    'create sql query
    strSQL = "SELECT jomast.fjobno as " & oNum & ", jomast.fpartno as " & pnum & ", jomast.fpartrev as " & rev & ", " & _
    "jomast.fquantity as " & q & ", jomast.fstatus as " & stat & ", jodrtg.fpro_id as " & id & "," & _
    "jodrtg.factschdst as " & sDate & ", jodrtg.factschdfn as " & tDate & ", jodrtg.fnqty_comp as " & qComp & _
    ", joitem.fdescmemo as " & memo & vbCrLf & _
    "FROM M2MDATA01.dbo.jodrtg jodrtg, M2MDATA01.dbo.jomast jomast, M2MDATA01.dbo.joitem joitem" & vbCrLf & _
    "WHERE ((jodrtg.fjobno = jomast.fjobno) AND (jomast.fjobno = joitem.fjobno)"

    'conditional sql statements
    If StrComp(jobno, "") Then
        strSQL = strSQL & " AND ((jomast.fjobno) = '" & jobno & "')"
    End If
    
    'conditional sql statements per order number
    Dim i As Integer
    i = 3
    Do While StrComp(Worksheets(strThisSht).Range("B" & i), "", vbTextCompare)
        If i = 3 Then
            strSQL = strSQL & " AND ((((jomast.fjobno)='" & Worksheets(strThisSht).Range("B" & i) & "')"
            If StrComp(Worksheets(strThisSht).Range("E" & i), "") Then
                strSQL = strSQL & " AND ((jodrtg.fpro_id)= '" & Worksheets(strThisSht).Range("E" & i) & "')"
            Else
                strSQL = strSQL & " AND ((jodrtg.fpro_id)= '" & procID & "')"
            End If
            strSQL = strSQL & ")"
        Else
            strSQL = strSQL & " or (((jomast.fjobno)='" & Worksheets(strThisSht).Range("B" & i) & "')"
            If StrComp(Worksheets(strThisSht).Range("E" & i), "") Then
                strSQL = strSQL & " AND ((jodrtg.fpro_id)= '" & Worksheets(strThisSht).Range("E" & i) & "')"
            Else
                strSQL = strSQL & " AND ((jodrtg.fpro_id)= '" & procID & "')"
            End If
            strSQL = strSQL & ")"
        End If
        i = i + 1
    Loop
    If i > 3 Then
        strSQL = strSQL & ")"
    End If
    
    'close sql statement
    strSQL = strSQL & ");"
    'Debug.Print strSQL
    
    'return sql statement
    burnSQL = strSQL
    
End Function

'inserts the given sales order into the database if it doesn't already exist
Public Sub insertSO(so As String)
    so = Right(so, 5)
    If Not checkSO(so) Then
        Dim strSQL As String
        strSQL = "INSERT INTO [Checked Sales Orders] (SO)" & vbCrLf & _
        "VALUES ('" & so & "')"
        DBCONT.Execute strSQL
    End If
End Sub

Public Function HasHangPic(partNumber As String) As Boolean
    Dim strSQL As String
    strSQL = "SELECT [Hang] FROM [Powder and Assembly Pictures] WHERE [Part Number] = '" & partNumber & "'"

    If TypeName(rs) <> "ADODB.Recordset" Then
        
        Set rs = CreateObject("ADODB.Recordset")
        
    End If
    rs.Open strSQL, DBCONT          'execute sql query on burntlist table

    If rs.RecordCount <> 0 Then          'return results of sql query
        Dim tmp As Integer
        tmp = rs(0)
                      
        If tmp < 0 Then
            HasHangPic = True
        Else
            HasHangPic = False
        End If
    End If
End Function

Public Function HasSkidPic(partNumber As String) As Boolean

    Dim strSQL As String
    strSQL = "SELECT [Skid] FROM [Powder and Assembly Pictures] WHERE [Part Number] = '" & partNumber & "'"

    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If
    rs.Open strSQL, DBCONT          'execute sql query on burntlist table

    If rs.RecordCount <> 0 Then          'return results of sql query
        Dim tmp As Integer
        tmp = rs(0)
        If tmp <> 0 Then
            HasSkidPic = True
        Else
            HasSkidPic = False
        End If
    End If
End Function

'checks if a given sales order exists in the database
Public Function checkSO(so As String) As Boolean
    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT COUNT(SO)" & vbCrLf & _
    "FROM [Checked Sales Orders]" & vbCrLf & _
    "WHERE SO='" & so & "'"
    
    'Debug.Print (strsql)
    
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        Dim tmp As Integer
        tmp = rs(0)
        If tmp > 0 Then
            checkSO = True
        Else
            checkSO = False
        End If
    End If
    
    rs.Close
End Function

'check if a job number has an XML file associated with it already
Public Function checkJobFile(jobno As String, filename As String) As String
    
    Dim strSQL As String, tempFilename As String
    tempFilename = filename
    strSQL = "SELECT XMLFileName" & vbCrLf & _
             "FROM JobPartNumber" & vbCrLf & _
             "WHERE JobNumber ='" & jobno & "'"
             
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        If StrComp(filename, "", vbTextCompare) = 0 Then
            tempFilename = rs(0)
        ElseIf StrComp(filename, rs(0), vbTextCompare) Then
            Dim tmpYN As String
            tmpYN = MsgBox(jobno & " - Old " & rs(0) & vbCrLf & "          - New " & filename & vbCrLf & "Would you like to assign a new XML to this Job?" & vbCrLf & "(If you do not actually print this will not update)", vbQuestion + vbYesNo, "Re-Assign XML")
            If tmpYN = vbNo Then
                tempFilename = rs(0)
            End If
        End If
    End If
    
    checkJobFile = tempFilename
    
    rs.Close
End Function

'inserts or updates the xml file and job number
Public Sub insertJobFile(jobno As String, filename As String)

    If StrComp(filename, "", vbTextCompare) = 0 Then
        Exit Sub
    End If
    
    Dim tempFilename As String, strSQL As String
    tempFilename = checkJobFile(jobno, "")

    If StrComp(tempFilename, "", vbTextCompare) Then
        strSQL = "UPDATE [JobPartNumber]" & vbCrLf & _
                 "SET XMLFileName = '" & filename & "'" & vbCrLf & _
                 "WHERE JobNumber = '" & jobno & "'"
    Else
        strSQL = "INSERT INTO [JobPartNumber] (JobNumber, XMLFileName)" & vbCrLf & _
        "VALUES ('" & jobno & "', '" & filename & "')"
    End If
    
    DBCONT.Execute strSQL
    
End Sub

'looks for an xml file for an order job
Public Function getOldJobFile(jobno As String) As String
    
    Dim strSQL As String, tempFilename As String
    tempFilename = ""
    strSQL = "SELECT XMLFileName" & vbCrLf & _
             "FROM JobPartNumber" & vbCrLf & _
             "WHERE JobNumber ='" & jobno & "'"
             
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        tempFilename = rs(0)
    End If
    
    getOldJobFile = tempFilename
    
    rs.Close
End Function
