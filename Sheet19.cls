VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Const prtySht = "Programming List"
Const proID = "Pro_ID"
Const jNum = "Job_Number"
Const pnum = "Part_Number"
Const stat = "Status"
Const rev = "Rev"
Const memo = "Memo"
Const partno = "PART NUMBER"
Const revNo = "REV #"
Const fDet = "Form #"
Const p150 = "150 Programmed"
Const p40 = "40 Programmed"
Const pr40 = "40 Robot Programmed"
Const connQry3 = "m2mdata01"


Public Sub Main()

    Call insert_Data
    Call clearColumn("B", prtySht)
    Call clearColumn("C", prtySht)
    Call clearColumn("D", prtySht)
    Call closeDatabase
    
End Sub
 
Private Sub insert_Data()
    Call connectDatabase
    
    
    Dim partName As String
    Dim pd150 As Integer
    Dim pd40 As Integer
    Dim i As Integer
    
    For i = 3 To Worksheets(prtySht).Cells(rows.count, "B").End(xlUp).row
        partName = Worksheets(prtySht).Range("B" & i).value
        If StrComp(partName, "", vbTextCompare) = 0 Then
            Exit For
        End If
        If StrComp(Worksheets(prtySht).Range("C" & i).value, "", vbTextCompare) Then
            pd150 = Int(Worksheets(prtySht).Range("C" & i).value)
        Else
            pd150 = 0
        End If
        If StrComp(Worksheets(prtySht).Range("D" & i).value, "", vbTextCompare) Then
            pd40 = Int(Worksheets(prtySht).Range("D" & i).value)
        Else
            pd40 = 0
        End If
        
        If pd150 <> 0 Or pd40 <> 0 Then
            Call insertPartNum(partName, pd150, pd40)
        End If
    Next i
    Call closeDatabase
End Sub

'check that the order number hasn't been burnt previously
'this is to prevent errors in BySoft
Private Function chkPartNum(pnum As String) As Boolean

    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT [Form Detail], [150 Programmed], [40 Programmed], [40 Robot Programmed]" & vbCrLf & _
    "FROM [Press programs]" & vbCrLf & _
    "WHERE [Form Detail]='" & pnum & "';"
    
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        chkPartNum = True
    Else
        chkPartNum = False
    End If
    
    rs.Close
End Function

'insert order number into burntlist table
Private Sub insertPartNum(pnum As String, pd150 As Integer, pd40 As Integer)
    Dim strSQL As String
    Dim insert As Boolean
    insert = False
    
    If pd150 Then
        insert = True
    End If
    If pd40 Then
        insert = True
    End If
    
    If insert Then
        If Not chkPartNum(pnum) Then
            strSQL = "SELECT [Form Detail]" & vbCrLf & _
            "FROM [Press programs]" & vbCrLf & _
            "WHERE [Form Detail]='" & pnum & "';"
                    
            If TypeName(rs) <> "ADODB.Recordset" Then
                Set rs = CreateObject("ADODB.Recordset")
            End If
        
            rs.Open strSQL, DBCONT          'execute sql query on burntlist table
            
            If rs.RecordCount <> 0 Then
                strSQL = "UPDATE [Press programs]" & vbCrLf & _
                "SET [150 Programmed] = '" & pd150 & " ', [40 Programmed] = '" & pd40 & "' " & vbCrLf & _
                "WHERE [Form Detail] = '" & pnum & "';"
            Else
                strSQL = "INSERT INTO [Press programs] ([Form Detail], [150 Programmed], [40 Programmed])" & vbCrLf & _
                "VALUES ('" & pnum & "', '" & pd150 & " ', '" & pd40 & "')" '
            End If
        End If
    End If
    
    On Error Resume Next
    DBCONT.Execute strSQL
End Sub


