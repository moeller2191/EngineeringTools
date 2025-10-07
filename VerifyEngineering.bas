Attribute VB_Name = "VerifyEngineering"
'used as a check for the laser burn list tools to verify that engineering has been checked
Public Function engCompleted(ordNum As String, strThisSht As String) As Boolean

    Dim strSQL As String
    strSQL = "SELECT jomast.fjobno as " & oNum & ", jomast.fstatus as " & stat & ", jomast.fpartno as " & pnum & vbCrLf & _
    "FROM M2MData01.dbo.jomast jomast" & vbCrLf & _
    "INNER JOIN M2MData01.dbo.jodrtg AS jodrtg ON jodrtg.fjobno = jomast.fjobno" & vbCrLf & _
    "WHERE (jomast.fjobno NOT LIKE '%-0000') AND (jomast.fjobno LIKE '" & ordNum & "-%') AND (jomast.fpartno LIKE 'SETUP JOB%')" & vbCrLf & _
    "ORDER BY jomast.fjobno"
    
    Call connQueryUpdate(connQry2, strSQL)
    
    Dim location As Variant, loc As Variant, contEng As Boolean, tmp As String
    location = findColumn(CStr(ame))
    contEng = True
    
    For index = 2 To Worksheets(strThisSht).Cells(rows.count, location(1)).End(xlUp).row
        tmp = ThisWorkbook.Sheets(strShtVer).Cells(index, location(1)).value
        If StrComp(tmp, "", vbTextCompare) Then
            If Not (StrComp("RELEASED", tmp, vbTextCompare) = 0 Or StrComp("COMPLETED", tmp, vbTextCompare) = 0 Or StrComp("CLOSED", tmp, vbTextCompare) = 0) Then
                contEng = False
            End If
        End If
    Next

    engCompleted = contEng
End Function

'get status of the job based on the name
Public Function findColumn(strName As String) As Variant
    Dim names As Variant, location(2) As Variant
    names = Array("OrderNumber", "Status", "Part_Number")

    For index = 1 To Worksheets(strShtVer).Cells(1, Columns.count).End(xlToLeft).column
    
        If StrComp(names(0), ThisWorkbook.Sheets(strShtVer).Cells(1, index).value, vbTextCompare) = 0 Then
            location(0) = index
        ElseIf StrComp(names(1), ThisWorkbook.Sheets(strShtVer).Cells(1, index).value, vbTextCompare) = 0 Then
            location(1) = index
        ElseIf StrComp(names(2), ThisWorkbook.Sheets(strShtVer).Cells(1, index).value, vbTextCompare) = 0 Then
            location(2) = index
        End If
    Next
    
    findColumn = location
End Function
Public Function sullairPPAP(jobno As String, rev As String, partno As String)
     
    Dim sullppap As Worksheet
    Dim purchaseorder As String
    
            
        Worksheets("SullairPPAP").Activate
        purchaseorder = InputBox("Enter the PO Number for " + jobno)
        Range("C6") = partno
        Range("E6") = rev
        Range("G6") = purchaseorder
        
        purchaseorder = 0
        
        
End Function
'check if the job has a FQC1RUN in the routing
Public Function findFQC1RUN(job As String) As Boolean
    
    
    Dim strSQL As String, qc As recordSheet
    strSQL = "SELECT jodrtg.fpro_id AS Pro_ID, jomast.fjobno as Job_Number" & vbCrLf & _
             "FROM M2MDATA01.dbo.jomast as jomast" & vbCrLf & _
             "INNER JOIN M2MDATA01.dbo.jodrtg as jodrtg ON jodrtg.fjobno = jomast.fjobno" & vbCrLf & _
             "WHERE jomast.fjobno = '" & job & "'"

    Call connQueryUpdate(connQry2, strSQL)
    
    Set qc = New recordSheet
    qc.setsheet (strShtVer)
    
    If qc.entryExists("Pro_ID", "FQC1RUN") Then
        findFQC1RUN = True
    Else
        findFQC1RUN = False
    End If
    
End Function
Public Function findSawWeld(job As String) As Boolean
    
    
    Dim strSQL As String, weldsaw As recordSheet
    Dim weld As Boolean
    Dim saw As Boolean
    strSQL = "SELECT jodrtg.fpro_id AS Pro_ID, jomast.fjobno as Job_Number" & vbCrLf & _
             "FROM M2MDATA01.dbo.jomast as jomast" & vbCrLf & _
             "INNER JOIN M2MDATA01.dbo.jodrtg as jodrtg ON jodrtg.fjobno = jomast.fjobno" & vbCrLf & _
             "WHERE jomast.fjobno = '" & job & "'"

    Call connQueryUpdate(connQry2, strSQL)
    
    Set weldsaw = New recordSheet
    weldsaw.setsheet (strShtVer)
    
    If weldsaw.entryExists("Pro_ID", "FSAW") Then
        saw = True
    Else
        saw = False
    End If
    
    If weldsaw.entryExists("Pro_ID", "FWELDB") Then
        weld = True
    Else
        weld = False
    End If
    
    If weld = True And saw = True Then
        MsgBox (jobno & vbNewLine & "Both Weld and Saw Operations Used. Ensure Saw is after Weld")
    End If
        
    
End Function

'check if the job has a FPOHV in the routing
Public Function findFPOHV(job As String) As Double
    
    
    Dim strSQL As String, qc As recordSheet
    strSQL = "SELECT jodrtg.fpro_id AS Pro_ID, jomast.fjobno as Job_Number" & vbCrLf & _
             "FROM M2MDATA01.dbo.jomast as jomast" & vbCrLf & _
             "INNER JOIN M2MDATA01.dbo.jodrtg as jodrtg ON jodrtg.fjobno = jomast.fjobno" & vbCrLf & _
             "WHERE jomast.fjobno = '" & job & "'"

    Call connQueryUpdate(connQry2, strSQL)
    
    Set qc = New recordSheet
    qc.setsheet (strShtVer)
    
    If qc.entryExists("Pro_ID", "FPOHV") Then
        findFPOHV = True
    Else
        findFPOHV = False
    End If
    
End Function

'check that the XML materials matches M2M BOM
Public Function checkMaterials(dict As Object, jobno As String) As Variant

    Dim key As Variant, matColl As Collection, count As Integer, M2M As Double
    Dim strSQL As String, usage As recordSheet, variance As Double
    Set matColl = New Collection
    count = 0
    strSQL = "SELECT fjobno, fbompart, ftotqty" & vbCrLf & _
             "FROM jodbom" & vbCrLf & _
             "WHERE jodbom.fjobno ='" & jobno & "'"
             
    Call connQueryUpdate(connQry3, strSQL)
    Set usage = New recordSheet
    usage.setsheet (bomTbl)
    
    'do not change this number unless authorized to do so by someone with the authority to do so
    variance = 0.01
    
    For Each key In dict.keys
        M2M = usage.getInfo("ftotqty", "fbompart", CStr(key))
        If Not (((dict(key) - variance) < M2M) And ((dict(key) + variance) > M2M)) Then
            matColl.add key & "?" & dict(key)
             
            count = 1
        End If
        
       
        
    Next
    If count <> 0 Then
        Set checkMaterials = matColl
    Else
        Set checkMaterials = Nothing
    End If
End Function


