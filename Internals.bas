Attribute VB_Name = "Internals"

'This function tells if an Internal job has a valid XML available, checks for dimensions on the part, and checks the materials in the XML against the M2M BOM
Public Sub internalsMain()

    Dim days As Integer
    If IsNumeric(Worksheets(strShtTool).Range("C" & 18).value) Then
        days = Worksheets(strShtTool).Range("C" & 18).value
    Else
        Call MsgBox("The value in cell C18 must be a numeric to run the Internals list", vbOKOnly, "User Error")
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT jodrtg.fpro_id AS " & opID & ", jomast.fjobno AS " & oNum & ", jomast.fquantity AS " & q & ", jomast.fpartrev AS Rev, joitem.fpartno AS " & pnum & ", CONVERT(char(10),jodrtg.factschdst,111) AS 'exp_8', jomast.fprodcl, joitem.fdescmemo as " & memo & ", joitem.fdesc AS " & desc & ", jomast.fstatus, jodrtg.foperno" & vbCrLf & _
             "from jodrtg" & vbCrLf & _
             "INNER JOIN joitem ON joitem.fjobno = jodrtg.fjobno" & vbCrLf & _
             "INNER JOIN jomast ON jomast.fjobno = jodrtg.fjobno" & vbCrLf & _
             "INNER JOIN (SELECT jodrtg.fjobno as jobno, MIN(jodrtg.foperno) as operno" & vbCrLf & _
             "    from jodrtg" & vbCrLf & _
             "    INNER JOIN jomast ON jomast.fjobno = jodrtg.fjobno" & vbCrLf & _
             "    WHERE jodrtg.fjobno LIKE 'I%0' AND jodrtg.factschdst>={ts '1975-01-01 00:00:00'} And jodrtg.factschdst<=(GETDATE() + 15)  AND jomast.fstatus In ('OPEN', 'STARTED')" & vbCrLf & _
             "    GROUP BY jodrtg.fjobno)" & vbCrLf & _
             "    AS table2 ON table2.jobno = jodrtg.fjobno AND table2.operno = jodrtg.foperno" & vbCrLf & _
             "WHERE jodrtg.fnpct_comp<$100 AND jodrtg.factschdst>={ts '1975-01-01 00:00:00'} And jodrtg.factschdst<=(GETDATE() + 15)  AND jomast.fstatus In ('OPEN', 'STARTED') AND jomast.fjobno LIKE 'I%0'" & vbCrLf & _
             "ORDER BY jomast.fjobno"
             
    Call connQueryUpdate(connQry, strSQL)
    
    Set rst = New recordSheet
    rst.setsheet (strShtQry)          'initialize rst record set
    Set fileNames = New Collection
    
    
    Do While rst.hasNext()
    
        Dim ops As Variant, partial As Boolean, op As Variant
        ops = Array("FNEST-L", "FLASERS", "FPUNCH", "FNEST-P")
        partial = True
        For Each op In ops
            If StrComp(op, rst.field(opID), vbTextCompare) = 0 Then
                partial = False
            End If
        Next
        If partial Then
            GoTo NextLoop
        End If
        Dim jNumField As String
        jNumField = rst.field(oNum)
        
        Dim rev As String
        Dim memoField As String
        Dim revField As String
        Dim descField As String
        
        memoField = rst.field(memo)
        rev = rst.field("Rev")
        descField = rst.field(desc)
        
        If Not searchQTR(memoField, descField) Then         'check for qtr in M2M, provide error if it was found
            strError = "QTR exists in memo/desc of M2M"
            rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
            GoTo NextLoop
        End If

        strFile = getFileName(rst.field(pnum), rev)      'get the file name for this record
        If StrComp(strFile, "", vbTextCompare) = 0 Then     'check that the file was found, provide error if it wasn't found
            strError = "Filename not found"
            rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
            GoTo NextLoop
        End If
        
        fileNames.add rst.field(oNum) & "?" & strFile & "?" & rst.field(q) & "?" & rst.field(pnum)
NextLoop:
        tmpFile = ""
    Loop
    
    Dim tmpValues As Variant
    For Each tmpValues In fileNames       'loop until all relevant jobs have their cutlist processed
        Dim tmpPart As partCollection, strValues As Variant, jobno As String
        strValues = Split(tmpValues, "?")
        jobno = strValues(0)
        tmpFile = strValues(1)
        Set tmpPart = New partCollection
        orderQty = CInt(strValues(2))
        Call tmpPart.init(1, strCutPath & "\" & CStr(strValues(1)), False)      'get list of parts in this job
        
        If Not tmpPart.hasParts Then                            'check that the xml had parts in it, provide an error if not
            strError = "There were no parts in the xml"
            rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
            GoTo NextFile
        End If
        
        If StrComp(strError, "", vbTextCompare) Then
            rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
            GoTo NextFile
        End If
        
        If StrComp(dimsError, "", vbTextCompare) Then
            'rejected for missing dimensions
            rejectedCount = rejectedJob(jobno, dimsError, tmpFile, rejectedCount)
            GoTo NextFile
        End If
        
        currentReport = strShtRep
        With Worksheets(currentReport)
            .PageSetup.CenterHorizontally = True
            With .Cells
                .Clear
                .ClearFormats
                .RowHeight = 15
                .ColumnWidth = 8.43
            End With
        End With
        
        If Not tmpPart.chkInsertData("cutlist", jobno, CStr(strValues(3))) Then
            If StrComp(strError, "", vbTextCompare) Then
                rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
                GoTo NextFile
            End If
            If StrComp(pressError, "", vbTextCompare) Then
                'rejected for missing press programs
                rejectedCount = rejectedJob(jobno, pressError, tmpFile, rejectedCount)
                GoTo NextFile
            End If
        Else
            Dim matDict As Object, tmp As Variant
            Set matDict = tmpPart.compileMaterials(jobno)
            'MsgBox ("hi" + tmpPart.compileMaterials(jobno))
            Dim tmpColl As Collection
            Set tmpColl = checkMaterials(matDict, jobno)
            
            If Not tmpColl Is Nothing Then
                For Each tmp In tmpColl
                    strValues = Split(CStr(tmp), "?")
                    strError = strError & " " & CStr(strValues(0)) & " - " & CStr(strValues(1))
                Next
                'rejected for material usage
                rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
                GoTo NextFile
            End If
            
            'Call productionTraveler
            If Worksheets(strShtTool).Shapes("Check Box 6").ControlFormat.value = 1 Then
                Worksheets(currentReport).PrintOut                       'print the report
            End If
            
            currentReport = strShtShp
            Call makeReport(currentReport, tmpPart, CStr(strValues(3)), jobno)
            
            currentReport = strShtLsPt
            Call makeReport(currentReport, tmpPart, CStr(strValues(3)), jobno)
            'Call traveler
        End If
NextFile:
    Next
    
    Set fileNames = Nothing
    Set rst = Nothing
End Sub
