Attribute VB_Name = "BOMandENGTRAV"
Option Explicit
Public cn As Object
Public Const pageHeight = 670.25
Public fstPg As Integer
Public count As Integer
Public totCellHeight As Double
Public pageCounter As Integer
Public pageStart As Integer
Public Const server = "KMI-M2M751"
Public Const db = "M2MDATA01"

'control for creating and printing Production Floor Traveler
Public Sub productionTraveler(jobno As String)
    Application.EnableCancelKey = xlDisabled
    
    Dim job As String, sqlQuery As String
    
    fstPg = 1
    job = jobno
    count = 1
    pageCounter = 0
    totCellHeight = Worksheets(prft).rows(1).Height
    pageStart = 1
    With Worksheets(prft)
        .PageSetup.CenterHorizontally = True
        With .Cells
            .Clear
            .ClearFormats
            .RowHeight = 15
            .ColumnWidth = 8.43
        End With
    End With
    
    Debug.Print (job)
    sqlQuery = "SELECT jomast.fjobno as " & oNum & ", jomast.fpartno as " & pnum & vbCrLf & _
               "FROM jomast" & vbCrLf & _
               "WHERE jomast.fpartno IN ('SETUP JOB FOR:', 'SETUP JOB') AND jomast.fjobno LIKE '" & Left(job, 5) & "%'"
    Debug.Print (sqlQuery)
    routingQuery (job)
    Call connQueryUpdate(connQry3, sqlQuery)
    Set rst2 = New recordSheet
    rst2.setsheet (bomTbl)
    Call createSMS(job, prft)
    
    If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 And Worksheets(strShtTool).Shapes("Check Box 12").ControlFormat.value = 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
        With Worksheets(prft)
            .PrintOut
        End With
    End If
End Sub

'Control for creating and printing BOM and Engineering Traveler
Public Sub traveler(jobno As String)
    Dim rs1 As Object
    Application.EnableCancelKey = xlDisabled
    Set rs1 = CreateObject("ADODB.Recordset")
    Set cn = CreateObject("ADODB.Connection")
    
    Dim job As String, sqlQuery As String

    cn.Open "Driver={SQL Server};Server=" & server & ";Database=" & db & ";"
    
    fstPg = 1
    job = jobno
    count = 1
    pageCounter = 0
    totCellHeight = Worksheets(pft).rows(1).Height
    pageStart = 1
    With Worksheets(pft)
        .PageSetup.CenterHorizontally = True
        With .Cells
            .Clear
            .ClearFormats
            .RowHeight = 15
            .ColumnWidth = 8.43
        End With
    End With
    With Worksheets(eft)
        .PageSetup.CenterHorizontally = True
        With .Cells
            .Clear
            .ClearFormats
            .RowHeight = 15
            .ColumnWidth = 8.43
        End With
    End With
    Debug.Print (job)
    sqlQuery = "SELECT jomast.fjobno as " & oNum & ", jomast.fpartno as " & pnum & vbCrLf & _
               " FROM jomast " & vbCrLf & _
               "WHERE jomast.fpartno IN ('SETUP JOB FOR:', 'SETUP JOB') AND jomast.fjobno LIKE '" & Left(job, 5) & "%'"
    Debug.Print (sqlQuery)
    routingQuery (job)
    Call connQueryUpdate(connQry3, sqlQuery)
    Set rst2 = New recordSheet
    rst2.setsheet (bomTbl)
    Dim job1 As String
    Call rst2.hasNext
    job1 = rst2.field(oNum)
    Set rst2 = Nothing
    Call createBOM(job)
    fstPg = 1
    job = rst.field(oNum)
    count = 2
    pageCounter = 0
    totCellHeight = Worksheets(eft).rows(1).Height
    pageStart = 1
    Call createSMS(job1, eft)
    
    If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 And Worksheets(strShtTool).Shapes("Check Box 13").ControlFormat.value = 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
        With Worksheets(pft)
            .PrintOut
        End With
    End If
    If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 And Worksheets(strShtTool).Shapes("Check Box 14").ControlFormat.value = 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
        With Worksheets(eft)
            .PrintOut
            Dim mysql As String
            mysql = "UPDATE jomast " & _
                    "SET fstatus = 'RELEASED', fact_rel = Convert(date, getdate()), ftrave_st = '1'," & _
                    "fhold_by = '" & Worksheets(strShtTool).Range("C" & 14).value & "T', fhold_dt = Convert(date, getdate())" & _
                    "where fjobno = '" & Left(jobno, 5) & "-0001'"
            Set rs1 = cn.Execute(mysql)
            
            Set rs1 = Nothing
        End With
    End If
    Set cn = Nothing
    
    
    'If VerifyEngineering.findFQC1RUN(jobno) Then
        
     '   Call VerifyEngineering.sullairPPAP(jobno, rev, rst.field(pnum))
        
      '  With Worksheets(sullppap)
       ' .PrintOut
        '.PrintOut
        '.PrintOut
        '.PrintOut
        'End With
    'End If
    
    
End Sub

'controller that determines what does and does not get put onto this form
Private Sub createSMS(jobNum As String, sheet As String)

    Call fthdrftr(jobNum, sheet)
    Dim description As String, descmemo As String, rst2 As recordSheet, sql As String
    Dim tempVar As Variant, nextBreak As Integer, pageCounter As Integer, tempCount As Integer
    
    sql = "SELECT CASE WHEN inmastx.fpartno IS NULL OR InMastx.fllotreqd <> 1 THEN 'C' ELSE CASE WHEN InMastx.fllotreqd = 1 AND InMastx.Flexpreqd <> 1 THEN 'D' ELSE 'F' END END as temporaryName, jomast.fjobno AS fbarcode, Jomast.fac, Jomast.fjobno as '" & oNum & "', Jomast.fstatus, Jomast.fpartno as " & pnum & ", Joitem.fdescmemo as " & memo & ", Jomast.fquantity as " & q & ", Jomast.fpartrev " & rev & ", CASE WHEN inmastx.fluseudrev IS NOT NULL AND inmastx.fluseudrev = 1 THEN jomast.fcudrev ELSE Jomast.fpartrev END as fcdisprev, Jomast.fmeasure as " & um & "," & _
          "Jomast.fddue_date as " & dueDate & ", Jomast.ftrave_dt, Jodrtg.foperno as " & opNum & ", Jodrtg.fpro_id as " & opID & ", Jodrtg.flBFLabor, Inwork.fcpro_name, Jodrtg.foperqty as " & opQty & ", Jomast.frel_dt, Jomast.fact_rel, Jodrtg.fsetuptime, Jodrtg.fuprodtime, Jodrtg.fmovetime, Jodrtg.fopermemo as " & opMemo & ", Joitem.fdesc as " & desc & ", Jodrtg.felpstime, Jomast.fsono, Jomast.fsono + Jomast.fkey AS sorelskey, Jomast.fsono + LEFT(Jomast.fkey,3) AS soitemkey, Jomast.ftype, Jomast.ftrave_st, 000000.000000 as Elapsed, joitem.fitem, jomast.fprodcl as " & pc & " " & vbCrLf & _
          "FROM jomast " & vbCrLf & _
          "INNER JOIN Joitem ON Joitem.fjobno = Jomast.fjobno " & vbCrLf & _
          "INNER JOIN jodrtg ON Jodrtg.fjobno = Jomast.fjobno " & vbCrLf & _
          "INNER JOIN Inwork ON Inwork.fcPro_Id = Jodrtg.fpro_id AND Inwork.fac = Jodrtg.fac " & vbCrLf & _
          "LEFT OUTER JOIN inmastx ON  inmastx.fpartno = jomast.fpartno AND inmastx.frev = jomast.fpartrev AND inmastx.fac = jomast.fac " & vbCrLf & _
          "WHERE  jomast.fjobno = '" & jobNum & "' " & vbCrLf & _
          "ORDER BY jomast.fjobno, jodrtg.foperno;"
          
    Call connQueryUpdate(connQry2, sql)
    
    Set rst2 = New recordSheet
    rst2.setsheet (strShtVer)
    Call formatting(sheet)
    pageCounter = 0
    
    Do While rst2.hasNext()
        If count = 5 Then
            description = Replace(rst2.field(desc), Chr(13), "")
            descmemo = Replace(rst2.field(memo), Chr(13), "")
            tempVar = Split(description & vbCrLf & descmemo, Chr(10), 60, vbTextCompare)
            Call insertDescription(tempVar, sheet)
        End If

        If insertOperation(rst2, False, sheet) Then
            totCellHeight = 0
            Call gotoNextPage(sheet)
            Call formatting(sheet)
        End If
        Call insertOperation(rst2, True, sheet)
        Debug.Print (vbTab & rst2.field(opNum) & ":" & rst2.field(opID))
    Loop
    
    If endReport(False, sheet) Then
        totCellHeight = 0
        Call gotoNextPage(sheet)
        Call formatting(sheet)
    End If
    Call endReport(True, sheet)
    Call gotoNextPage(sheet)
    
End Sub

'query that displays the routing in numerical order
Private Sub routingQuery(jobNum As String)
    Dim description As String, descmemo As String, rst2 As recordSheet, sql As String
    Dim tempVar As Variant, nextBreak As Integer, pageCounter As Integer, tempCount As Integer
    
    sql = "SELECT CASE WHEN inmastx.fpartno IS NULL OR InMastx.fllotreqd <> 1 THEN 'C' ELSE CASE WHEN InMastx.fllotreqd = 1 AND InMastx.Flexpreqd <> 1 THEN 'D' ELSE 'F' END END as temporaryName, jomast.fjobno AS fbarcode, Jomast.fac, Jomast.fjobno as '" & oNum & "', Jomast.fstatus, Jomast.fpartno as " & pnum & ", Joitem.fdescmemo as " & memo & ", Jomast.fquantity as " & q & ", Jomast.fpartrev " & rev & ", CASE WHEN inmastx.fluseudrev IS NOT NULL AND inmastx.fluseudrev = 1 THEN jomast.fcudrev ELSE Jomast.fpartrev END as fcdisprev, Jomast.fmeasure as " & um & "," & _
          "Jomast.fddue_date as " & dueDate & ", Jomast.ftrave_dt, Jodrtg.foperno as " & opNum & ", Jodrtg.fpro_id as " & opID & ", Jodrtg.flBFLabor, Inwork.fcpro_name, Jodrtg.foperqty as " & opQty & ", Jomast.frel_dt, Jomast.fact_rel, Jodrtg.fsetuptime, Jodrtg.fuprodtime, Jodrtg.fmovetime, Jodrtg.fopermemo as " & opMemo & ", Joitem.fdesc as " & desc & ", Jodrtg.felpstime, Jomast.fsono, Jomast.fsono + Jomast.fkey AS sorelskey, Jomast.fsono + LEFT(Jomast.fkey,3) AS soitemkey, Jomast.ftype, Jomast.ftrave_st, 000000.000000 as Elapsed, joitem.fitem, jomast.fprodcl as " & pc & " " & vbCrLf & _
          "FROM jomast " & vbCrLf & _
          "INNER JOIN Joitem ON Joitem.fjobno = Jomast.fjobno " & vbCrLf & _
          "INNER JOIN jodrtg ON Jodrtg.fjobno = Jomast.fjobno " & vbCrLf & _
          "INNER JOIN Inwork ON Inwork.fcPro_Id = Jodrtg.fpro_id AND Inwork.fac = Jodrtg.fac " & vbCrLf & _
          "LEFT OUTER JOIN inmastx ON  inmastx.fpartno = jomast.fpartno AND inmastx.frev = jomast.fpartrev AND inmastx.fac = jomast.fac " & vbCrLf & _
          "WHERE  jomast.fjobno = '" & jobNum & "' " & vbCrLf & _
          "ORDER BY jomast.fjobno, jodrtg.foperno;"
          
    Call connQueryUpdate(connQry2, sql)
End Sub

'creates the header and footer for the BOM
Private Sub bomhdrftr(jobno As String)
    Dim lHdr As String, cHdr As String, rHdr As String, rev As String, memoField As String, revField As String
    memoField = rst.field("Memo")
    revField = rst.field("Rev")
    rev = getRevNum(rst.field(oNum), memoField, revField, "")       'grab the rev number
    lHdr = Chr(10) & Chr(10) & Chr(10) & "&8Part Number " & rst.field(pnum) & Chr(10) & "Quantity " & rst.field(q) & "    Revision  " & rev
    Debug.Print (Len(lHdr) & "  " & lHdr)
    Worksheets(pft).PageSetup.LeftHeader = lHdr
    Worksheets(pft).PageSetup.CenterHeader = "&" & Chr(34) & "Arial" & Chr(34) & "&C&B&16Production Floor Traveler&B" & Chr(10) & "&" & Chr(34) & "BC C39 3/1 Narrow" & Chr(34) & "&28*C" & jobno & "*" & Chr(10) & "&" & Chr(34) & "Arial" & Chr(34) & "&12Job Order " & jobno
    rHdr = "&8Date: &D" & Chr(10) & "Time - &T" & Chr(10) & "Page# &P"
    Debug.Print (Len(rHdr) & "  " & rHdr)
    Worksheets(pft).PageSetup.RightHeader = rHdr
    Worksheets(pft).PageSetup.LeftFooter = "&B&10SHOP FLOOR"
End Sub

'creates the header and footer for the floor travelers
Private Sub fthdrftr(job1 As String, sheet As String)

    If StrComp("-0000", Right(job1, 5), vbTextCompare) And findFQC1RUN(job1) Then
        With Worksheets(sheet).PageSetup.CenterHeaderPicture
            .filename = headerPath
            .Height = 500
            .Width = 500
        End With
    End If
    
    Dim lHdr As String, cHdr As String, rHdr As String, rev As String, memoField As String, revField As String
    memoField = rst.field("Memo")
    revField = rst.field("Rev")
    rev = getRevNum(rst.field(oNum), memoField, revField, "")       'grab the rev number
    lHdr = Chr(10) & Chr(10) & Chr(10) & "&8Part Number " & rst.field(pnum) & Chr(10) & "Quantity " & rst.field(q) & "    Revision  " & rev
    Debug.Print (Len(lHdr) & "  " & lHdr)
    Worksheets(sheet).PageSetup.LeftHeader = lHdr
    Worksheets(sheet).PageSetup.CenterHeader = "&" & Chr(34) & "Arial" & Chr(34) & "&C&B&16Production Floor Traveler&B" & Chr(10) & "&" & Chr(34) & "BC C39 3/1 Narrow" & Chr(34) & "&28*C" & job1 & "*" & Chr(10) & "&" & Chr(34) & "Arial" & Chr(34) & "&12Job Order " & job1
    rHdr = "&8Date: &D" & Chr(10) & "Time - &T" & Chr(10) & "Page# &P"
    Debug.Print (Len(rHdr) & "  " & rHdr)
    Worksheets(sheet).PageSetup.RightHeader = rHdr
    Worksheets(sheet).PageSetup.LeftFooter = "&B&10SHOP FLOOR"
    With Worksheets(sheet)
        .Range("B" & 1).value = "Sales Order"
        .Range("C" & 1).HorizontalAlignment = xlCenter
        .Range("C" & 1).value = rst.field("fsono")
        .Range("D" & 1).value = "Loc  " & rst.field("flocation")
        .Range("E" & 1).value = "Bin  " & rst.field("fbinno")
        .Range("G" & 1).value = "Customer"
        With .Range("H" & 1 & ":J" & 1)
            .Merge
            .value = rst.field("fcompany")
        End With
        .Range("K" & 1).value = "PC  " & rst.field("fprodcl")
    End With
End Sub

'sets the formatting for the given sheet
Private Sub formatting(sheet)
    With Worksheets(sheet)
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 11
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 8
        With .Range("C" & count + 1)
            .value = "Operation"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("D" & count + 1)
            .value = "Work Center"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("E" & count)
            .value = "Operation"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("E" & count + 1)
            .value = "Quantity"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("A" & count + 2 & ":K" & count + 2)
            .Merge
            .RowHeight = 2
            .Interior.Color = RGB(128, 128, 128)
        End With
        With .Range("A" & count + 3 & ":K" & count + 3)
            .Merge
            .RowHeight = 4
        End With
        .Range("B" & count + 4).value = "Description"
        totCellHeight = totCellHeight + .rows(count & ":" & count + 3).Height
        count = 5
    End With
End Sub

'inserts the description in to the given sheet
Private Sub insertDescription(tempVar As Variant, sheet As String)
    Dim sent As Variant
    With Worksheets(sheet)
        For Each sent In tempVar
            If StrComp(sent, "", vbTextCompare) Then
                Debug.Print (sent)
                        Dim tmpStr As Variant
                        Dim tmpSent As Variant
                        Dim tmpString As String
                If count = 5 Then
                    If fstPg = 1 Then
                        totCellHeight = totCellHeight + rows(count).Height
                        count = count + 1
                        fstPg = 0
                    End If
                    If Len(sent) > 90 Then
                        tmpString = insertDesc(CStr(sent))
                        tmpStr = Split(tmpString, Chr(10))
                        For Each tmpSent In tmpStr
                                If StrComp(tmpSent, "", vbTextCompare) Then
                                If count = 5 Then
                                    If fstPg = 1 Then
                                        totCellHeight = totCellHeight + rows(count).Height
                                        count = count + 1
                                        fstPg = 0
                                    End If
                                        .Range("C" & count + pageCounter & ":K" & count + pageCounter).Merge
                                        .Range("C" & count + pageCounter) = tmpSent
                                Else
                                    .Range("B" & count + pageCounter & ":K" & count + pageCounter).Merge
                                    .Range("B" & count + pageCounter) = tmpSent
                                End If
                                totCellHeight = totCellHeight + .rows(count).Height
                                count = count + 1
                            End If
                        Next
                    Else
                        .Range("C" & count + pageCounter & ":K" & count + pageCounter).Merge
                        .Range("C" & count + pageCounter) = sent
                        totCellHeight = totCellHeight + .rows(count).Height
                        count = count + 1
                    End If
                Else
                    If Len(sent) > 90 Then
                        tmpString = insertDesc(CStr(sent))
                        tmpStr = Split(tmpString, Chr(10))
                        For Each tmpSent In tmpStr
                            If StrComp(tmpSent, "", vbTextCompare) Then
                                .Range("B" & count + pageCounter & ":K" & count + pageCounter).Merge
                                .Range("B" & count + pageCounter) = tmpSent
                                totCellHeight = totCellHeight + .rows(count).Height
                                count = count + 1
                            End If
                        Next
                    Else
                        .Range("B" & count + pageCounter & ":K" & count + pageCounter).Merge
                        .Range("B" & count + pageCounter) = sent
                        totCellHeight = totCellHeight + .rows(count).Height
                        count = count + 1
                    End If
                End If
            End If
        Next
        With .Range("A" & count + pageCounter & ":K" & count + pageCounter)
            .Merge
            .RowHeight = 4
        End With
        totCellHeight = totCellHeight + .rows(count).Height
        count = count + 1
        With .Range("A" & count + pageCounter & ":K" & count + pageCounter)
            .Merge
            .RowHeight = 2
            .Interior.Color = RGB(128, 128, 128)
        End With
        totCellHeight = totCellHeight + .rows(count).Height
        count = count + 1
        With .Range("A" & count + pageCounter & ":K" & count + pageCounter)
            .Merge
            .RowHeight = 4
        End With
        totCellHeight = totCellHeight + .rows(count).Height
        count = count + 1
    End With
End Sub

'inserts this operation from the routing on the given sheet
Private Function insertOperation(r As recordSheet, add As Boolean, sheet As String) As Boolean
    
    Dim mycount As Integer, potentheight As Double
    potentheight = 0
    mycount = count
    
    If add Then
        With Worksheets(sheet)
            With .Range("G" & mycount + pageCounter & ":K" & mycount + 3 + pageCounter).Borders
                .Color = RGB(0, 0, 0)
                .LineStyle = xlContinuous
                .weight = xlMedium
            End With
            .Range("G" & mycount + pageCounter) = "Date"
            With .Range("G" & mycount + pageCounter)
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
            End With
            .Range("G" & mycount + 1 + pageCounter) = "Clock No."
            With .Range("G" & mycount + 1 + pageCounter)
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
            End With
            .Range("G" & mycount + 2 + pageCounter) = "Good Parts"
            With .Range("G" & mycount + 2 + pageCounter)
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
            End With
            .Range("G" & mycount + 3 + pageCounter) = "Bad Parts"
            With .Range("G" & mycount + 3 + pageCounter)
                .Font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
            End With
            .Range("E" & mycount + pageCounter) = r.field(opQty)
            .Range("D" & mycount + pageCounter) = r.field(opID)
            .Range("C" & mycount + pageCounter) = r.field(opNum)
            With .Range("B" & mycount + pageCounter)
                .Font.Size = 36
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Name = "BC C39 3/1 Narrow"
            End With
            .Range("B" & mycount + pageCounter) = "*" & r.field(opNum) & "*"
            With .Range("B" & mycount + pageCounter & ":B" & mycount + 1 + pageCounter)
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            .Range("B" & mycount + 2 + pageCounter).Font.Size = 6
            .Range("B" & mycount + 2 + pageCounter) = "Operation Description Detail:"
        End With
    End If
    potentheight = totCellHeight + Worksheets(sheet).rows(mycount + pageCounter & ":" & mycount + 3 + pageCounter).Height
    mycount = mycount + 4
    
    potentheight = potentheight + insertMemo(r, mycount, add, opMemo, False, sheet)
    potentheight = potentheight + insertPadding(mycount, add, sheet)
    
    If Not add Then
        If potentheight > pageHeight Then
            insertOperation = True
        Else
            insertOperation = False
        End If
    Else
        totCellHeight = potentheight
        count = mycount
    End If
End Function

'breaks apart the description and memo fields to smaller sizes that fit onto the sheet
Private Function insertDesc(desc As Variant) As String
    Dim word As Variant, tmpVar As Variant, text As String, totalText As String
    tmpVar = Split(desc, " ")
    text = ""
    totalText = ""
    For Each word In tmpVar
        Dim length As Integer
        length = Len(text) + Len(word) + 1
        If length < 75 Then
            text = text & " " & word
        Else
            totalText = totalText & Chr(10) & text
            text = word
        End If
    Next
    insertDesc = totalText
End Function

'inserts the memo into the given sheet
Private Function insertMemo(r As recordSheet, mycount As Integer, add As Boolean, field As String, note As Boolean, sheet As String) As Double
    Dim descmemo As String, tempVar As Variant, str As Variant, potentheight As Double
    If note Then
        descmemo = "Notes:          " & r.field(field)
    Else
        descmemo = r.field(field)
    End If
    
    If StrComp(Trim(r.field(field)), "", vbTextCompare) = 0 Then
        insertMemo = 0
        Exit Function
    End If
    
    tempVar = Split(descmemo, Chr(10), 60, vbTextCompare)
    potentheight = 0
    
    With Worksheets(sheet)
        For Each str In tempVar
            If Len(CStr(str)) < 85 Then
                If add Then
                    .Range("B" & mycount + pageCounter & ":K" & mycount + pageCounter).Merge
                    .Range("B" & mycount + pageCounter) = str
                End If
                potentheight = potentheight + .rows(mycount).Height
                mycount = mycount + 1
            Else
                Dim tmpString As String, tmpStr As Variant, tmpSent As Variant
                tmpString = insertDesc(vbTab & CStr(str))
                tmpStr = Split(tmpString, Chr(10))
                For Each tmpSent In tmpStr
                    If StrComp(tmpSent, "", vbTextCompare) Then
                        If add Then
                            .Range("B" & mycount + pageCounter & ":K" & mycount + pageCounter).Merge
                            .Range("B" & mycount + pageCounter) = tmpSent
                        End If
                        totCellHeight = totCellHeight + .rows(mycount).Height
                        mycount = mycount + 1
                    End If
                Next
            End If
        Next
    End With
    If add Then
        count = mycount
    End If
    insertMemo = potentheight
End Function

'inserts padding between rows in the given sheet
Private Function insertPadding(mycount As Integer, add As Boolean, sheet As String) As Double
    Dim potentheight As Double
    potentheight = 0
    If add Then
        With Worksheets(sheet)
            With .Range("A" & mycount + pageCounter & ":K" & mycount + pageCounter)
                .Merge
                .RowHeight = 4
            End With
            With .Range("A" & mycount + 1 + pageCounter & ":K" & mycount + 1 + pageCounter)
                .Merge
                .RowHeight = 2
                .Interior.Color = RGB(128, 128, 128)
            End With
            With .Range("A" & mycount + 2 + pageCounter & ":K" & mycount + 2 + pageCounter)
                .Merge
                .RowHeight = 4
            End With
        End With
    End If
    potentheight = Worksheets(pft).rows(mycount + pageCounter & ":" & mycount + pageCounter + 2).Height
    mycount = mycount + 3
    If add Then
        count = mycount
    End If
    insertPadding = potentheight
End Function

'Puts a declaration at the end of the given report that it is the end of the report
Private Function endReport(add As Boolean, sheet As String) As Boolean
    Dim mycount As Integer, potentheight As Double
    mycount = count + 2
    If add Then
        With Worksheets(sheet)
            With .Range("A" & mycount + pageCounter & ":D" & mycount + pageCounter)
                .RowHeight = 7
                .Merge
                .Borders(xlEdgeBottom).weight = xlThin
            End With
            With .Range("A" & mycount + 1 + pageCounter & ":D" & mycount + 1 + pageCounter)
                .RowHeight = 7
                .Merge
            End With
            totCellHeight = totCellHeight + .rows(mycount + pageCounter + 1).Height
            With .Range("E" & mycount + pageCounter & ":G" & mycount + 1 + pageCounter)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Merge
            End With
            .Range("E" & mycount + pageCounter) = "End of Report"
            With .Range("H" & mycount + pageCounter & ":K" & mycount + pageCounter)
                .Merge
                .Borders(xlEdgeBottom).weight = xlThin
            End With
            With .Range("H" & mycount + 1 + pageCounter & ":K" & mycount + 1 + pageCounter)
                .Merge
            End With
        End With
    End If

    mycount = mycount + 2
    
    If add Then
        count = mycount
        totCellHeight = totCellHeight + Worksheets(sheet).rows(mycount + pageCounter & ":" & mycount + pageCounter + 1).Height
    Else
        potentheight = totCellHeight + 14
        If potentheight > pageHeight Then
            endReport = True
        Else
            endReport = False
        End If
    End If
End Function

'controller to create and print the BOM report
Private Sub createBOM(jobNum As String)
    
    Call bomhdrftr(jobNum)
    Dim rst2 As recordSheet, sql As String, rst3 As recordSheet, tempCount As Integer
    
    sql = "SELECT jodbom.fbompart, jodbom.fbomrev, jodbom.fbomdesc, jodbom.fnoperno, Jomast.fsono, jodbom.fbommeas, CASE jodbom.flextend WHEN 1 THEN jodbom.ftotqty * jomast.fquantity ELSE jodbom.factqty END as Qty_Req,CASE jodbom.fqty_iss WHEN 0 THEN jodbom.fpoqty ELSE jodbom.fqty_iss END as Qty_Iss, jodbom.fsub_job, jodbom.fbomsource, jodbom.fpoqty, jodbom.fqty_iss, CASE WHEN Inmast.flocbfdef IS NULL THEN Inmast.flocate1 ELSE Inmast.flocbfdef END AS loc, CASE WHEN Inmast.fbinbfdef IS NULL THEN Inmast.fbin1 ELSE Inmast.fbinbfdef END AS bin, jodbom.fstdmemo AS bommemo" & vbCrLf & _
          "From jomast" & vbCrLf & _
          "INNER JOIN jodbom ON jodbom.fjobno = jomast.fjobno" & vbCrLf & _
          "LEFT OUTER JOIN inmast ON inmast.fpartno = jodbom.fbompart AND inmast.frev = jodbom.fbomrev" & vbCrLf & _
          "WHERE jomast.fjobno = '" & jobNum & "' AND jodbom.fnoperno = '0'" & vbCrLf & _
          "ORDER BY jodbom.fbompart;"
    
    Call connQueryUpdate(connQry3, sql)
    
    Set rst2 = New recordSheet
    rst2.setsheet (strShtVer)
    Set rst3 = New recordSheet
    rst3.setsheet (bomTbl)
    
    Call bomFormatting
    Call insertUsedOp(rst2, 0, True)
    If rst3.emptyQuery() Then
        Do While rst3.hasNext()
            If insertBOMItem(rst3, False) Then
                totCellHeight = 0
                Call gotoNextPage(pft)
                Call bomFormatting
            End If
            Call insertBOMItem(rst3, True)
        Loop
    Else
        If insertNoMats(False) Then
            totCellHeight = 0
            Call gotoNextPage(pft)
            Call bomFormatting
        End If
        Call insertNoMats(True)
    End If
    
    tempCount = count
    If insertPadding(tempCount, False, pft) > pageHeight Then
        totCellHeight = 0
        Call gotoNextPage(pft)
        Call bomFormatting
    Else
        totCellHeight = insertPadding(count, True, pft)
    End If
        
    Do While rst2.hasNext()
        If insertUsedOp(rst2, rst2.field(opNum), False) Then
            totCellHeight = 0
            Call gotoNextPage(pft)
            Call bomFormatting
        End If
        Call insertUsedOp(rst2, rst2.field(opNum), True)
        
    sql = "SELECT jodbom.fbompart, jodbom.fbomrev, jodbom.fbomdesc, jodbom.fnoperno, Jomast.fsono, jodbom.fbommeas, CASE jodbom.flextend WHEN 1 THEN jodbom.ftotqty * jomast.fquantity ELSE jodbom.factqty END as Qty_Req,CASE jodbom.fqty_iss WHEN 0 THEN jodbom.fpoqty ELSE jodbom.fqty_iss END as Qty_Iss, jodbom.fsub_job, jodbom.fbomsource, jodbom.fpoqty, jodbom.fqty_iss, CASE WHEN Inmast.flocbfdef IS NULL THEN Inmast.flocate1 ELSE Inmast.flocbfdef END AS loc, CASE WHEN Inmast.fbinbfdef IS NULL THEN Inmast.fbin1 ELSE Inmast.fbinbfdef END AS bin, jodbom.fstdmemo AS bommemo" & vbCrLf & _
          "From jomast" & vbCrLf & _
          "INNER JOIN jodbom ON jodbom.fjobno = jomast.fjobno" & vbCrLf & _
          "LEFT OUTER JOIN inmast ON inmast.fpartno = jodbom.fbompart AND inmast.frev = jodbom.fbomrev" & vbCrLf & _
          "WHERE jomast.fjobno = '" & jobNum & "' AND jodbom.fnoperno = '" & rst2.field(opNum) & "'"
              
        Call connQueryUpdate(connQry3, sql)
        
        Set rst3 = New recordSheet
        rst3.setsheet (bomTbl)
    
        If rst3.emptyQuery() Then
            Do While rst3.hasNext()
                If StrComp(rst3.field("fbomsource"), "S", vbTextCompare) Then
                    Debug.Print ("Found one")
                End If
                If insertBOMItem(rst3, False) Then
                    totCellHeight = 0
                    Call gotoNextPage(pft)
                    Call bomFormatting
                End If
                Call insertBOMItem(rst3, True)
            Loop
        Else
            If insertNoMats(False) Then
                totCellHeight = 0
                Call gotoNextPage(pft)
                Call bomFormatting
            End If
            Call insertNoMats(True)
        End If
        
        tempCount = count
        If insertPadding(tempCount, False, pft) > pageHeight Then
            totCellHeight = 0
            Call gotoNextPage(pft)
            Call bomFormatting
        Else
            totCellHeight = insertPadding(count, True, pft)
        End If
    Loop
    
    If endReport(False, pft) Then
        totCellHeight = 0
        Call gotoNextPage(pft)
        Call formatting(pft)
    End If
    Call endReport(True, pft)
    Call gotoNextPage(pft)
    
End Sub

'controller that determines what goes onto the BOM report
Private Function insertBOMItem(rst3 As recordSheet, add As Boolean) As Boolean
    Dim mycount As Integer, potentheight As Double
    mycount = count
    potentheight = 0
    If add Then
        With Worksheets(pft)
            With .Range("A" & mycount + pageCounter & ":D" & mycount + pageCounter)
                .Merge
                .value = rst3.field("fbompart")
                .HorizontalAlignment = xlLeft
            End With
            With .Range("E" & mycount + pageCounter)
                .value = "Rev:  " & rst3.field("fbomrev")
                .HorizontalAlignment = xlLeft
            End With
            With .Range("B" & mycount + pageCounter + 1 & ":F" & mycount + pageCounter + 1)
                .Merge
                .value = Trim(rst3.field("fbomdesc"))
                .HorizontalAlignment = xlCenter
            End With
            With .Range("G" & mycount + pageCounter)
                .value = rst3.field("Qty_Req") & " " & rst3.field("fbommeas")
                .HorizontalAlignment = xlRight
            End With
            With .Range("H" & mycount + pageCounter)
                .value = rst3.field("Qty_Iss")
                .HorizontalAlignment = xlCenter
            End With
            With .Range("I" & mycount + pageCounter)
                If StrComp(rst3.field("fbomsource"), "B", vbTextCompare) Then
                    .value = rst3.field("fbomsource")
                Else
                    .value = "NS"
                End If
                .HorizontalAlignment = xlCenter
            End With
            With .Range("J" & mycount + pageCounter)
                .Font.Size = 9
                .value = rst3.field("fsub_job")
                .HorizontalAlignment = xlCenter
            End With
            With .Range("K" & mycount + pageCounter)
                .value = rst3.field("fonhand")
                .HorizontalAlignment = xlCenter
            End With
            If StrComp(rst3.field("fbomsource"), "B", vbTextCompare) Then
                If StrComp(Trim(rst3.field("loc")), "", vbTextCompare) <> 0 Then
                    With .Range("H" & mycount + pageCounter + 1 & ":I" & mycount + pageCounter + 1)
                        .Merge
                        .value = "Loc:  " & rst3.field("loc")
                    End With
                End If
                If StrComp(Trim(rst3.field("bin")), "", vbTextCompare) <> 0 Then
                    With .Range("J" & mycount + pageCounter + 1 & ":K" & mycount + pageCounter + 1)
                        .Merge
                        .value = "Bin:  " & rst3.field("bin")
                    End With
                End If
            Else
                With .Range("H" & mycount + pageCounter + 1 & ":K" & mycount + pageCounter + 1)
                    .Merge
                    .value = "Buy Item"
                End With
            End If
        End With
    End If
    mycount = mycount + 2
    potentheight = Worksheets(pft).rows(pageCounter + 1 & ":" & mycount + pageCounter).Height
    If StrComp(rst3.field("bommemo"), "", vbTextCompare) Then
        potentheight = potentheight + insertMemo(rst3, mycount, add, "bommemo", True, pft)
    End If
    If Not add Then
        If potentheight > pageHeight Then
            insertBOMItem = True
        Else
            insertBOMItem = False
        End If
    Else
        totCellHeight = potentheight
        count = mycount
    End If
End Function

'inserts the operation that this BOM item is used at
Private Function insertUsedOp(rst2 As recordSheet, num As Integer, add As Boolean) As Boolean
    Dim mycount As Integer, potentheight As Double
    potentheight = totCellHeight
    mycount = count
    With Worksheets(pft)
        If add Then
            If num = 0 Then
                With .Range("A" & mycount + pageCounter & ":C" & mycount + pageCounter)
                    .Merge
                    .value = "Used In Oper: 0"
                    .HorizontalAlignment = xlLeft
                End With
                With .Range("D" & mycount + pageCounter & ":F" & mycount + pageCounter)
                    .Merge
                    .value = "Work Center: None"
                    .HorizontalAlignment = xlLeft
                End With
                With .Range("G" & mycount + pageCounter & ":K" & mycount + pageCounter)
                    .Merge
                    .value = "None"
                    .HorizontalAlignment = xlLeft
                End With
            Else
                With .Range("A" & mycount + pageCounter & ":C" & mycount + pageCounter)
                    .Merge
                    .value = "Used In Oper: " & rst2.field(opNum)
                    .HorizontalAlignment = xlLeft
                End With
                With .Range("D" & mycount + pageCounter & ":F" & mycount + pageCounter)
                    .Merge
                    .value = "Work Center: " & rst2.field(opID)
                    .HorizontalAlignment = xlLeft
                End With
                With .Range("G" & mycount + pageCounter & ":K" & mycount + pageCounter)
                    .Merge
                    .value = rst2.field("fcpro_name")
                    .HorizontalAlignment = xlLeft
                End With
            End If
        End If
        mycount = mycount + 1
    End With
    potentheight = Worksheets(pft).rows(pageStart & ":" & mycount + pageCounter).Height
    If Not add Then
        If potentheight > pageHeight Then
            insertUsedOp = True
        Else
            insertUsedOp = False
        End If
    Else
        totCellHeight = potentheight
        count = mycount
    End If
End Function

'if there are no mats at a workcenter then insert "No Materials Assigned"
Private Function insertNoMats(add As Boolean) As Boolean
    Dim mycount As Integer, potentheight As Double
    mycount = count + 1
    potentheight = 0
    If add Then
            With Worksheets(pft).Range("A" & count + pageCounter & ":E" & count + pageCounter)
                .Merge
                .Font.Size = 8
                .value = "No Materials Assigned"
            End With
    End If
    potentheight = Worksheets(pft).rows(pageStart & ":" & mycount + pageCounter).Height
    
    If add Then
        count = mycount
        totCellHeight = potentheight
    Else
        If potentheight > pageHeight Then
            insertNoMats = True
        Else
            insertNoMats = False
        End If
    End If
    
End Function

'creates the formating for the BOM
Private Sub bomFormatting()
    Dim mycount As Integer
    mycount = count
    With Worksheets(pft)
        With .Range("A" & mycount & ":K" & mycount)
            .Merge
            .Font.Size = "25"
            .RowHeight = 26
            .value = "BOM"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        With .Range("B" & mycount + 1 & ":C" & mycount + 1)
            .Merge
            .value = "Component Part"
            .HorizontalAlignment = xlLeft
        End With
        With .Range("B" & mycount + 2 & ":F" & mycount + 2)
            .Merge
            .value = "Component Part Description"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("G" & mycount + 1)
            .value = "Quantity"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("G" & mycount + 2)
            .value = "Required"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("H" & mycount + 1)
            .value = "Quantity"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("H" & mycount + 2)
            .value = "Iss/Rec"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("I" & mycount + 2)
            .value = "Src"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("J" & mycount + 2)
            .value = "PO / JO"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("K" & mycount + 1)
            .value = "On-Hand"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("K" & mycount + 2)
            .value = "Quantity"
            .HorizontalAlignment = xlCenter
        End With
        With .Range("A" & mycount + 3 & ":K" & mycount + 3)
            .Merge
            .RowHeight = 2
            .Interior.Color = RGB(128, 128, 128)
        End With
        With .Range("A" & mycount + 4 & ":K" & mycount + 4)
            .Merge
            .RowHeight = 4
        End With
        totCellHeight = .rows(mycount & ":" & mycount + 4).Height
        count = 6
    End With
End Sub

'function that takes the cursor to the next page
Private Sub gotoNextPage(sheet As String)
    Dim index As Integer, pHeight As Double
    index = pageStart
    pHeight = 0
    
    Do While pHeight <= pageHeight
        pHeight = pHeight + Worksheets(sheet).rows(index).RowHeight
        index = index + 1
    Loop
    
    Debug.Print (index - 1 & " : " & pHeight)
    pageStart = index - 1
    pageCounter = pageStart - 1
    count = index - 1
End Sub

'function that gets the cell height of the combined height of all populated rows
Public Sub cellheights(sheet As String)
    Dim index As Integer, heights As Double
    
    index = 62
    heights = 0
    
    Do While index < 120
        heights = heights + Worksheets(sheet).rows(index).RowHeight
        index = index + 1
    Loop
    
    Debug.Print (index - 1 & " : " & heights)
End Sub
