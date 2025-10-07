Attribute VB_Name = "Main"
Option Explicit
Public DBCONT As Object         'Object that accesses M2MDATA01 database
Public rs As Object             'Object that stores a record set
Public rst As recordSheet       'object that stores a record set
Public rst2 As recordSheet
Public rst3 As recordSheet
Public matRS As Object          'object that stores a record set
Public engRS As Object
Public strError As String       'stores errors that get printed to column M on burnlist tab
Public dxfError As String       'stores missing dxf #s
Public pressError As String     'stores missing press program #s
Public dimsError As String
Public Const bomTbl = "BOMData"
Public Const pft = "PrintBOMReport"
Public Const eft = "PrintEngReport"
Public Const prft = "ProductionRouting"
Public Const sullppap = "SullairPPAP"
Public Const strNewPath = "\\kmi-solidworks22\solidworks22common\CUT LIST XML\New"          'path to folder where new cutlists
Public Const strLegPath = "\\kmi-solidworks22\solidworks22common\CUT LIST XML\Legacy"       'path to cutlist legacy folder
Public Const strCutPath = "\\kmi-solidworks22\solidworks22common\CUT LIST XML"              'path to cutlist xml folder
Public Const strArchivePath = "\\kmi-solidworks22\solidworks22common\CUT LIST XML"          'path to full XML archive for indexing
Public Const strDBPath = "\\kmi-solidworks22\solidworks22common\CUT LIST XML\xmlCutlist\JobNoBurnt.accdb"       'path to reference access database
Public Const headerPath = "\\kmi-exch2013\m$\KMISHARE\ENGINEERING\Solid Works Integration\ToolAndTemplates\FAI.png"
Public Const strShtQry = "query"                'reference to query sheet
Public Const strShtMat = "Materials"            'reference to materials sheet
Public Const strShtRep = "PrintReportForm"      'reference to printreportform sheet
Public Const strShtTool = "TravelerTool"        'reference to Engineering Traveler sheet
Public Const strByTool = "BySoft"               'reference to BySoft sheet
Public Const strShtSigTool = "SigmaNestBurnlistTool"    'reference to Sigmanest sheet
Public Const strShtVer = "engVerify"            'reference to engVerify sheet
Public Const strShtShp = "ShipLoose"            'reference to ShipLoose sheet
Public Const strShtLsPt = "LoosePart"           'reference to LoosePart sheet
Public Const evrythgSht = "Everything in BOM"   'reference to Everything in BOM sheet
Public Const xmlPath = "//xml/transactions/transaction/document/configuration/references"   'initial path to follow when opening an XML file
Public Const connQry = "m2mdata02"                                              'ODBCconnection name for sheet query
Public Const connQry2 = "m2mdata03"                                             'ODBCconnection name for engVerify
Public Const connQry3 = "m2mdata05"                                             'ODBCconnection name for BOMdata
Public Const oNum = "OrderNumber"                                                           'names of columns from the SQL query
Public Const stat = "Status"
Public Const id = "ProID"
Public Const tDate = "TargetDate"
Public Const sDate = "StartDate"
Public Const qComp = "QtyComp"
Public Const dueDate = "Due_Date"
Public Const pnum = "Part_Number"
Public Const rev = "Rev"
Public Const q = "Quantity"
Public Const memo = "Memo"
Public Const desc = "Description"
Public Const um = "Unit_Measure"
Public Const loc = "Location"
Public Const bin = "Bin"
Public Const pc = "Product_Class"
Public Const opNum = "Operation_Number"
Public Const opID = "Operation_ID"
Public Const opMemo = "Operation_Memo"
Public Const opQty = "Operation_Quantity"
Public tmpFile As String                'temp file that stores part names of files that won't load
Public Const dxfInch = "\\kmi-exch2013\cadtemp\ENGINEERING\DXF-INCH"
Public reportTitle As String
Public currentReport As String
Public globDict As Object
Public orderQty As Integer
Public rejectedCount As Integer
Public fileNames As Collection
Public desc2 As String
Public desc3 As String
Private Sub Workbook_Open()

CheckBox9.value = True

Call createReport_Btn_Click

End Sub
Public Sub createReport_Btn_Click()
    


    If Worksheets(strShtTool).Shapes("Check Box 14").ControlFormat.value = 1 Then       'check that the user`s initials are present when "Print Engineering Traveler" is checked
        If StrComp(Worksheets(strShtTool).Range("C" & 14), "", vbTextCompare) = 0 Then
            MsgBox "Your initials must be present to print and release the Engineering Traveler."
            Exit Sub
        End If
    End If

    If lookForBadPartNames() Then           'check for bad xml file names
        Exit Sub
    End If
    
    Call mptLgcyFldr                        'empties the files out of the "New" folder moving them into the active folder "CUT LIST XML" and moving old revisions from the "CUT LIST XML" folder into the "Legacy" Folder

    'declare variables
    Dim strFile As String
    Dim rejected As Boolean
    Dim dateTime As String
    Dim sucessCount As Integer
    Dim printFail As Integer
    Dim i As Integer
    Dim dxfCount As Integer
    Dim jobDict As Object
    
    'clears the data from each column given below
    Call clearColumn("D", strShtTool)
    Call clearColumn("E", strShtTool)
    Call clearColumn("F", strShtTool)
    Call clearColumn("G", strShtTool)
    Call clearColumn("H", strShtTool)
    Call clearColumn("I", strShtTool)

    'initialize variables
    dateTime = Now()                            'initialize the date variable
    dateTime = Replace(dateTime, "/", "-")      'format the time so it can be used as a filename
    dateTime = Replace(dateTime, ":", "_")      'format the time so it can be used as a filename
    rejected = False
    rejectedCount = 3
    printFail = 2
    sucessCount = 3
    i = 3
    dxfCount = 2
    desc2 = ""
    desc3 = ""
    Set jobDict = CreateObject("Scripting.dictionary")
    tmpFile = ""
    
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If
    
    Call connectDatabase                        'create database connection
    
    If Worksheets(strShtTool).Shapes("Check Box 9").ControlFormat.value = 1 Then        'check if internals should be verified
        Call internalsMain              'perform check on internals
    End If
    
    Call connQueryUpdate(connQry, getSQL(strShtTool))     'execute the sql query and refresh on query sheet
    Set rst = New recordSheet
    rst.setsheet (strShtQry)          'initialize rst record set
    Set fileNames = New Collection
    
    Do While rst.hasNext()              'loops over all rows on the "query" sheet
    
        Dim jNumField As String
        jNumField = rst.field(oNum)     'extract the job number from the current row
        
        Dim rev As String
        Dim memoField As String
        Dim revField As String
        Dim descField As String
        Dim salesOrder As String
        
        'initialize variables with data from the "query" sheet
        memoField = rst.field("Memo")
        revField = rst.field("Rev")
        descField = rst.field("Description")
        salesOrder = rst.field("fsono")
        
        If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
            If StrComp(Left(jNumField, 1), "I", vbTextCompare) <> 0 And Not checkSO(Right(salesOrder, 5)) Then      'verify that sales order has been checked
                strError = "Sales Order still needs checked"
                tmpFile = Right(salesOrder, 5)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
            End If
        
            If Not searchQTR(memoField, descField) Then         'check for qtr in M2M, provide error if it was found
                strError = "QTR exists in memo/desc of M2M - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
            End If
        End If
        
        
        If Not searchSIMTO(memoField, descField) Then         'check for Similar in M2M, provide error if it was found
                strError = "Similar To Job - Please Review - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchPrintReq(memoField, descField) Then         'check for `PRINTREQ in M2M, provide error if it was found
                strError = "Print Request Pending - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchSDR(memoField, descField) Then         'check for SDR in M2M, provide error if it was found
                strError = "SDR Pending Review - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchQC(memoField, descField) Then         'check for `QCgeneral in M2M, provide error if it was found
                strError = "Waiting on Engineering QC - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                
                GoTo NextLoop
        End If
                
        If Not searchQCBM(memoField, descField) Then         'check for `QCBM in M2M, provide error if it was found
                strError = "Waiting on Brian Moeller to QC - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                
                GoTo NextLoop
                
        End If
        
        If Not searchQCRH(memoField, descField) Then         'check for `QCRH in M2M, provide error if it was found
                strError = "Waiting on Aaron Foster to QC - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                
                GoTo NextLoop
        End If
        
        If Not searchXML(memoField, descField) Then         'check for `xml in M2M, provide error if it was found
                strError = "Waiting on XML - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchReady(memoField, descField) Then         'check for `ready in M2M, provide error if it was found
                strError = "Ready for Engineering - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchDesignReview(memoField, descField) Then         'check for `Training in M2M, provide error if it was found
                strError = "Design Review Pending - " + rst.field(pnum)
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        If Not searchPendingCancel(memoField, descField) Then         'check for `Training in M2M, provide error if it was found
                strError = "Cancel Pending"
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
        End If
        
        'CAUSING ISSUES FOR SOME REASON BMM 10/22/19
        
       'If Not HasHangPic(rst.field(pnum)) = True And Worksheets(strShtTool).Shapes("Check Box 19").ControlFormat.value <> 1 Then
       '       strError = rst.field(pnum) + "   No HANG Pic - Please Review"
       '         rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
        '        GoTo NextLoop
        'End If
        'If Not HasSkidPic(rst.field(pnum)) = True And Worksheets(strShtTool).Shapes("Check Box 19").ControlFormat.value <> 1 Then
        '       strError = "No SKID Pic - Please Review"
        '        rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
        '        GoTo NextLoop
        'End If
        
        Dim ops As Variant, partial As Boolean, op As Variant
        
        rev = getRevNum(rst.field(oNum), memoField, revField, "")       'grab the rev number
        If StrComp(rev, "", vbTextCompare) = 0 Then                     'check the rev was found, provide error if it wasn't found
        
            ops = Array("FNEST-L", "FLASERS", "FPUNCH", "FNEST-P")
            partial = True
            For Each op In ops                                                  'check that one of the Workcenters listed above is the first Workcenter in the routing
                If StrComp(op, rst.field(opID), vbTextCompare) = 0 Then
                    partial = False
                End If
            Next
            
            If partial Then                                                     'if this job`s routing does not begin with one of the Workcenters listed above, continue with the code inside the if statement
                If Worksheets(strShtTool).Shapes("Check Box 12").ControlFormat.value = 1 Then
                    Call productionTraveler(jNumField)                                         'print the Production Floor Traveler
                End If
                If Worksheets(strShtTool).Shapes("Check Box 13").ControlFormat.value = 1 Or Worksheets(strShtTool).Shapes("Check Box 14").ControlFormat.value = 1 Then
                    Call traveler(jNumField)                                                   'print the BOM and Engineering Traveler
                End If
                Worksheets(strShtTool).Range("H" & sucessCount) = rst.field(oNum)
                sucessCount = sucessCount + 1
                strError = "No Laser or Punch Operation"
                tmpFile = "No cutlist needed"
            ElseIf StrComp(strError, "", vbTextCompare) = 0 Then               'error message
                strError = "Missing Revision, make sure it's `REV### followed by a space (` is the grave accent in the upper lefthand corner of the keyboard)"
            End If
            rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
            GoTo NextLoop
        End If
        
        strFile = getFileName(rst.field(pnum), rev)                     'get the file name for this record
        If StrComp(strFile, "", vbTextCompare) = 0 Then                 'check that the file was found, provide error if it wasn't found
        
            ops = Array("FNEST-L", "FLASERS", "FPUNCH", "FNEST-P")
            partial = True
            For Each op In ops                                                  'check that one of the Workcenters listed above is the first Workcenter in the routing
                If StrComp(op, rst.field(opID), vbTextCompare) = 0 Then
                    partial = False
                End If
            Next
            
            If partial Then                                                     'if this job`s routing does not begin with one of the Workcenters listed above, continue with the code inside the if statement
                If Worksheets(strShtTool).Shapes("Check Box 12").ControlFormat.value = 1 Then       'check if the Production Floor Traveler should be printed
                    Call productionTraveler(jNumField)
                End If
                If Worksheets(strShtTool).Shapes("Check Box 13").ControlFormat.value = 1 Or Worksheets(strShtTool).Shapes("Check Box 14").ControlFormat.value = 1 Then 'check if the BOM or Engineering Traveler should be printed
                    Call traveler(jNumField)
                End If
                Worksheets(strShtTool).Range("H" & sucessCount) = rst.field(oNum)
                sucessCount = sucessCount + 1
                strError = "No Laser or Punch Operation"
                tmpFile = "No cutlist needed"
            ElseIf StrComp(strError, "", vbTextCompare) = 0 Then               'error message
                strError = "Filename not found"
                tmpFile = rst.field(pnum) & "_REV" & Right(rev, 2) & "*.xml"
            End If
            rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
            GoTo NextLoop
        End If
        If StrComp(rst.field("fstatus"), "CLOSED", vbTextCompare) = 0 Or StrComp(rst.field("fstatus"), "COMPLETED", vbTextCompare) = 0 Then         'checks if a job is COMPLETED or CLOSED in M2M
            strFile = getOldJobFile(rst.field(oNum))                                                                'checks if the job has had an XML file associated with it already
            If StrComp(strFile, "", vbTextCompare) = 0 Then
                strError = "There is no filename associated with this job"
                tmpFile = rst.field(pnum) & "_REV" & rev & "*.xml"
                rejectedCount = rejectedJob(rst.field(oNum), strError, tmpFile, rejectedCount)
                GoTo NextLoop
            End If
        Else
            strFile = checkJobFile(rst.field(oNum), strFile)
        End If
        
        
        fileNames.add rst.field(oNum) & "?" & strFile & "?" & rst.field(q) & "?" & rst.field(pnum)                  'insert required information into a collection for furthur processing after all desired jobs, that can be, have been inserted into this collection
NextLoop:
        tmpFile = ""
    Loop
    
    Dim tmpValues As Variant
    For Each tmpValues In fileNames                                                                                 'loop all jobs in the collection have been processed
        Dim tmpPart As partCollection, strValues As Variant, jobno As String
        strValues = Split(tmpValues, "?")
        jobno = strValues(0)
        Call rst.goToRow(jobno, oNum)
        tmpFile = strValues(1)
        Set tmpPart = New partCollection
        orderQty = CInt(strValues(2))
        Dim tempFilepath As String
        tempFilepath = Dir$(strCutPath & "\" & CStr(strValues(1)))                                                  'check for the part number in the "CUT LIST XML" folder first
        If StrComp(tempFilepath, "", vbTextCompare) = 0 Then
            tempFilepath = strLegPath & "\" & Dir$(strLegPath & "\" & CStr(strValues(1)))
        Else
            tempFilepath = strCutPath & "\" & tempFilepath
        End If
        Call tmpPart.init(1, tempFilepath, False)                                                                   'get collection of parts in this job
        
        If Not tmpPart.hasParts Then                                                                                'check that the xml had parts in it, provide an error if not
            strError = "There were no parts in the xml"
            rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
            GoTo NextFile
        End If
        
        If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                              'check that the "Test lists" check box is checked, if it is, skip the following code
            If StrComp(strError, "", vbTextCompare) Then
                rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
                GoTo NextFile
            End If
            
            If StrComp(dimsError, "", vbTextCompare) Then                                                           'check that the dimensions were all found
                rejectedCount = rejectedJob(jobno, dimsError, tmpFile, rejectedCount)
                GoTo NextFile
            End If
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
        
        If Not tmpPart.chkInsertData("cutlist", jobno, CStr(strValues(3))) Then                                     'create the cutlist if the data is valid
            If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
                If StrComp(strError, "", vbTextCompare) Then
                    rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
                    GoTo NextFile
                ElseIf StrComp(pressError, "", vbTextCompare) Then
                    rejectedCount = rejectedJob(jobno, pressError, tmpFile, rejectedCount)
                    GoTo NextFile
                Else
                    rejectedCount = rejectedJob(jobno, "No Items for Cutlist or some other unhandled error", tmpFile, rejectedCount)
                    GoTo NextFile
                End If
            End If
        Else
            
            
            'compile and check the materials in the XML against the BOM
            Dim matDict As Object, tmp As Variant
            Set matDict = tmpPart.compileMaterials(jobno)
            Dim tmpColl As Collection
            Set tmpColl = checkMaterials(matDict, jobno)
            
            If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
                If Not tmpColl Is Nothing Then                                                                      'if there is something in the tmpColl variable after performing the materials check then print out the appropriate error
                    For Each tmp In tmpColl
                        strValues = Split(CStr(tmp), "?")
                        strError = strError & " " & CStr(strValues(0)) & " - " & CStr(strValues(1))
                    Next
                    rejectedCount = rejectedJob(jobno, strError, tmpFile, rejectedCount)
                    GoTo NextFile
                End If

                Worksheets(strShtTool).Range("H" & sucessCount) = jobno    'add to list of successfully added orders
                sucessCount = sucessCount + 1
                If StrComp(dxfError, "", vbTextCompare) Then
                    Worksheets(strShtTool).Range("I" & dxfCount + 1) = dxfError
                    dxfError = ""
                    dxfCount = dxfCount + 1
                End If
                
                If Worksheets(strShtTool).Shapes("Check Box 12").ControlFormat.value = 1 Then
                    Call productionTraveler(jobno)                                                                             'create the Production Floor Traveler
                End If
                
                If Worksheets(strShtTool).Shapes("Check Box 6").ControlFormat.value = 1 Then                        'check if the "Print Cutlist Reports" box is checked, if it is print out the cutlist reports
                    Worksheets(currentReport).PrintOut                       'print the report
                    Call insertJobFile(jobno, tmpFile)                                                              'Update/Insert the job and XML into the database for future reference if needed
                End If
            End If
            
            'START check bom for bogus
                                
                Dim Cell As Range
                Worksheets("BOMData").Activate
                Columns("A:A").Select
    
                Set Cell = Selection.Find(What:="HDW10-2424001000000000000", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HDW10-2424001000000000000 on " + jobno & vbNewLine & "This number Should generate correct number with Zinc yellow")
                End If
                
                Set Cell = Selection.Find(What:="HDW14-0002004000000000000", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HDW14-0002004000000000000 on " & jobno & vbNewLine & "This number Should generate correct number with Zinc yellow")
                End If
                
                Set Cell = Selection.Find(What:="HHW70-000009035", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HHW70-000009035 on " & jobno & vbNewLine & "This number should be HHW70-000009037")
                End If
                
                Set Cell = Selection.Find(What:="HDW14-2050021000.75000000", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HDW14-2050021000.75000000 on " & jobno & vbNewLine & "This number Should generate correct number with Zinc yellow")
                End If
                
                Set Cell = Selection.Find(What:="HDW14-2050021001.00000000", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HDW14-2050021001.00000000 on " & jobno & vbNewLine & "This number Should generate correct number with Zinc yellow")
                End If
                
                Set Cell = Selection.Find(What:="HDW16-1824021000000000000", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox ("Check the BOM for HDW16-1824021000000000000 on " & jobno & vbNewLine & "This number Should generate correct number with Zinc yellow")
                End If
                
                Set Cell = Selection.Find(What:="STRUCC", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox (jobno + " has structural channel. Make sure ft in the bom is divisible by 20 ft and notify purchasing that you have released the job.")
                End If
                
                Set Cell = Selection.Find(What:="STRUCT", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox (jobno + " has structural tube. Make sure ft in the bom is divisible by 20 ft and notify purchasing that you have released the job.")
                End If
                
                Set Cell = Selection.Find(What:="STRUCA", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox (jobno + " has structural angle. Make sure ft in the bom is divisible by 20 ft and notify purchasing that you have released the job.")
                End If
                
                Set Cell = Selection.Find(What:="DHWARE-DY", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox (jobno + " has paino hinge. Make sure ft in the bom is divisible by 8 ft and notify purchasing that you have released the job.")
                End If
                
                Set Cell = Selection.Find(What:="STRUT-", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                Else
                MsgBox (jobno + " has a strut item. **If the item is saw cut strut** Make sure ft in the bom is divisible by 10 ft and notify purchasing that you have released the job.")
                End If
                
                Set Cell = Selection.Find(What:="SKID-", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
                If Cell Is Nothing Then
                MsgBox (jobno + " is missing skid")
                Else
                End If
                
                'Set Cell = Selection.Find(What:="(ENTER BOM ITEM HERE)", After:=ActiveCell, LookIn:=xlFormulas, _
                'LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                'MatchCase:=False, SearchFormat:=False)
                'If Cell Is Nothing Then
                'Else
                'MsgBox ("Check the BOM for (ENTER BOM ITEM HERE) on " & jobno & vbNewLine & "(ENTER ADDITIONAL MESSAGE HERE")
                'End If
                
                
                Sheets("TravelerTool").Activate
                Range("A1").Select

                               
            'END bogus checkj
            
            'start sullairPPAP check
            'If VerifyEngineering.findFQC1RUN(jobno) = True Then
            
                'Call VerifyEngineering.sullairPPAP(jobno, rev, rst.field(pnum))
            
            'End If
            'end SullairPPAP check
            Call findSawWeld(jobno)
            
            currentReport = strShtShp
            Call makeReport(currentReport, tmpPart, CStr(strValues(3)), jobno)                                             'create the Ship Loose List
            
            currentReport = strShtLsPt
            Call makeReport(currentReport, tmpPart, CStr(strValues(3)), jobno)                                             'Create the Loose Parts List
            If Worksheets(strShtTool).Shapes("Check Box 13").ControlFormat.value = 1 Or Worksheets(strShtTool).Shapes("Check Box 14").ControlFormat.value = 1 Then
                Call traveler(jobno)                                                                                           'create the BOM and Engineering Traveler
            End If
        End If
NextFile:
    Next
    
    Call closeDatabase                      'close database and record sets
End Sub

Public Sub makeReport(current As String, tmpPart As partCollection, partno As String, jobno As String)
    With Worksheets(current)
        .Cells.Clear                              'reset sheet to default
        .Cells.ClearFormats
        .Cells.RowHeight = 15
        .Cells.ColumnWidth = 8.43
        
        If StrComp("ShipLoose", current, vbTextCompare) = 0 Then
            reportTitle = "PARTS SHIPPED LOOSE FOR JOB " & jobno
            Call genericFormat_ShipReport(partno)           'set generic format of report
        Else
            reportTitle = "DMG PARTS LIST FOR JOB " & jobno
            Call genericFormat_LooseReport(partno)           'set generic format of report
        End If
        
        If insertData(currentReport, tmpPart) Then
            If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
                If Worksheets(strShtTool).Shapes("Check Box 6").ControlFormat.value = 1 Then                        'check that the form should be printed
                    Worksheets(currentReport).Cells.PrintOut
                    If StrComp("loosepart", current, vbTextCompare) Then                                            'Print an extra Loose Part List
                        Worksheets(currentReport).Cells.PrintOut
                    End If
                End If
            End If
        End If
    End With
End Sub

'check that the order number hasn't been burnt previously
'this is to prevent errors in BySoft
Public Function chkOrderNo(ordNo As String) As Boolean

    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT COUNT(OrderNumber)" & vbCrLf & _
    "FROM Burntlist" & vbCrLf & _
    "WHERE OrderNumber='" & ordNo & "';"

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        Dim tmp As Integer
        tmp = rs(0)
        If tmp > 0 Then
            chkOrderNo = True
        Else
            chkOrderNo = False
        End If
    End If
    
    rs.Close
End Function

'check if the file is already in the Legacy database
'if so, return empty string
'else return the filename
Public Function isLegacy(strFile As String) As String

    If StrComp(strFile, "", vbTextCompare) = 0 Then     'check that the strFile string is not empty string
        isLegacy = ""                                   'return if it is
        Exit Function
    End If

    Dim strSQL As String            'create sql query
    strSQL = "SELECT COUNT(fileName) AS numFiles" & vbCrLf & _
    "FROM Legacy" & vbCrLf & _
    "WHERE fileName='" & strFile & "'"
    
    rs.Open strSQL, DBCONT          'execute sql query on the Legacy table
    
    If rs.RecordCount > 0 Then              'return results
        If rs(0) > 0 Then
            Call deleteFile(strFile, strCutPath)
            isLegacy = ""
        Else
            isLegacy = strFile
        End If
    End If
    
    rs.Close
End Function

'insert order number into burntlist table
Public Sub insertOrderNo(ordNo As String)
    Dim strSQL As String
    strSQL = "INSERT INTO Burntlist (OrderNumber)" & vbCrLf & _
    "VALUES ('" & ordNo & "')"
    DBCONT.Execute strSQL
End Sub

'create sql query for the burnlist
Public Function getSQL(strThisSht) As String
    
    'declare variables
    Dim strSQL As String, strSQL2 As String, strSQL3 As String, strList As String
    Dim jobno As String
    Dim release As Integer
    Dim status As String
    
    status = Range("G11").Value2
    
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
    
    'create sql query
    strSQL = "SELECT table2." & oNum & ", table2." & pnum & ", table2." & opNum & ", jodrtg.fpro_id as " & opID & ", CONVERT(char(10),jodrtg.factschdst,111) as " & dueDate & ", table2." & q & ", table2." & rev & ", table2.fsono, OnHand.flocation, OnHand.fbinno, table2.fprodcl, table2.fcompany, joitem.fdescmemo as " & memo & ", joitem.fdesc AS " & desc & ", table2.fstatus" & vbCrLf & _
             "FROM jodrtg" & vbCrLf & _
             "INNER JOIN (SELECT Jomast.fjobno as " & oNum & ",CONVERT(char(10), jomast.fddue_date,111) as " & dueDate & ", jomast.fpartno as " & pnum & ", MIN(jodrtg.foperno) as " & opNum & ", jomast.fquantity as " & q & ", jomast.fpartrev as " & rev & ", jomast.fsono, jomast.fprodcl, jomast.fcompany, jomast.fstatus" & vbCrLf & _
             "FROM jomast" & vbCrLf & _
             "INNER JOIN jodrtg ON jomast.fjobno = jodrtg.fjobno" & vbCrLf & _
             "WHERE  jomast.fjobno IN ("
             
    strSQL2 = ")" & vbCrLf & _
              "GROUP BY jomast.fjobno, jomast.fpartno, jomast.fddue_date, jomast.fquantity, jomast.fpartrev, jomast.fsono, jomast.fprodcl, jomast.fcompany, jomast.fstatus)" & vbCrLf & _
              "AS table2" & vbCrLf & _
              "ON (table2." & oNum & " = jodrtg.fjobno AND table2." & opNum & " = jodrtg.foperno)" & vbCrLf & _
              "INNER JOIN joitem ON joitem.fjobno = table2." & oNum & vbCrLf & _
              "FULL OUTER JOIN (SELECT Inonhd.*" & vbCrLf & _
              "FROM Inonhd" & vbCrLf & _
              "INNER JOIN Inmastx ON Inmastx.fpartno = Inonhd.fpartno AND Inmastx.frev = Inonhd.fpartrev AND Inmastx.flocbfdef = Inonhd.flocation" & vbCrLf & _
              "WHERE Inonhd.flocation <> '10')" & vbCrLf & _
              "AS OnHand ON OnHand.fpartno = table2." & pnum & " AND OnHand.fpartrev = table2." & rev & vbCrLf & _
              "WHERE table2." & oNum & " IN ("
             
    strSQL3 = ")" & vbCrLf & _
              "ORDER BY " & oNum & ", " & dueDate

    'conditional sql statements
    If StrComp(jobno, "") Then
        strSQL = strSQL & " AND ((jomast.fjobno) = '" & jobno & "')"
    End If
    
    'conditional sql statements per order number
    Dim i As Integer, added As Integer
    added = 0
    For i = 3 To Worksheets(strThisSht).Cells(rows.count, "B").End(xlUp).row
        Dim jobSearch As String
        jobSearch = Worksheets(strThisSht).Range("B" & i)
        
        If StrComp(jobSearch, "", vbTextCompare) Then
            added = added + 1
            jobSearch = Left(jobSearch, 10) '& "-0000"
            If added = 1 Then
                strList = "'" & jobSearch & "'"
            Else
                strList = strList & ", '" & jobSearch & "'"
            End If
        End If
    Next
    
    strSQL = strSQL + strList + strSQL2 + strList + strSQL3
    
    
    'Debug.Print (strSQL)
    'return sql statement
    If added <> 0 Then
        getSQL = strSQL
    Else
        getSQL = ""
        Debug.Print ("exit macro")
        End
    End If
    
End Function

'look for the text "qtr" in the memo and/or description fields of M2M
Public Function searchQTR(memo As String, desc As String) As Boolean
    If InStr(1, memo, "qtr", vbTextCompare) Then
        searchQTR = False
    Else
        If InStr(1, desc, "qtr", vbTextCompare) Then
            searchQTR = False
        Else
            searchQTR = True
        End If
    End If
End Function

'look for the text "similar" in the memo and/or description fields of M2M
Public Function searchSIMTO(memo As String, desc As String) As Boolean
    If InStr(1, memo, "similar", vbTextCompare) Then
        searchSIMTO = False
    Else
        If InStr(1, desc, "similarx", vbTextCompare) Then
            searchSIMTO = False
        Else
            searchSIMTO = True
        End If
    End If
End Function

'look for the text "`PrintReq" in the memo and/or description fields of M2M
Public Function searchPrintReq(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`PrintReq", vbTextCompare) Then
        searchPrintReq = False
    Else
        If InStr(1, desc, "'PrintReq", vbTextCompare) Then
            searchPrintReq = False
        Else
            searchPrintReq = True
        End If
    End If
End Function

'look for the text "sdr" in the memo and/or description fields of M2M
Public Function searchSDR(memo As String, desc As String) As Boolean
    If InStr(1, memo, "sdr", vbTextCompare) Then
        searchSDR = False
    Else
        If InStr(1, desc, "sdr", vbTextCompare) Then
            searchSDR = False
        Else
            searchSDR = True
        End If
    End If
End Function

'look for the text "`qcgeneral" in the memo and/or description fields of M2M
Public Function searchQC(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`qcneral", vbTextCompare) Then
        searchQC = False
    Else
        If InStr(1, desc, "`qcgeneral", vbTextCompare) Then
            searchQC = False
        Else
            searchQC = True
        End If
    End If
End Function

'look for the text "`qcbm" in the memo and/or description fields of M2M
Public Function searchQCBM(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`qcbm", vbTextCompare) Then
        searchQCBM = False
    Else
        If InStr(1, desc, "`qcbm", vbTextCompare) Then
            searchQCBM = False
        Else
            searchQCBM = True
        End If
    End If
End Function

'look for the text "`qcbm" in the memo and/or description fields of M2M
Public Function searchQCRH(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`qcRH", vbTextCompare) Then
        searchQCRH = False
    Else
        If InStr(1, desc, "`qcbm", vbTextCompare) Then
            searchQCRH = False
        Else
            searchQCRH = True
        End If
    End If
End Function

'look for the text "`xml" in the memo and/or description fields of M2M
Public Function searchXML(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`xml", vbTextCompare) Then
        searchXML = False
    Else
        If InStr(1, desc, "`xml", vbTextCompare) Then
            searchXML = False
        Else
            searchXML = True
        End If
    End If
End Function

'look for the text "`ready" in the memo and/or description fields of M2M
Public Function searchReady(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`ready", vbTextCompare) Then
        searchReady = False
    Else
        If InStr(1, desc, "`ready", vbTextCompare) Then
            searchReady = False
        Else
            searchReady = True
        End If
    End If
End Function

'look for the text "`desrev" in the memo and/or description fields of M2M
Public Function searchDesignReview(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`desrev", vbTextCompare) Then
        searchDesignReview = False
    Else
        If InStr(1, desc, "`desrev", vbTextCompare) Then
            searchDesignReview = False
        Else
            searchDesignReview = True
        End If
    End If
End Function
'look for the text "`desrev" in the memo and/or description fields of M2M
Public Function searchPendingCancel(memo As String, desc As String) As Boolean
    If InStr(1, memo, "`cancel", vbTextCompare) Then
        searchPendingCancel = False
    Else
        If InStr(1, desc, "`cancel", vbTextCompare) Then
            searchPendingCancel = False
        Else
            searchPendingCancel = True
        End If
    End If
End Function
'check if an item exists in a collection
Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:

    Contains = False
End Function

'writes the relevant information to the TravelerTool when a job fails to print at any point
Public Function rejectedJob(job As String, reason As String, filename As String, rejectedCount As Integer) As Integer
    
    Worksheets(strShtTool).Range("D" & rejectedCount) = job
    Worksheets(strShtTool).Range("E" & rejectedCount) = reason
    Worksheets(strShtTool).Range("F" & rejectedCount) = filename
    tmpFile = ""
    strError = ""
    pressError = ""
    dxfError = ""
    dimsError = ""
    rejectedJob = rejectedCount + 1
End Function

'scan the entire XML archive for indexing purposes
Public Function scanXMLArchive() As Collection
    Dim xmlFiles As Collection
    Dim strFileName As String
    Set xmlFiles = New Collection
    
    'scan main folder
    strFileName = Dir$(strArchivePath & "\" & "*.xml")
    Do Until StrComp(strFileName, "") = 0
        xmlFiles.Add strArchivePath & "\" & strFileName
        strFileName = Dir
    Loop
    
    'scan legacy folder  
    strFileName = Dir$(strLegPath & "\" & "*.xml")
    Do Until StrComp(strFileName, "") = 0
        xmlFiles.Add strLegPath & "\" & strFileName
        strFileName = Dir
    Loop
    
    'scan new folder
    strFileName = Dir$(strNewPath & "\" & "*.xml")
    Do Until StrComp(strFileName, "") = 0
        xmlFiles.Add strNewPath & "\" & strFileName
        strFileName = Dir
    Loop
    
    Set scanXMLArchive = xmlFiles
End Function

Sub CheckBox6_Click()

End Sub

Sub CheckBox9_Click()

End Sub

Sub CheckBox3_Click()

End Sub

Sub CheckBox2_Click()

End Sub

Sub CheckBox10_Click()

End Sub

Sub CheckBox12_Click()

End Sub

Sub CheckBox13_Click()

End Sub

Sub CheckBox14_Click()

End Sub
