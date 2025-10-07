VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub createBurnlist_Btn_Click()

    If fileOperations.lookForBadPartNames() Then
        Exit Sub
    End If
    
    Call mptLgcyFldr

    'declare variables
    Dim fso As Object
    Dim oFile As Object
    Dim strFile As String
    Dim xmlFile As Object
    Dim backup As Object
    Dim strXML As String
    Dim rejected As Boolean
    Dim rejectedCount As Integer
    Dim dateTime As String
    Dim sucessCount As Integer
    
    Call clearColumn("L", strByTool)
    Call clearColumn("M", strByTool)
    Call clearColumn("N", strByTool)
    
    Set fso = CreateObject("Scripting.FileSystemObject")                                            'set file system object
    Set xmlFile = fso.createTextFile(strCutPath & "\xmlCutlist\cutlist.erp")                        'initialize xml file
    Set backup = fso.createTextFile(strCutPath & "\xmlCutlist\erp backup" & dateTime & ".erp")     'initialize xml file backup
    'Set xmlFile = fso.createTextFile("C:\Users\epart\Desktop\cutlist.erp")                        'initialize xml file
    'Set backup = fso.createTextFile("C:\Users\epart\Desktop\" & dateTime & ".erp")     'initialize xml file backup
    dateTime = Now()        'initialize the date variable
    dateTime = Replace(dateTime, "/", "-")      'format the time so it can be used as a filename
    dateTime = Replace(dateTime, ":", "_")      'format the time so it can be used as a filename
    
    strXML = ""             'initialize other variables
    rejected = False
    rejectedCount = 2
    sucessCount = 0
    
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If
    
    Call connectDatabase                        'create database connection
    Call connQueryUpdate(connQry, burnSQL(strByTool))     'execute the sql query and refresh on query sheet
    Set rst = New recordSheet
    rst.setsheet (strShtQry)          'initialize rst record set
        
    xmlFile.writeline "<ErpExchange>"           'open outer XML tags
    backup.writeline "<ErpExchange>"
    xmlFile.writeline vbTab & "<Orders>"
    backup.writeline vbTab & "<Orders>"
    
    Dim i As Integer
    i = 3

    Do While rst.hasNext()      'loop until all relevant jobs have burnlist xml
        Dim jNumField As String
        jNumField = rst.field(oNum)
        Dim tempFilepath As String
        If engCompleted(CStr(Split(rst.field(oNum), "-")(0)), strByTool) Then
            strFile = getOldJobFile(jNumField)
            If StrComp(Left(rst.field(oNum), 1), "I", vbTextCompare) = 0 Then
                strFile = getFileName(rst.field(pnum), getRevNum(rst.field(oNum), rst.field("Memo"), rst.field("Rev"), ""))
                tempFilepath = strCutPath & "\" & strFile
            ElseIf StrComp(strFile, "", vbTextCompare) = 0 Then
                strError = "There is no File associated with this Job"
            Else
                tempFilepath = Dir$(strCutPath & "\" & strFile)
                If StrComp(tempFilepath, "", vbTextCompare) = 0 Then
                    tempFilepath = strLegPath & "\" & Dir$(strLegPath & "\" & strFile)
                Else
                    tempFilepath = strCutPath & "\" & tempFilepath
                End If
            End If
        Else
            strFile = ""
            strError = "Engineering not completed."
        End If
        
        If Len(strFile) <> 0 Then   'create xml for ErpOrder if a filename was returned
            
            Dim tmpXML As String
            tmpXML = addXML(rst.field(q), tempFilepath, False)      'grab sub-parts
            
            If StrComp(tmpXML, "") Then     'append order to xml string if the sub-parts exist
                strXML = vbTab & vbTab & "<ErpOrder>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<ImportType>NewOrder</ImportType>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<OrderNumber>" & Trim(UCase(rst.field(oNum))) & "</OrderNumber>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<StartDate>" & Format(rst.field(sDate), "yyyy-mm-dd") & "</StartDate>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<TargetDate>" & Format(rst.field(tDate), "yyyy-mm-dd") & "</TargetDate>" & vbCrLf
                'strXML = strXML & vbTab & vbTab & vbTab & "<Info1" & "Hello World" & "</Info1>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<ProductionStrategy>MaterialAdministrationOrder</ProductionStrategy>" & vbCrLf
                strXML = strXML & vbTab & vbTab & vbTab & "<Automatic>True</Automatic>" & vbCrLf
                strXML = strXML & tmpXML
                strXML = strXML & vbTab & vbTab & "</ErpOrder>"
                Dim order As String, tempOrder() As String
                tempOrder = Split(rst.field(oNum), "-")
                If engCompleted(tempOrder(0), strByTool) Then
                    xmlFile.writeline (strXML)      'write the order to the xml file
                    backup.writeline (strXML)       'write the order to the backup file
                Else
                    rejected = True
                    If StrComp(strError, "", vbTextCompare) Then                'error message
                        strError = strError & "   Engineering is not yet completed"
                    Else
                        strError = "Engineering is not yet completed"
                    End If
                End If
                tmpXML = ""
                sucessCount = sucessCount + 1
                Worksheets(strByTool).Range("O" & sucessCount) = rst.field(oNum)      'add to list of successfully add orders
            Else
                If StrComp(strError, "", vbTextCompare) = 0 Then          'error message
                    strError = "Error loading data from XML"
                End If
                rejected = True
            End If
        Else
            If StrComp(strError, "", vbTextCompare) = 0 Then               'error message
                strError = "Filename not found"
                Worksheets(strByTool).Range("N" & rejectedCount + 1) = tmpFile
                tmpFile = ""
            End If
        End If
        If StrComp(strError, "", vbTextCompare) Then                       'add rejected order numbers to "Rejected Job Orders" list with reason why
            rejectedCount = rejectedCount + 1
            Worksheets(strByTool).Range("L" & rejectedCount) = rst.field(oNum)
            Worksheets(strByTool).Range("M" & rejectedCount) = strError
            strError = ""
        End If
        i = i + 1
        Call insertJobFile(rst.field(oNum), strFile)
    Loop
    
    xmlFile.writeline vbTab & "</Orders>"       'close outer XML tags
    backup.writeline vbTab & "</Orders>"
    xmlFile.writeline "</ErpExchange>"
    backup.writeline "</ErpExchange>"
    
    xmlFile.Close                               'close and clean up files
    backup.Close
    Set fso = Nothing
    Set oFile = Nothing
    
    Dim temp As String
    For i = 1 To Worksheets(strByTool).Cells(rows.count, "O").End(xlUp).row            'add order numbers that were burnt to reference access database
        temp = Worksheets(strByTool).Range("O" & i)
        If StrComp(temp, "") <> 0 Then
            Worksheets(strByTool).Range("O" & i) = ""
        End If
    Next
End Sub
