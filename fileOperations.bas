Attribute VB_Name = "fileOperations"
'clears all data from the given column on a specific sheet
Public Sub clearColumn(column As String, sheet As String)

    Dim pos As Integer
    For pos = 3 To Worksheets(sheet).Cells(rows.count, column).End(xlUp).row
        Worksheets(sheet).Range(column & pos) = ""
    Next
End Sub

'looks for faulty xml file names
Public Function lookForBadPartNames() As Boolean
    Dim strFileName As String
    strFileName = Dir$(strCutPath & "\" & "*.xml")
    
    Do Until StrComp(strFileName, "") = 0           'loop until dir returns no more filenames

        If (strFileName Like "*REL#####.XML") Then
            MsgBox "Filename should be of <M2M Part Number>_REV##_REL##.xml format" & vbCrLf & "and " & strFileName & vbCrLf & "does not meet the standard format."
            lookForBadPartNames = True
            Exit Function
        End If
        
        strFileName = Dir
    Loop
    
    lookForBadPartNames = False
End Function

'Function to determine if a string is alphanumeric
Public Function AlphaNumeric(pValue) As Boolean

    'declare variables
   Dim LPos As Integer
   Dim LChar As String
   Dim LValid_Values As String


   LPos = 1        'Start at first character in pValue

   LValid_Values = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789."           'Set up values that are considered to be alphanumeric 'NTF ADD . 02/21/2023

   While LPos <= Len(pValue)           'Test each character in pValue

      LChar = Mid(pValue, LPos, 1)        'Single character in pValue

      If InStr(LValid_Values, LChar) = 0 Then             'If character is not alphanumeric, return FALSE
         AlphaNumeric = False
         Exit Function
      End If

      LPos = LPos + 1             'Increment counter

   Wend

   AlphaNumeric = True         'Value is alphanumeric, return TRUE

End Function

'clear out the contents of the legacy folder, while inserting the filenames into the legacy database
Public Sub mptLgcyFldr()

    Dim strFileName As String
    strFileName = Dir$(strNewPath & "\" & "*.xml")
    Dim legacyFiles As Collection
    Set legacyFiles = New Collection
    Dim currentLegacy As Collection
    Dim fileToggle As Boolean
    
    Do Until StrComp(strFileName, "") = 0           'loop until dir returns no more filenames
        
        Dim curFile As Variant
        Dim index As Integer
        index = 1
        For Each curFile In legacyFiles
            If StrComp(Left(CStr(curFile), InStr(1, CStr(curFile), "_REV", vbTextCompare) - 1), Left(strFileName, InStr(1, strFileName, "_REV", vbTextCompare) - 1), vbTextCompare) = 0 Then
                Dim curRev As Variant, curRel As Integer, newRev As Variant, newRel As Integer
                curRev = Mid(CStr(curFile), InStr(1, CStr(curFile), "_REV", vbTextCompare) + 4, 2)
                curRel = CInt(Mid(CStr(curFile), InStr(1, CStr(curFile), "_REL", vbTextCompare) + 4, InStr(1, CStr(curFile), ".XML", vbTextCompare) - (InStr(1, CStr(curFile), "_REL", vbTextCompare) + 4)))
                newRev = Mid(strFileName, InStr(1, strFileName, "_REV", vbTextCompare) + 4, 2)
                newRel = CInt(Mid(strFileName, InStr(1, strFileName, "_REL", vbTextCompare) + 4, InStr(1, strFileName, ".XML", vbTextCompare) - (InStr(1, strFileName, "_REL", vbTextCompare) + 4)))
                Debug.Print (curRev & " " & curRel & " " & newRev & " " & newRel)
                If IsNumeric(curRev) And IsNumeric(newRev) Then
                    If CInt(curRev) < CInt(newRev) Then
                        legacyFiles.Remove (index)
                        legacyFiles.add item:=strFileName
                    ElseIf CInt(curRev) = CInt(newRev) Then
                        If curRel < newRel Then
                            legacyFiles.Remove (index)
                            legacyFiles.add item:=strFileName
                        End If
                    End If
                Else
                    If StrComp(CStr(curRev), CStr(newRev), vbTextCompare) < 0 Then
                        legacyFiles.Remove (index)
                        legacyFiles.add item:=strFileName
                    ElseIf StrComp(CStr(curRev), CStr(newRev), vbTextCompare) = 0 Then
                        If curRel < newRel Then
                            legacyFiles.Remove (index)
                            legacyFiles.add item:=strFileName
                        End If
                    End If
                End If
                index = 0
            Else
                index = index + 1
            End If
            
            If index = 0 Then
                Exit For
            End If
            
        Next curFile
        If legacyFiles.count = 0 Then
            legacyFiles.add item:=strFileName
        ElseIf index <> 0 Then
            legacyFiles.add item:=strFileName
        End If
        strFileName = Dir
    Loop

    Dim file As Variant
    
    For Each file In legacyFiles
    
        Dim thisFile As String
        thisFile = Left(CStr(file), InStr(1, CStr(file), "_REV", vbTextCompare) - 1) & "_REV*"
        strFileName = Dir$(strCutPath & "\" & Left(thisFile, InStr(1, thisFile, "_REV", vbTextCompare) - 1) & "*")
        Set currentLegacy = New Collection
        
        Dim tmpName As String
        Do Until strFileName = ""
            tmpName = Left(strFileName, InStr(1, strFileName, "_REV", vbTextCompare) - 1) & "_REV*"
            If StrComp(thisFile, tmpName, vbTextCompare) = 0 Then
                currentLegacy.add strFileName
            End If
            strFileName = Dir
        Loop
        
        Dim currentFile As Variant
        Dim cFile As String
        For Each currentFile In currentLegacy
            cFile = CStr(currentFile)
            'copy file from main CUT LIST XML folder to legacy folder
            Call FileCopy(strCutPath & "\" & cFile, strLegPath & "\" & CStr(cFile))
            'delete file from main CUT LIST XML folder
            Call deleteFile(cFile, strCutPath)
        Next currentFile
        
        Call FileCopy(strNewPath & "\" & CStr(file), strCutPath & "\" & CStr(file))
        Call deleteFile(CStr(file), strNewPath)
        strFileName = Dir$(strNewPath & "\" & Left(thisFile, InStr(1, thisFile, "_REV", vbTextCompare) - 1) & "*")
        
        Do Until strFileName = ""
            tmpName = Left(strFileName, InStr(1, strFileName, "_REV", vbTextCompare) - 1) & "_REV*"
            If StrComp(thisFile, tmpName, vbTextCompare) = 0 Then
                currentLegacy.add strFileName
            End If
            strFileName = Dir
        Loop
        
        For Each currentFile In currentLegacy
            cFile = CStr(currentFile)
            'delete file from legacy folder and CUT LIST XML
            Call deleteFile(cFile, strNewPath)
        Next currentFile
        
    Next file
    
End Sub

'sub to delete files based on folder location and filename
Public Sub deleteFile(strFile As String, path As String)

    Dim file As String
    file = path & "\" & strFile
    'check that filename exists
    'if so, remove any attribute (read-only)
    'delete file
    If (Dir(file) <> "") Then
        SetAttr file, vbNormal
        Kill file
    End If
End Sub



'grabs the first valid rev from the three that are supplied
Public Function getRevNum(oNum As String, memo As String, rev As String, Optional shtRev As Variant) As String
    Dim temp As String
    If Not IsMissing(shtRev) Then           'check that this function was passed a rev from the main sheet
        If AlphaNumeric(CStr(shtRev)) Then        'check that the rev is alphanumeric
            If InStr(1, CStr(shtRev), "NS", vbTextCompare) = 0 Then       'check that the rev is not NS
                If StrComp(CStr(shtRev), "", vbTextCompare) Then
                    getRevNum = CStr(shtRev)                            'return rev
                    Exit Function
                End If
            End If
        End If
    End If
    If AlphaNumeric(rev) Then       'check that the rev parameter is alphanumeric
        If InStr(1, rev, "NS", vbTextCompare) = 0 Then          'check that the rev parameter is not NS
            getRevNum = rev         'return rev
            Exit Function
        End If
    End If
    getRevNum = getRevNumatMemo(memo)       'grab and return rev from the memo field of M2M
End Function

'returns rev from the memo field of M2M
Public Function getRevNumatMemo(memo As String) As String
    Dim revNum As Variant
    If InStr(1, memo, "`rev", vbTextCompare) Then        'check that "rev" exists in the field
        Dim splStr() As String
        splStr = Split(memo, " ")           'split the memo by spaces
        For Each revNum In splStr                       'loop through each token to find the token containing "rev"
            If InStr(1, revNum, "`rev", vbTextCompare) = 1 Then          'search for "rev" in the token
                If Len(revNum) = 7 Then                    'check that the length of the token is 6
                    Dim tmp As String
                    tmp = Right(revNum, 3)                 'grab the 3 characters on the right (the rev number)
                    If InStr(1, tmp, "*", vbTextCompare) = 0 Then       'check that the 3 characters does not contain a "*"
                        If IsNumeric(tmp) Then          'check that the token is numeric
                            revNum = tmp                'token to return
                            Exit For
                        ElseIf AlphaNumeric(tmp) Then   'check that the token is alphanumeric
                            If InStr(1, tmp, "NS", vbTextCompare) = 0 Then      'check that the token does not contain "NS"
                                revNum = tmp                'token to return
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If Not IsNull(revNum) Then              'check if revNum is null, return value
        getRevNumatMemo = revNum
    Else
        getRevNumatMemo = ""
    End If
End Function

'create a basic structure for the filename to search for,
'use wildcards in the place of the REL number
Public Function getFileName(partno As String, rev As String) As String
    
    'declare variables
    Dim strFileName As String
    Dim tmpStrFileName As String

    If IsNumeric(Trim(UCase(rev))) Then     'check that the rev is numeric
        If Len(Trim(UCase(rev))) <> 2 Then      'make the rev exactly 2 digits
            If Len(Trim(UCase(rev))) <> 3 Then
                strFileName = Trim(UCase(partno)) & "_REV0" & Trim(UCase(rev)) & "*.xml"
            Else
                strFileName = Trim(UCase(partno)) & "_REV" & Right(Trim(UCase(rev)), 2) & "*.xml"
            End If
        Else
            strFileName = Trim(UCase(partno)) & "_REV" & Trim(UCase(rev)) & "*.xml"
        End If
        
    Else
        If StrComp(Trim(UCase(rev)), "NS") = 0 Or StrComp(Trim(UCase(rev)), "") = 0 Then        'check that the rev is not NS
            strFileName = ""
        Else
            strFileName = Trim(UCase(partno)) & "_REV" & Right(Trim(UCase(rev)), 2) & "*.xml"
        End If
    End If
    
    'check for 3 digit rev #, then check for 2
    tmpStrFileName = strFileName
    
    strFileName = Trim(UCase(partno)) & "_REV" & Trim(UCase(rev)) & "*.xml"
    tmpFile = strFileName
    getFileName = searchFileName(strFileName)       'pass the predicted filename into searchFileName then return the filename
    
    If StrComp(getFileName, "") = 0 Then
        tmpFile = tmpStrFileName
        strFileName = tmpStrFileName
        getFileName = searchFileName(strFileName)
    End If
End Function

'search for the filename with the given structure
'wildcards in the place of the REL number
Public Function searchFileName(strFileName As String) As String
    
    If StrComp(strFileName, "") = 0 Then
        searchFileName = ""
        Exit Function
    End If
    
    'declare variables
    Dim strFile As String
    Dim tempFile As String
    Dim intTempRel As Integer
    Dim intRel As Integer
    tempFile = ""
    intRel = 0
    intTempRel = 0
    strFile = ""
    
    strFile = Dir$(strCutPath & "\" & strFileName)      'retrieve filename

    Do Until StrComp(strFile, "") = 0           'loop until there are no more filenames that match the given filename structure
    
        intTempRel = CInt(Mid(strFile, InStr(1, strFile, "REL", vbTextCompare) + 3, 2))

        If intTempRel > intRel Then             'looking for the highest REL number for files that match the given filename structure
            tempFile = strFile
            intRel = intTempRel
        End If

        strFile = Dir               'grab next filename
    Loop
    
    If StrComp(tempFile, "") Then           'grab file that best fit the criteria
        strFile = tempFile
    End If

    searchFileName = strFile                    'return file
    
End Function

'fucntion that initiates the creation of the xml tags and values
Public Function addXML(quant As Integer, strFileName As String, isCinci As Boolean) As String

    'declare variables
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim parXML As partCollection
    Dim partsList As New Collection
    Dim strXML As String
    Dim part As partInfo
    
    Set parXML = New partCollection
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    If xmlDoc.Load(strFileName) Then            'check that the file will load

        Set xmlNode = xmlDoc.SelectSingleNode(xmlPath)              'set the known xml path
        
        Dim tmpChildNodes As Object
        If xmlNode Is Nothing Then
            If StrComp(strError, "", vbTextCompare) Then
                strError = strError & "    xml is incomplete"
            Else
                strError = "xml is incomplete"
            End If
            addXML = ""
            Exit Function
        End If
        Set tmpChildNodes = xmlNode.ChildNodes
        'grab childnodes of that path and
        'pass them into nestRef for further processing
        Call parXML.setColl(parXML.getParts(quant, tmpChildNodes))
        Set partsList = parXML.consolColl
        'check that nestRef returned a non-empty string,
        'complete the part xml, and return the string
        Dim tmpprtLst As Collection
        Set tmpprtLst = New Collection
        For Each part In partsList
            If part.cutPart Then
               tmpprtLst.add part
            End If
        Next
        
        Set partsList = tmpprtLst
        
        If partsList.count <> 0 Then
            Dim tmpDebit As Integer
            If isCinci Then
                For Each part In partsList
                    If part.getIsTemplate Then
                        If part.getCutTemplate Then
                            tmpDebit = 1
                        Else
                            tmpDebit = 0
                        End If
                    Else
                        tmpDebit = WorksheetFunction.RoundUp(part.getQuant / part.getYQty, 0)
                    End If
                    strXML = strXML & "PART," & part.getProg & "," & part.getQuality & "," & part.getThickness & "," & _
                    tmpDebit & ",DXF,g:\engineering\dxf-inch\" & part.getProg & ".dxf," & rst.field(oNum) & "," & vbCrLf
                    tmpDebit = 0
                Next
            Else
                strXML = vbTab & vbTab & vbTab & "<Parts>" & vbCrLf
                For Each part In partsList
                    If part.getIsTemplate Then
                        If part.getCutTemplate Then
                            tmpDebit = 1
                        Else
                            tmpDebit = 0
                        End If
                    Else
                        tmpDebit = WorksheetFunction.RoundUp(part.getQuant / part.getYQty, 0)
                    End If
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & "<ErpPart>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<BysoftCode>" & part.getProg & "</BysoftCode>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<Debit>" & tmpDebit & "</Debit>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<MaterialCode>" & part.getQuality & "</MaterialCode>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<Measure>Inch</Measure>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<Thickness>" & part.getThickness & "</Thickness>" & vbCrLf
                    'strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<Info1>" & "HelloWorld" & "</Info1>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<RotationAllowance>Angle" & part.getRot & "</RotationAllowance>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & vbTab & "<FillPart>False</FillPart>" & vbCrLf
                    strXML = strXML & vbTab & vbTab & vbTab & vbTab & "</ErpPart>" & vbCrLf
                    tmpDebit = 0
                Next
                strXML = strXML & vbTab & vbTab & vbTab & "</Parts>" & vbCrLf
            End If
            addXML = strXML
        Else
            addXML = ""
        End If
    End If
    
    Set xmlDoc = Nothing                'close the xmlDoc
    
End Function

'set generic format of the report
Public Sub genericFormat_Report(partno As String)
    Dim padWidth As Double
    Dim checkWidth As Double
    padWidth = 0.5
    checkWidth = 5
    With Worksheets(currentReport)                                  'use report sheet
        With .Range("A1:AA1")                                   'first row settings
            .RowHeight = 36#
            .Font.Size = 28
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Merge
            .value = reportTitle
            .Font.Bold = True
        End With
    
        With .Range("A2:D2")                                    'second row settings
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
    
        With .Range("E2:N2")
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .value = partno
            .Font.Bold = True
        End With
        With .Range("O2:P2")
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        With .Range("Q2:AA2")
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .value = "INITIAL UPON COMPLETION"
            .Font.Bold = True
        End With
    
        .rows(3).Font.Bold = True                           'third row settings
        .rows(3).HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 17                      'NTF CHANGED FROM 10
        .Columns("B").ColumnWidth = padWidth
        .Columns("C").ColumnWidth = 39
        .Columns("D").ColumnWidth = padWidth
        .Columns("E").ColumnWidth = 3
        .Columns("F").ColumnWidth = padWidth
        .Columns("G").ColumnWidth = 3
        .Columns("H").ColumnWidth = padWidth
        .Columns("I").ColumnWidth = 5
        .Columns("J").ColumnWidth = padWidth
        .Columns("K").ColumnWidth = 5
        .Columns("L").ColumnWidth = padWidth
        .Columns("M").ColumnWidth = 5
        .Columns("N").ColumnWidth = padWidth
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = padWidth
        .Columns("Q").ColumnWidth = checkWidth
        .Columns("R").ColumnWidth = padWidth
        .Columns("S").ColumnWidth = checkWidth
        .Columns("T").ColumnWidth = padWidth
        .Columns("U").ColumnWidth = checkWidth
        .Columns("V").ColumnWidth = padWidth
        .Columns("W").ColumnWidth = checkWidth
        .Columns("X").ColumnWidth = padWidth
        .Columns("Y").ColumnWidth = checkWidth
        .Columns("Z").ColumnWidth = padWidth
        .Columns("AA").ColumnWidth = checkWidth
        
        .Range("A3").value = "PROG"
        .Range("C3").value = "PART DESCRIPTION"
        .Range("E3").value = "QTY"
        .Range("G3").value = "YQTY"
        .Range("I3").value = "XAX"
        .Range("K3").value = "YAX"
        .Range("M3").value = "GA"
        .Range("O3").value = "QUALITY"
        .Range("Q3").value = "ENG"
        .Range("S3").value = "NST"
        .Range("U3").value = "LSR"
        .Range("W3").value = "PCH"
        .Range("Y3").value = "FRM"
        .Range("AA3").value = "PEM"
        
        With .Range("A4:AA4")                               'fourth row settings
            .Merge
            .RowHeight = 4
            .Interior.Color = RGB(128, 128, 128)
        End With
    End With
End Sub

'set generic format of the report
Public Sub genericFormat_ShipReport(partno As String)
    Dim padWidth As Double
    Dim checkWidth As Double
    padWidth = 0.5
    checkWidth = 15
    With Worksheets(currentReport)                                  'use report sheet
        With .Range("A1:K1")                                   'first row settings
            .RowHeight = 36#
            .Font.Size = 28
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Merge
            .value = reportTitle
            .Font.Bold = True
        End With
        
        With .Range("A2:K2")                                    'second row settings
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .value = partno
            .Font.Bold = True
        End With
    
        .rows(3).Font.Bold = True                           'third row settings
        .rows(3).HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = padWidth
        .Columns("C").ColumnWidth = 50
        .Columns("D").ColumnWidth = padWidth
        .Columns("E").ColumnWidth = 4
        .Columns("F").ColumnWidth = padWidth
        .Columns("G").ColumnWidth = checkWidth
        .Columns("H").ColumnWidth = padWidth
        .Columns("I").ColumnWidth = checkWidth
        .Columns("J").ColumnWidth = padWidth
        .Columns("K").ColumnWidth = checkWidth
        
        .Range("A3").value = "PROG"
        .Range("C3").value = "PART DESCRIPTION"
        .Range("E3").value = "QTY"
        .Range("G3").value = "QTY SHIPPED"
        .Range("I3").value = "QTY BACKORDER"
        .Range("K3").value = "PICKED BY"
        
        With .Range("A4:K4")                               'fourth row settings
            .Merge
            .RowHeight = 4
            .Interior.Color = RGB(128, 128, 128)
        End With
    End With
End Sub


'set generic format of the report
Public Sub genericFormat_LooseReport(partno As String)
    Dim padWidth As Double
    Dim checkWidth As Double
    padWidth = 0.5
    checkWidth = 6
    With Worksheets(currentReport)                                  'use report sheet
        With .Range("A1:Q1")                                   'first row settings
            .RowHeight = 36#
            .Font.Size = 28
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Merge
            .value = reportTitle
            .Font.Bold = True
        End With
    
        With .Range("A2:D2")                                    'second row settings
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .value = partno
            .Font.Bold = True
        End With
        With .Range("E2:Q2")
            .Merge
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
            .value = "INITIAL UPON COMPLETION"
            .Font.Bold = True
        End With
    
        .rows(3).Font.Bold = True                           'third row settings
        .rows(3).HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = padWidth
        .Columns("C").ColumnWidth = 50
        .Columns("D").ColumnWidth = padWidth
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = padWidth
        .Columns("G").ColumnWidth = 4
        .Columns("H").ColumnWidth = padWidth
        .Columns("I").ColumnWidth = checkWidth
        .Columns("J").ColumnWidth = padWidth
        .Columns("K").ColumnWidth = checkWidth
        .Columns("L").ColumnWidth = padWidth
        .Columns("M").ColumnWidth = checkWidth
        .Columns("N").ColumnWidth = padWidth
        .Columns("O").ColumnWidth = checkWidth
        .Columns("P").ColumnWidth = padWidth
        .Columns("Q").ColumnWidth = checkWidth
        
        .Range("A3").value = "PART NUMBER"
        .Range("C3").value = "PART DESCRIPTION"
        .Range("E3").value = "PAINTED"
        .Range("G3").value = "QTY"
        .Range("I3").value = "WELD"
        .Range("K3").value = "Q.C."
        .Range("M3").value = "CLEAN"
        .Range("O3").value = "PAINT"
        .Range("Q3").value = "SHIP"
        
        With .Range("A4:Q4")                               'fourth row settings
            .Merge
            .RowHeight = 4
            .Interior.Color = RGB(128, 128, 128)
        End With
    End With
End Sub


'function to insert data into the report sheet and return if the insert was successful
Public Function insertData(text As String, p As partCollection) As Boolean

    'declare and initialize variables
    insertData = False
    Dim parts As Collection
    Dim strFile As String
    Dim chkPrs As Boolean
    Dim partsArr() As Variant
    Dim parArrLen As Integer
    
    chkPrs = True
    partsArr = p.getPartsArr
    parArrLen = UBound(partsArr) - LBound(partsArr) + 1

    If (parArrLen <> 0) Then                                        'check if the parts collection contains data
    
        Dim i As Integer
        Dim items As Integer
        Dim part As Variant
        Dim h As Integer
        
        i = 0
        items = 2
        
        
        Call clearColumn("A", evrythgSht)
        Call clearColumn("B", evrythgSht)
        Call clearColumn("C", evrythgSht)
        Call clearColumn("D", evrythgSht)
        Call clearColumn("E", evrythgSht)
        Call clearColumn("F", evrythgSht)
        Call clearColumn("G", evrythgSht)
        
        Worksheets(evrythgSht).Range("A" & 1) = "Program Name"
        Worksheets(evrythgSht).Range("B" & 1) = "getMake"
        Worksheets(evrythgSht).Range("C" & 1) = "getAdd"
        Worksheets(evrythgSht).Range("D" & 1) = "Material"
        Worksheets(evrythgSht).Range("E" & 1) = "Quantity"
        Worksheets(evrythgSht).Range("F" & 1) = "Quality"
        Worksheets(evrythgSht).Range("G" & 1) = "Thickness"
        
        Call rowPadding(2 * i + 5, text)                                      'insert a row of padding
        For h = 0 To parArrLen - 1                                          'loop through all the parts in the collection of parts
            Dim temp As Integer
            
            Worksheets(evrythgSht).Range("A" & items) = partsArr(h).everythingName()
            Worksheets(evrythgSht).Range("B" & items) = partsArr(h).getMake()
            Worksheets(evrythgSht).Range("C" & items) = partsArr(h).getAdd()
            Worksheets(evrythgSht).Range("D" & items) = partsArr(h).getMaterial()
            Worksheets(evrythgSht).Range("E" & items) = partsArr(h).getQuant()
            Worksheets(evrythgSht).Range("F" & items) = partsArr(h).getQuality()
            Worksheets(evrythgSht).Range("G" & items) = partsArr(h).getThickness()
            items = items + 1
            
            temp = populateRow(partsArr(h), 2 * i + 6, text)
            If temp > 0 Then                                                  'populate data for this part
                Call rowPadding(2 * i + 7, text)                              'insert a row of padding
                chkPrs = chkPrs And checkPress(partsArr(h))
                insertData = True                                             'set return value to true
                i = i + temp
                Call checkDXF(partsArr(h))
            End If
            
        Next
        Call endingLines(2 * i + 7, text)                                     'insert ending lines on the bottom of the report
        
    End If
    If StrComp(strError, "", vbTextCompare) Then
        Debug.Print (strError)
        insertData = False
    Else
        Debug.Print (strError)
        insertData = insertData And chkPrs
    End If
    
End Function

Public Function checkPress(parts As Variant) As Boolean
    
    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT [150 Programmed], [40 Programmed]" & vbCrLf & _
    "FROM [Press Programs]" & vbCrLf & _
    "WHERE [Form Detail]='" & parts.getProg() & "';"

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        Dim tmp1 As Integer, tmp2 As Integer
        tmp1 = rs(0)
        tmp2 = rs(1)
        If tmp1 > 0 Then
            checkPress = True
        ElseIf tmp2 > 0 Then
            checkPress = True
        Else
            pressError = pressError & " " & parts.getProg()
            checkPress = False
        End If
    Else
        pressError = pressError & " " & parts.getProg()
        checkPress = False
    End If
    
    rs.Close
    
End Function

'checks that dxf exists for a given part
Public Sub checkDXF(part As Variant)

    Dim strFile As String
    
    strFile = Dir$(dxfInch & "\" & part.getProg() & ".dxf")
    
    If StrComp(strFile, "", vbTextCompare) = 0 Then
        If StrComp(dxfError, "", vbTextCompare) = 0 Then
            dxfError = part.getProg()
        Else
            dxfError = dxfError & " " & part.getProg()
        End If
    End If
    
End Sub

'add padding between the rows of data
Public Sub rowPadding(row As Integer, text As String)

    Dim strCol As String
    
    If StrComp(text, "cutlist", vbTextCompare) = 0 Then
        strCol = ":AA"
    ElseIf StrComp(text, "shiploose", vbTextCompare) = 0 Then
        strCol = ":K"
    ElseIf StrComp(text, "loosepart", vbTextCompare) = 0 Then
        strCol = ":Q"
    End If
    
    With Worksheets(currentReport).Range("A" & row & strCol & row)
        .Merge
        .RowHeight = 4
    End With
    
End Sub

'populate the row with data from the part
Public Function populateRow(part As Variant, row As Integer, text As String) As Integer

    Dim column()
    Dim tradRound As Integer
    Dim whichList As Integer
    Dim rowsPopulated As Integer
    rowpopulated = 0
    
    tradRound = 0.0000000001
    
    If StrComp(text, "cutlist", vbTextCompare) = 0 Then
        whichList = 0
        column = Array("Q", "S", "U", "W", "Y", "AA")
    ElseIf StrComp(text, "shiploose", vbTextCompare) = 0 Then
        whichList = 1
        column = Array("G", "I", "K")
    ElseIf StrComp(text, "loosepart", vbTextCompare) = 0 Then
        whichList = 2
        column = Array("I", "K", "M", "O", "Q")
    End If
    
    If part.addToReport(whichList) Then                     'check that the part can be added to the report
        If StrComp(part.getDesc2, "", vbTextCompare) = 0 Then
        ElseIf StrComp(part.getDesc2, desc2, vbTextCompare) <> 0 Then
            desc2 = part.getDesc2
            Worksheets(currentReport).Range("C" & row).value = desc2
            Worksheets(currentReport).Range("C" & row).Font.Bold = True
            Call rowPadding(row + 1, text)
            Call rowPadding(row + 3, text)
            rowsPopulated = rowsPopulated + 1
            row = row + 2
            If StrComp(part.getDesc3, "", vbTextCompare) <> 0 Then
                desc3 = part.getDesc3
                Worksheets(currentReport).Range("C" & row).value = desc3
                Worksheets(currentReport).Range("C" & row).Font.Bold = True
                Call rowPadding(row + 1, text)
                Call rowPadding(row + 3, text)
                rowsPopulated = rowsPopulated + 1
                row = row + 2
            Else
                desc3 = ""
            End If
        ElseIf StrComp(part.getDesc3, desc3, vbTextCompare) <> 0 Then
            desc3 = part.getDesc3
            Worksheets(currentReport).Range("C" & row).value = desc3
            Worksheets(currentReport).Range("C" & row).Font.Bold = True
            Call rowPadding(row + 1, text)
            Call rowPadding(row + 3, text)
            rowsPopulated = rowsPopulated + 1
            row = row + 2
        End If
        
    
        With Worksheets(currentReport)                          'use the report sheet
            .Range("A" & row).value = part.getProg()        'get the program
            .Range("C" & row).value = part.getDesc()        'get the description
            
            If part.getIsTemplate Then
                If part.getCutTemplate Then
                    .Range("E" & row).value = 1             'set quantity to 1
                Else
                    .Range("E" & row).value = 0             'set quantity to 0
                End If
                
            Else
                If Not part.getHardwareLot And Not whichList = 2 Then
                    .Range("E" & row).value = part.getQuant() * orderQty / part.getYQty()   'get the quantity
                End If
            End If
            
            If whichList = 0 Then
                .Range("G" & row).value = part.getYQty()        'get the yield
            
                If part.xAxisIsNull Then                        'check if the xAxis variable is null
                    .Range("I" & row).value = ""
                Else
                    .Range("I" & row).value = Round(part.getXAxis + tradRound, 2)
                End If
                
                If part.yAxisIsNull Or part.getYAxis = 0 Then   'check if the yAxis variable is null
                    .Range("K" & row).value = ""
                Else
                    .Range("K" & row).value = Round(part.getYAxis + tradRound, 2)
                End If
                
                .Range("M" & row).value = part.getGauge         'get the gauge of the material
                .Range("O" & row).value = part.getQuality       'get the material type
            ElseIf whichList = 2 Then
                If part.getPowdered() Then
                    .Range("E" & row).value = "YES"
                Else
                    .Range("E" & row).value = "NO"
                End If
                .Range("G" & row).value = part.getQuant() * orderQty
            End If
            
            Dim i As Integer
                
            For i = 0 To UBound(column)                     'create boxes to initialize in
            
                If Not part.getHardwareLot Then
                    With .Range(column(i) & row).Borders
                        .Color = RGB(128, 128, 128)
                        .LineStyle = xlContinuous
                        .weight = xlMedium
                    End With
                End If
            Next
            rowsPopulated = rowsPopulated + 1
            row = row + 2
        End With
        
        If part.getHardwareLot Then
            
            Dim par As partInfo
            Dim partCol As Collection
            Set partCol = part.getParts()
            
            For Each par In partCol
                Call rowPadding(row - 1, text)
                With Worksheets(currentReport)                          'use the report sheet
                    .Range("C" & row).value = par.getDesc()        'get the description
                    .Range("E" & row).value = par.getQuant()       'get the quantity
                    
                    For i = 0 To UBound(column)                     'create boxes to initialize in
                    
                        With .Range(column(i) & row).Borders
                            .Color = RGB(128, 128, 128)
                            .LineStyle = xlContinuous
                            .weight = xlMedium
                        End With
                        
                    Next
                    
                    rowsPopulated = rowsPopulated + 1
                    row = row + 2
                End With
            Next
            
        End If
        
    Else
        rowsPopulated = 0
    End If
    populateRow = rowsPopulated
    
End Function

'insert the lines on the bottom of the report
Public Sub endingLines(row As Integer, text As String)

    Dim strCol As String
    
    If StrComp(text, "cutlist", vbTextCompare) = 0 Then
        strCol = ":AA"
    ElseIf StrComp(text, "shiploose", vbTextCompare) = 0 Then
        strCol = ":K"
    ElseIf StrComp(text, "loosepart", vbTextCompare) = 0 Then
        strCol = ":Q"
    End If
    
    Dim i As Integer
    
    For i = 0 To 2
    
        With Worksheets(currentReport).Range("A" & row + i & strCol & row + i)
            .Borders(xlEdgeBottom).weight = xlThin
            .RowHeight = 20
        End With
        
    Next
    
End Sub

