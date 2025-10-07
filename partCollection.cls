VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "partCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private partCol As Collection
Private materialDict As Object
Public seeToggle As Boolean

Public Sub init(quant As Integer, strFileName As String, isCinci As Boolean)
    'initialize the partcol variable
    Set partCol = getXML(quant, strFileName, isCinci)
End Sub

Public Function hasParts() As Boolean
    'checks that the collection has parts
    If Not partCol Is Nothing Then
        hasParts = True
    Else
        hasParts = False
    End If
End Function

Public Function getPartsArr() As Variant
    'converts the part collection to a sorted array
    getPartsArr = sort(partCol)
End Function

Public Function compileMaterials(jobno As String) As Object
    'returns a dictionary with an entry for each material
    'the materials are summed together and paired with their respective key
    Set materialDict = CreateObject("Scripting.dictionary")
    Dim part As Variant
    For Each part In partCol
        
            If (part.getIsTemplate() And part.getCutTemplate()) Then
                'Make sure the material for a template is spread out over the entire quantity
                Dim strSQL As String, quant As recordSheet
                strSQL = "SELECT fjobno, fquantity" & vbCrLf & _
                         "FROM jomast" & vbCrLf & _
                         "WHERE jomast.fjobno = '" & jobno & "'"
            
                Call connQueryUpdate(connQry3, strSQL)
                Debug.Print (jobno)
                Set quant = New recordSheet
                quant.setsheet (bomTbl)
                quant.hasNext
            
                If StrComp(part.getMaterial, "", vbTextCompare) And part.addToReport(0) Then
                    If materialDict.exists(part.getMaterial) Then
                        materialDict(part.getMaterial) = materialDict(part.getMaterial) + (part.getArea / quant.field("fquantity") / part.getYQty)
                    Else
                        materialDict.add item:=(part.getArea / quant.field("fquantity") / part.getYQty), key:=part.getMaterial
                    End If
                End If
            ElseIf Not part.getIsTemplate() Then
                If StrComp(part.getMaterial, "", vbTextCompare) And part.addToReport(0) Then
                    If materialDict.exists(part.getMaterial) Then
                          materialDict(part.getMaterial) = materialDict(part.getMaterial) + (part.getArea * part.getQuant / part.getYQty)
                    Else
                        materialDict.add item:=(part.getArea * part.getQuant / part.getYQty), key:=part.getMaterial
                    End If
                End If
            End If
        'handle exception
        If part.getweighttoggle = "Yes" Then
            materialDict(part.getMaterial) = part.getweightoverride
            seeToggle = True
        Else
            seeToggle = False
        End If
        
    Next
    Dim key As Variant
    For Each key In materialDict.keys
        'If seeToggle = False Then
            materialDict(key) = materialDict(key) * getMatWgtSc(CStr(key)) / 144
            
        'Else
         '   materialDict(key) = materialDict(key)

        'End If
    Next
    Set compileMaterials = materialDict
End Function

'function to insert data into the report sheet and return if the insert was successful
Public Function chkInsertData(text As String, jobno As String, partno As String) As Boolean

    If findFQC1RUN(jobno) Then
        reportTitle = "****FAI****           CUT LIST FOR JOB " & jobno
    Else
        reportTitle = "CUT LIST FOR JOB " & jobno
    End If
    Call genericFormat_Report(partno)           'set generic format of report
    
    'declare and initialize variables
    chkInsertData = False
    Dim chkPrs As Boolean
    Dim partsArr() As Variant
    Dim parArrLen As Integer
    
    chkPrs = True
    partsArr = getPartsArr
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
            If temp > 0 Then                                                  'populate data for this part                                           'set return value to true
                i = i + temp
                Call rowPadding(2 * i + 5, text)                              'insert a row of padding
                chkPrs = chkPrs And checkPress(partsArr(h))
                chkInsertData = True
                Call checkDXF(partsArr(h))
            End If
            
        Next
        Call endingLines(2 * i + 7, text)                                     'insert ending lines on the bottom of the report
        
    End If
    
    If Worksheets(strShtTool).Shapes("Check Box 10").ControlFormat.value <> 1 Then                          'check that the "Test lists" check box is checked, if it is, skip the following code
        If StrComp(strError, "", vbTextCompare) Then
            'Debug.Print (strError)
            chkInsertData = False
        Else
            'Debug.Print (strError)
            chkInsertData = chkInsertData And chkPrs
        End If
    Else
        chkInsertData = True
    End If
    
End Function

Private Function checkPress(parts As Variant) As Boolean
    'checks that the press program has been marked as complete or not possible for the the bystronics
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
            If StrComp(pressError, "", vbTextCompare) = 0 Then
                pressError = parts.getProg() & "needs Press Program"
                checkPress = False
            Else
                pressError = pressError & ", " & parts.getProg() & " needs Press Program"
                checkPress = False
            End If
        End If
    Else
        If StrComp(pressError, "", vbTextCompare) = 0 Then
            pressError = parts.getProg() & " needs Press Program"
            checkPress = False
        Else
            pressError = pressError & ", " & parts.getProg() & " needs Press Program"
            checkPress = False
        End If
    End If
    
    rs.Close
    
End Function

 
Private Sub checkDXF(part As Variant)
    'checks for a dxf of each program in the dxfinch folder
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

Private Function getXML(quant As Integer, strFileName As String, isCinci As Boolean) As Collection
    'takes the xml file and creates a collection of all the parts on the file
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
            strError = "xml is incomplete"
            Set getXML = Nothing
            Exit Function
        End If
        
        Set tmpChildNodes = xmlNode.ChildNodes
        
        'grab childnodes of that path and
        'pass them into getParts for further processing
        Set getXML = getParts(quant, tmpChildNodes)
    Else
        Set getXML = Nothing
        strError = "There is an error in the XML"
    End If
End Function

'recursive function to get make items
'this function is designed to stop its recursive calls at the
'last "make" item found for an ErpPart it will still loop
'to the next "make" item for whatever 'depth' it's currently at in the xml
Public Function getParts(quant As Integer, xml As Object, Optional ByVal isHardware As Boolean, Optional descr2 As String, Optional descr3 As String) As Collection

    Dim partList As Collection
    Set partList = New Collection
    Dim tempPart As partInfo
    Dim tempPartList As Collection
    Dim refNode As Object
    Dim config As String, attrib As String
    Dim description2 As String
    Dim description3 As String
    
    Set tempPartList = New Collection
    config = "configuration"
    attrib = "attribute"
    description2 = ""
    description3 = ""
    If Not IsMissing(descr2) Then
        If StrComp(descr2, "", vbTextCompare) Then
            description2 = descr2
        End If
    End If
    If Not IsMissing(descr3) Then
        If StrComp(descr3, "", vbTextCompare) Then
            description3 = descr3
        End If
    End If
    
    For Each refNode In xml     'loops over each childnode that was passed into the function as xml (a list of nodes)
    
        Dim nodeName As String
        
        nodeName = refNode.nodeName
        
        Select Case refNode.nodeName
        
            Case attrib
            Case config
                Dim partQuant As partInfo
                Set partQuant = getData(refNode.ChildNodes(), quant)        'send the node list to the getQuantity() function
                
                If Not partQuant Is Nothing And partQuant.makeXML Or Not IsMissing(isHardware) Then                                   'check that makeXML is true
                    If partQuant.getHardwareLot Then
                        If StrComp(partQuant.getDesc2, "", vbTextCompare) Then
                            description2 = partQuant.getDesc2
                        End If
                        If StrComp(partQuant.getDesc3, "", vbTextCompare) Then
                            description3 = partQuant.getDesc3
                        End If
                        Set tempPartList = getParts(partQuant.getQuantity, refNode.ChildNodes(), True, description2, description3)
                        Dim par As partInfo
                        For Each par In tempPartList
                            Call partQuant.addPart(par)
                        Next
                        Set tempPartList = New Collection
                    ElseIf Not partQuant.getAdd Then
                        If refNode.HasChildNodes Then
                            Set tempPartList = getParts(partQuant.getQuantity, refNode.ChildNodes(), False, description2, description3)
                            If Not tempPartList Is Nothing Then
                                Dim part As Variant
                                For Each part In tempPartList
                                    partList.add item:=part
                                Next
                            End If
                            Set tempPartList = New Collection
                        End If
                    End If
                    partList.add item:=partQuant
                    
                    'see if the configuration has childnodes
                    'with a nested configuration tag and
                    'if it does, step into that childnode
                ElseIf partQuant.getMake Then
                     If refNode.HasChildNodes Then
                        Set tempPartList = getParts(partQuant.getQuantity, refNode.ChildNodes(), False, description2, description3)
                        If Not tempPartList Is Nothing Then
                            For Each part In tempPartList
                                partList.add item:=part
                            Next
                        End If
                        Set tempPartList = New Collection
                    End If
                End If
            Case Else
                'if the node is not a configuration tag, attempt
                'to step into the xml another childnode deeper
                If refNode.HasChildNodes Then
                    Set tempPartList = getParts(quant, refNode.ChildNodes(), False, description2, description3)
                    If Not tempPartList Is Nothing Then
                        For Each part In tempPartList
                            partList.add item:=part
                        Next
                    End If
                    Set tempPartList = New Collection
                End If
        End Select
    Next
    If partList.count = 0 Then
        Set getParts = Nothing
        Exit Function
    End If
    Set getParts = partList
End Function

'function that returns a string for a single make item
Private Function getData(xml As Object, Optional quantity As Variant, Optional descr2 As String, Optional descr3 As String) As partInfo
    Dim quant As Integer
    quant = 1
    If Not IsMissing(quantity) Then
        If IsNumeric(quantity) Then
            quant = CInt(quantity)
        End If
    End If
    
    'declare variables
    Dim attNames
    Dim part As partInfo
    Dim xNode As Object
    Set part = New partInfo
    Call part.init
    Call part.setMake
    attNames = Array("Name", "PDMrotation", "PDMrawmatlnum", "Reference Count", "PDMsmparttoggle", "PDMgangqty", "PDMchilddesc", "PDMpartweight", "State", "PDMmaxX", "PDMmaxY", "PDMpunch", "PDMsawcut", "PDMlooseprt", "PDMshiploose", "PDMpowdered", "PDMtemplate", "PDMtemplatecut", "PDMhardwarelot", "PDMassemblycut", "PDMshear", "PDMchilddesc2", "PDMchilddesc3", "PDMweighttoggle", "PDMweightoverride")
    
    For Each xNode In xml   'loop through all nodes in xml (a list of nodes)
        
        If StrComp(xNode.nodeName, "attribute") = 0 Then    'check if a node/tag is named attribute
            Dim tagAttr As String
            tagAttr = xNode.Attributes.getNamedItem("name").text    'grab text from the "name" field of the attribute tag
            
            Dim str As String
            str = xNode.Attributes.getNamedItem("value").text
            
            Select Case tagAttr     'Select case based on value in tagAttr
                
                Case attNames(0)    'set Name
                    part.setProg (str)
                
                Case attNames(1)    'set PDMrotation
                    part.setRot (str)
               
                Case attNames(2)    'set PDMrawmatlnum
                    part.setMaterial (str)
                
                Case attNames(3)    'set Reference Count
                    Call part.setQuant(quant * CInt(str))
                
                Case attNames(4)    'set PDMsmparttoggle
                    part.setMake (str)
                
                Case attNames(5)    'set PDMgangqty
                    part.setYQty (str)
                    
                Case attNames(6)    'set PDMchilddesc
                    part.setDesc (str)
                    
                Case attNames(7)    'set PDMpartweight
                    part.setWeight (str)
                    
                Case attNames(8)    'set State
                    part.setState (str)
                
                Case attNames(9)    'set xAxis
                    part.setXAxis (str)
                
                Case attNames(10)   'set yAxis
                    part.setYAxis (str)
                    
                Case attNames(11)   'PDMpunch
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setPunch (True)
                    Else
                        part.setPunch (False)
                    End If
                    
                Case attNames(12)   'PDMsawcut
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setSaw (True)
                    Else
                        part.setSaw (False)
                    End If
                    
                Case attNames(13)   'PDMlooseprt
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setLoosePart (True)
                    Else
                        part.setLoosePart (False)
                    End If
                    
                Case attNames(14)   'PDMshiploose
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setShipLoose (True)
                    Else
                        part.setShipLoose (False)
                    End If
                    
                Case attNames(15)   'PDMpowdered
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setPowdered (True)
                    Else
                        part.setPowdered (False)
                    End If
                    
                Case attNames(16)   'PDMtemplate
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setIsTemplate (True)
                    Else
                        part.setIsTemplate (False)
                    End If
                    
                Case attNames(17)   'PDMtemplatecut
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setCutTemplate (True)
                    Else
                        part.setCutTemplate (False)
                    End If
                    
                Case attNames(18)   'PDMhardwarelot
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setHardwareLot (True)
                    Else
                        part.setHardwareLot (False)
                    End If
                    
                Case attNames(19)   'PDMassemblycut
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setAssemblyCut (True)
                    Else
                        part.setAssemblyCut (False)
                    End If
                    
                Case attNames(20)   'PDMshear
                    If StrComp(str, "1", vbTextCompare) = 0 Then
                        part.setShear (True)
                    Else
                        part.setShear (False)
                    End If
                    
                Case attNames(21)    'set PDMchilddesc2
                    If Not IsMissing(descr2) And StrComp(descr2, "", vbTextCompare) Then
                        part.setDesc2 (descr2)
                    Else
                        part.setDesc2 (str)
                    End If
                    
                Case attNames(22)    'set PDMchilddesc3
                    If Not IsMissing(descr3) And StrComp(descr3, "", vbTextCompare) Then
                        part.setDesc3 (descr3)
                    Else
                        part.setDesc3 (str)
                    End If
                Case attNames(23)    'set PDMweighttoggle
                    part.setweighttoggle (str)
                
                Case attNames(24)    'set PDMweightoverride
                    If IsNumeric(str) Then
                        part.setweightoverride (1# * str)
                    Else
                        part.setweightoverride (0)
                    End If
                    
            End Select
        Else
            Exit For
        End If
    Next
    
    If part.getMake And part.getExtCut Then             'check if part is a make item and that extention is psm or sldprt
        If part.getMatIsMissing And part.getExtCut And StrComp("SWELD", part.getProg, vbTextCompare) Then   'check for missing materials
            If StrComp(strError, "", vbTextCompare) = 0 Then
                strError = part.getProg & " Material is missing"
            Else
                strError = strError & ", " & part.getProg & " Material is missing"
            End If
        End If
        If ((part.verDims And StrComp("SWELD", part.getProg, vbTextCompare)) Or (part.verOneDim And StrComp("SWELD", part.getProg, vbTextCompare) = 0)) Then    '
            If StrComp(dimsError, "", vbTextCompare) = 0 Then
                dimsError = part.getProg & " has missing dimensions"
            Else
                dimsError = dimsError & ", " & part.getProg & " has missing dimensions"
            End If
        End If
    End If
    part.orderAxis
    Set getData = part      'return the part
End Function

'Combines identical parts that have identical desc2 & desc3 then Sorts the given collection using the Arrays.MergeSort algorithm.
' O(n log(n)) time
' O(n) space
Private Function sort(col As Collection) As Variant
    Dim tempColl As Collection
    Set tempColl = New Collection
    Dim notInColl As Boolean
    Dim par As Variant
    notInColl = True
    For Each par In col
        Dim index As Integer
        For index = 1 To CInt(tempColl.count)
            If StrComp(tempColl(index).getActProg, par.getActProg, vbTextCompare) = 0 And StrComp(tempColl(index).getDesc2, par.getDesc2, vbTextCompare) = 0 And StrComp(tempColl(index).getDesc3, par.getDesc3, vbTextCompare) = 0 And StrComp(tempColl(index).getProg, par.getProg, vbTextCompare) = 0 Then
                notInColl = False
                tempColl.item(index).addQuant (par.getActQuant)
                Exit For
            End If
        Next
        If notInColl Then
            tempColl.add item:=par
        End If
        notInColl = True
    Next par
    Set col = tempColl
    Dim a() As Variant
    a = ToArray(col)
    arrSort a()
    sort = a()
End Function

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Private Function ToArray(col As Collection) As Variant
    Dim a() As Variant
    ReDim a(0 To col.count - 1)
    Dim i As Long
    i = 0
    Dim par As Variant
    For Each par In col
        Set a(i) = par
        i = i + 1
    Next par
    ToArray = a()
End Function

'Sorts the array using the MergeSort algorithm (follows the Java legacyMergesort algorithm
'O(n*log(n)) time; O(n) space
Private Sub arrSort(ByRef a() As Variant)
    MergeSort a, 0, UBound(a) - LBound(a)
End Sub

Private Sub MergeSort(ByRef dest() As Variant, low As Long, high As Long)
    Dim length As Integer
    Dim i As Integer, p As Integer
    Dim temp As partInfo

    length = high - low + 1

    ' insertion sort on small arrays
    If length <= 7 Then
        i = low
        Dim j As Integer
        Do While i <= high
            j = i
            Do While True
                If (j <= low) Then
                    Exit Do
                End If
                'Debug.Print (dest(j - 1).getDesc2 & " " & dest(j - 1).getDesc3 & " " & dest(j - 1).getProg & " - " & dest(j).getDesc2 & " " & dest(j).getDesc3 & " " & dest(j).getProg)
                If StrComp(dest(j - 1).getDesc2, dest(j).getDesc2, vbTextCompare) = 0 Then
                    If StrComp(dest(j - 1).getDesc3, dest(j).getDesc3, vbTextCompare) = 0 Then
                        If (StrComp(dest(j - 1).getProg, dest(j).getProg, vbTextCompare) > 0) Then
                            Call swap(dest(), j - 1, j)
                        Else
                            Exit Do
                        End If
                    ElseIf StrComp(dest(j - 1).getDesc3, dest(j).getDesc3, vbTextCompare) > 0 Then
                        Call swap(dest(), j - 1, j)
                    Else
                        Exit Do
                    End If
                ElseIf StrComp(dest(j - 1).getDesc2, dest(j).getDesc2, vbTextCompare) > 0 Then
                    Call swap(dest(), j - 1, j)
                Else
                    Exit Do
                End If
                j = j - 1 'decrement j
            Loop
            i = i + 1 'increment i
        Loop
        Exit Sub
    End If

    'recursively sort halves of dest into src
    Dim lowLeft As Long, highLeft As Long, lowRight As Long, highRight As Long
    lowLeft = low
    highLeft = (low + high) / 2
    highLeft = highLeft - 1
    lowRight = (low + high) / 2
    highRight = high
    MergeSort dest, lowLeft, highLeft
    MergeSort dest, lowRight, highRight
    Dim index As Integer
    index = low
    
    Dim copy1() As Variant
    Dim copy2() As Variant
    ReDim copy1(lowLeft To highLeft)
    ReDim copy2(lowRight To highRight)
    Dim mytemp As Integer
    
    For index = 0 To highLeft - lowLeft
        mytemp = index + lowLeft
        Set copy1(index + lowLeft) = dest(index + lowLeft)
    Next
    
    For index = 0 To highRight - lowRight
        mytemp = index + lowRight
        Set copy2(index + lowRight) = dest(index + lowRight)
    Next
    
    'merge sorted halves into dest
    index = lowLeft
    i = lowLeft
    p = lowRight
    Do While i <= highLeft And p <= highRight
        If StrComp(copy1(i).getDesc2, copy2(p).getDesc2, vbTextCompare) = 0 Then
            If StrComp(copy1(i).getDesc3, copy2(p).getDesc3, vbTextCompare) = 0 Then
                If (StrComp(copy1(i).getProg, copy2(p).getProg, vbTextCompare) < 1) Then
                    Set dest(index) = copy1(i)
                    i = i + 1
                Else
                    Set dest(index) = copy2(p)
                    p = p + 1
                End If
            ElseIf StrComp(copy1(i).getDesc3, copy2(p).getDesc3, vbTextCompare) < 1 Then
                Set dest(index) = copy1(i)
                i = i + 1
            Else
                Set dest(index) = copy2(p)
                p = p + 1
            End If
        ElseIf StrComp(copy1(i).getDesc2, copy2(p).getDesc2, vbTextCompare) < 1 Then
            Set dest(index) = copy1(i)
            i = i + 1
        Else
            Set dest(index) = copy2(p)
            p = p + 1
        End If
        index = index + 1
    Loop
    If i <= highLeft Then
        Do While i <= highLeft
            Set dest(index) = copy1(i)
            i = i + 1
            index = index + 1
        Loop
    ElseIf p <= highRight Then
        Do While p <= highRight
            Set dest(index) = copy2(p)
            p = p + 1
            index = index + 1
        Loop
    End If
End Sub

Private Sub swap(ByRef a() As Variant, i As Integer, p As Integer)
    'swaps two partInfo entries
    Dim temp As partInfo
    Set temp = a(i)
    Set a(i) = a(p)
    Set a(p) = temp
End Sub

Private Sub printArr(arr() As Variant)
    'used for testing to verify the sort algorithm was functioning as intended
    Dim temp As Variant, index As Integer
    index = 0
    
    For Each temp In arr
        Debug.Print (index & " - " & temp.getDesc2 & temp.getDesc3 & temp.getProg)
        index = index + 1
    Next
    Debug.Print (vbCrLf)
End Sub



Private Function getMatWgtSc(mat As String) As Double
    'returns the scrap factor times the per sq foot weight of the material
    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT ScrapFactor * Pounds AS lbSc" & vbCrLf & _
    "FROM [MaterialTable]" & vbCrLf & _
    "WHERE MaterialPartNo='" & mat & "'"
    
    If TypeName(rs) <> "ADODB.Recordset" Then
        Set rs = CreateObject("ADODB.Recordset")
    End If

    rs.Open strSQL, DBCONT          'execute sql query on burntlist table
    
    If rs.RecordCount > 0 Then          'return results of sql query
        getMatWgtSc = rs(0)
    Else
        getMatWgtSc = 0
    End If
    
    rs.Close
End Function

Public Function consolColl() As Collection
    'removes multiple entries of the same program number summing the total together
    Dim tmpArr As Variant, parArrLen As Integer
    tmpArr = getPartsArr
    parArrLen = UBound(tmpArr) - LBound(tmpArr) + 1
    Dim index As Integer, temp As Collection
    Set temp = New Collection
    For index = 1 To parArrLen
        temp.add item:=tmpArr(index - 1)
    Next
    Set consolColl = temp
End Function

Public Sub setColl(newColl As Collection)
    'sets the collection equal to a new collection
    Set partCol = newColl
End Sub
