VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "partInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private prog As String              'stores the program number for bysoft
Private desc As String              'stores the part description
Private desc2 As String              'stores the part description
Private desc3 As String              'stores the part description
Private qty As Integer              'stores the quanity to be produced
Private yQty As Integer             'stores the yield quantity
Private xAxis As Double             'stores the xAxis length if it exists
Private yAxis As Double             'stores the yAxis length if it exists
Private gauge As String             'stores the gauge of the material
Private quality As String           'stores the material type
Private rot As Integer              'stores the rotational allowance
Private material As String          'stores the materail code from M2M
Private thickness As Double         'stores the thickness of the material
Private make As Boolean             'determines if a part is a make item or not
Private weight As Double            'stores the weight of the part
Private state As String             'stores the state of the part
Private add As Boolean              'stores value to add part to list
Private saw As Boolean              'stores value to saw toggle
Private shear As Boolean            'stores value to shear toggle
Private powdered As Boolean         'stores value to powdered toggle
Private loosePart As Boolean        'stores value to loose part toggle
Private shiploose As Boolean        'stores value to ship loose toggle
Private punch As Boolean            'stores value to punch toggle
Private isTemplate As Boolean       'stores value to identify template
Private cutTemplate As Boolean      'stores the value to cut templates
Private assemblyCut As Boolean      'stores whether something is cut at assebmly
Private hardwareLot As Boolean      'stores whether hardware is shipped loose
Private matIsMissing As Boolean     'stores value that knows if a raw material exists
Private partCol As Collection       'stores hardware in a hardware lot
Private ext As String               'stores the file extension
Public wgtoverride As Double    'stores weight override
Private wgttoggle As String       'stores weighttoggle

Public Sub init()
    'initializes the partcol variable as a new collection
    Set partCol = New Collection
End Sub

Public Function cutPart() As Boolean
    'function that defines if a part should be placed on the burnlist or not
        cutPart = False
    
    If shear Then
        cutPart = False
        Exit Function
    End If
    
    If saw Then
        cutPart = False
        Exit Function
    End If
        
    If assemblyCut Then
        cutPart = False
        Exit Function
    End If
    
    'If loosePart Then
        'cutPart = False
        'Exit Function
    'End If
    
    If hardwareLot Then
        cutPart = False
        Exit Function
    End If
    
    If shiploose Then
        cutPart = False
        Exit Function
    End If
    
    If Not make Then        'return false if this is not a make item
        cutPart = False
        Exit Function
    ElseIf StrComp(prog, "", vbTextCompare) = 0 Then    'check that the program variable is not empty
        cutPart = False
        Exit Function
    ElseIf IsNull(rot) Then     'check that the rotation allowance has a value
        rot = 90
    ElseIf IsNull(qty) Then     'check that qty is not empty
        cutPart = False
        Exit Function
    ElseIf qty = 0 Then     'check that qty is not zero
        cutPart = False
        Exit Function
    ElseIf StrComp(quality, "", vbTextCompare) = 0 Then     'check that quality is not empty
        cutPart = False
        Exit Function
    ElseIf IsNull(thickness) Then       'check that the thickness is not null
        cutPart = False
        Exit Function
    ElseIf thickness = 0 Then       'check that the thickness is not zero
        cutPart = False
        Exit Function
    ElseIf StrComp(ext, "psm", vbTextCompare) And StrComp(ext, "sldprt", vbTextCompare) Then
        cutPart = False
        Exit Function
    End If
    
    cutPart = True
End Function

Public Function makeXML() As Boolean
    'verifies that the part has all the information to create the XML
    makeXML = False
    
    If shear Then
        makeXML = True
        Exit Function
    End If
    
    If saw Then
        makeXML = True
        Exit Function
    End If
        
    If assemblyCut Then
        makeXML = True
        Exit Function
    End If
    
    If loosePart Then
        makeXML = True
        Exit Function
    End If
    
    If hardwareLot Then
        makeXML = True
        Exit Function
    End If
    
    If shiploose Then
        makeXML = True
        Exit Function
    End If
    
    If Not make Then        'return false if this is not a make item
        makeXML = False
        Exit Function
    ElseIf StrComp(prog, "", vbTextCompare) = 0 Then    'check that the program variable is not empty
        makeXML = False
        Exit Function
    ElseIf IsNull(rot) Then     'check that the rotation allowance has a value
        rot = 90
    ElseIf IsNull(qty) Then     'check that qty is not empty
        makeXML = False
        Exit Function
    ElseIf qty = 0 Then     'check that qty is not zero
        makeXML = False
        Exit Function
    ElseIf StrComp(quality, "", vbTextCompare) = 0 Then     'check that quality is not empty
        makeXML = False
        Exit Function
    ElseIf IsNull(thickness) Then       'check that the thickness is not null
        makeXML = False
        Exit Function
    ElseIf thickness = 0 Then       'check that the thickness is not zero
        makeXML = False
        Exit Function
    End If
    
    makeXML = True
    
End Function

Public Function addRemakeToReport() As Boolean
    'determines if it`s ok to add a program to the remake cutlist
    If Not make Then
        addRemakeToReport = False
        Exit Function
    ElseIf StrComp(prog, "", vbTextCompare) = 0 Then
        addRemakeToReport = False
        Exit Function
    ElseIf IsNull(qty) Then
        addRemakeToReport = False
        Exit Function
    ElseIf qty = 0 Then
        addRemakeToReport = False
        Exit Function
    ElseIf StrComp(gauge, "", vbTextCompare) = 0 Then
        addRemakeToReport = False
        Exit Function
    ElseIf StrComp(quality, "", vbTextCompare) = 0 Then
        addRemakeToReport = False
        Exit Function
    End If
    If IsNull(yQty) Then
        yQty = 1
        Exit Function
    ElseIf yQty < 1 Then
        yQty = 1
    End If
    addRemakeToReport = True
End Function

Public Function addToReport(formsOnly As Integer) As Boolean
    'determines if it`s ok to add a program to the current report
    addToReport = False
    
    Select Case formsOnly
    
        Case 0
        
            If hardwareLot Then
                addToReport = False
                Exit Function
            End If
        
            If saw Then
                If qty = 0 Then
                    addToReport = False
                    Exit Function
                End If
                addToReport = True
                Exit Function
            End If
            
            If shear Then
                If qty = 0 Then
                    addToReport = False
                    Exit Function
                End If
                addToReport = True
                Exit Function
            End If
            
            If assemblyCut Then
                If qty = 0 Then
                    addToReport = False
                    Exit Function
                End If
                addToReport = True
                Exit Function
            End If
        
            If Not make Then
                addToReport = False
                Exit Function
            ElseIf StrComp(prog, "", vbTextCompare) = 0 Then
                addToReport = False
                Exit Function
            ElseIf IsNull(qty) Then
                addToReport = False
                Exit Function
            ElseIf qty = 0 Then
                addToReport = False
                Exit Function
            ElseIf StrComp(gauge, "", vbTextCompare) = 0 Then
                addToReport = False
                Exit Function
            ElseIf StrComp(quality, "", vbTextCompare) = 0 Then
                addToReport = False
                Exit Function
            End If
            
            If IsNull(yQty) Then
                yQty = 1
                Exit Function
            ElseIf yQty < 1 Then
                yQty = 1
            End If
            
            addToReport = True
            
        Case 1
        
            If shiploose And make Then
                If qty = 0 Then
                    addToReport = False
                    Exit Function
                End If
                addToReport = True
            End If
    
        Case 2
        
            If hardwareLot Then
                addToReport = False
                Exit Function
            End If

            If loosePart And make Then
                If qty = 0 Then
                    addToReport = False
                    Exit Function
                End If
                addToReport = True
            End If
    
    End Select
    
End Function

Public Function getQuantity() As Integer
    'returns the quantity of 0 if it's not set to make or hardwareLot,
    '1 for templates marked to be cut all the time
    '0 for templates not marked to be cut all the time
    'qty for anything else
    If make Or hardwareLot Then
        If isTemplate Then
            If cutTemplate Then
                getQuantity = 1
            Else
                getQuantity = 0
            End If
        Else
            getQuantity = qty
        End If
    Else
        getQuantity = 0
    End If
    
End Function

Public Function xAxisIsNull() As Boolean
    'checks for null values in the xaxis
    xAxisIsNull = IsNull(xAxis)
End Function

Public Function yAxisIsNull() As Boolean
    'checks for null values in the yaxis
    yAxisIsNull = IsNull(yAxis)
End Function

Public Sub setProg(program As String)
    'sets the name of the program
    Dim tmp() As String
    
    tmp = Split(program, ".")
    Call setExt(tmp(UBound(tmp)))
    If StrComp(tmp(1), "psm", vbTextCompare) = 0 Then
        add = True
    ElseIf StrComp(tmp(1), "sldprt", vbTextCompare) = 0 Then
        add = True
    Else
        add = False
    End If
    
    prog = tmp(0)
    
End Sub

Public Function getProg() As String
    'when marked "saw" returns sweld
    '"shear" returns shear
    '"assemblyCut" returns assembly cut
    '"hardwareLot" returns hardware lot
    'otherwise order the axis and return the name of the program
    If saw Then
        yAxis = 0
        getProg = "SWELD"
        
    ElseIf shear Then
        getProg = "SHEAR"
    ElseIf assemblyCut Then
        getProg = "ASSEMBLY CUT"
    ElseIf hardwareLot Then
        getProg = "HARDWARE LOT"
    Else
        Me.orderAxis
        getProg = prog
    End If
    
End Function

Public Function getActProg() As String
    'returns the name of the program
    'regardless of anything else
    getActProg = prog
End Function

Public Sub setDesc(description As String)
    'sets desc
    desc = description
End Sub

Public Function getDesc() As String
    'returns nothing for hardwareLot
    'otherwise returns the desc
    If hardwareLot Then
        getDesc = ""
    Else
        getDesc = desc
    End If
End Function

Public Sub setDesc2(description As String)
    'sets desc2
    desc2 = description
End Sub

Public Function getDesc2() As String
    'returns nothing for hardwareLot
    'otherwise returns desc2
    If hardwareLot Then
        getDesc2 = ""
    Else
        getDesc2 = desc2
    End If
End Function

Public Sub setDesc3(description As String)
    'sets desc3
    desc3 = description
End Sub

Public Function getDesc3() As String
    'returns nothing for hardwareLot
    'otherwise returns desc3
    If hardwareLot Then
        getDesc3 = ""
    Else
        getDesc3 = desc3
    End If
End Function

Public Sub setQuant(quantity As String)
    'sets the quantity
    If IsNumeric(quantity) Then
        qty = CInt(quantity)
    End If
    
End Sub

Public Function getQuant() As Integer
    'returns 1 for templates marked to always cut
    'returns 0 for templates marked to not always cut
    'returns the quantity otherwise
    If isTemplate Then
        If cutTemplate Then
            getQuant = 1
        Else
            getQuant = 0
        End If
    Else
        getQuant = qty
    End If
End Function

Public Function getActQuant() As Integer
    'returns the quantity
    getActQuant = qty
End Function

Public Sub addQuant(quantity As Integer)
    'adds to the quantity (used for multiples entries of the same item, summing them together for clarity)
    qty = qty + quantity
End Sub

Public Sub setYQty(yQuantity As String)
    'sets the yeild of this program
    If IsNumeric(yQuantity) Then
        yQty = CInt(yQuantity)
    End If
    
End Sub

Public Function getYQty() As Integer
    'returns the yeild of this program
    If yQty < 1 Then
        getYQty = 1
    Else
        getYQty = yQty
    End If
    
End Function

Public Sub setXAxis(axis As String)
    'sets the xaxis
    If IsNumeric(axis) Then
        xAxis = CDbl(axis)
    End If
    
End Sub

Public Function getXAxis() As Double
    'reutrns the xaxis
    getXAxis = xAxis
End Function

Public Sub setYAxis(axis As String)
    'sets the yaxis
    If IsNumeric(axis) Then
        yAxis = CDbl(axis)
    End If
    
End Sub

Public Function getYAxis() As Double
    'returns the yaxis
    getYAxis = yAxis
End Function

Public Sub setGauge(ga As String)
    'sets the gauge for this program
    gauge = ga
End Sub

Public Function getGauge() As String
    'returns the gauge for this program
    getGauge = gauge
End Function

Public Sub setQuality(qual As String)
    'sets the quality of the material for this program
    quality = qual
End Sub

Public Function getQuality() As String
    'returns the quality of the material
    getQuality = quality
End Function

Public Sub setRot(rotation As String)
    'sets the allowed rotation of this part
    If IsNumeric(rotation) Then
        rot = CInt(rotation)
    Else
        rot = 90
    End If
    
End Sub

Public Function getRot() As Integer
    'returns the allowable rotation for this part
    getRot = rot
End Function
Public Sub setweightoverride(override As Double)
    'weight override value
     wgtoverride = override
End Sub
Public Function getweightoverride() As Double
    'returns bypass weight override value
    
    getweightoverride = wgtoverride

End Function
Public Sub setweighttoggle(toggle As String)
    'bypass weight
     wgttoggle = toggle
End Sub
Public Function getweighttoggle() As String
    'returns bypass weight
    getweighttoggle = wgttoggle
End Function


Public Sub setMaterial(materialCode As String)
    'sets the material for this program
    material = materialCode
    
    If StrComp(materialCode, "", vbTextCompare) = 0 Then
        Exit Sub
    ElseIf StrComp(materialCode, "CONFIG FILE, IGNORE", vbTextCompare) = 0 Then
        Exit Sub
    End If
    
    Dim strSQL As String                'create sql query to check for order number in the Burnt list database
    strSQL = "SELECT BysoftMaterialCode, Thickness, Gauge" & vbCrLf & _
    "FROM MaterialTable" & vbCrLf & _
    "WHERE MaterialPartNo='" & materialCode & "';"

    'Debug.Print (strSQL)
    rs.Open strSQL, DBCONT          'execute sql query on MaterialTable
    
    If rs.RecordCount > 0 Then
        gauge = rs("Gauge")
        thickness = rs("Thickness")
        quality = rs("BysoftMaterialCode")
    Else
        quality = ""
        gauge = ""
        thickness = 0
        matIsMissing = True
        rs.Close
        Exit Sub
    End If
    rs.Close
End Sub

Public Function getMaterial() As String
    'returns the material for this program
    getMaterial = material
End Function

Public Sub setThickness(thick As String)
    'sets the thickness of the material
    If IsNumeric(thick) Then
        thickness = CDbl(thick)
    End If
    
End Sub

Public Function getThickness() As Double
    'returns the thinkness of the material
    getThickness = thickness
End Function

Public Sub setMake(Optional mk As Variant)
    'sets the make toggle to true if this is a make program
    If Not IsMissing(mk) Then
        If StrComp(mk, "make", vbTextCompare) Then
            make = False
        End If
    Else
        make = True
    End If
    
End Sub

Public Function getMake() As Boolean
    'returns true if the item is a hardwarelot
    'otherwise returns the value in the make toggle
    If harwarelot Then
        getMake = hardwareLot
    Else
        getMake = make
    End If
End Function

Public Sub setWeight(wt As String)
    'sets the weight of the program
    If IsNumeric(wt) Then
        weight = CDbl(wt)
    End If
    
End Sub

Public Function getWeight() As Double
    'returns the weight of this program
    getWeight = weight
End Function

Public Sub setState(st As String)
    'sets the state of the files at the time that this xml was created
    state = st
End Sub

Public Function getState() As String
    'returns the state
    getState = state
End Function

Public Sub setAdd(ad As Boolean)
    'sets the add toggle
    add = ad
End Sub

Public Function getAdd() As Boolean
    'returns the value of the add toggle
    getAdd = add
End Function

Public Sub setSaw(isSaw As Boolean)
    'sets the saw toggle
    saw = isSaw
End Sub

Public Function getSaw() As Boolean
    'reutrns the value of the saw toggle
    getSaw = saw
End Function

Public Sub setLoosePart(loose As Boolean)
    'sets the loosepart toggle
    loosePart = loose
End Sub

Public Function getLoosePart() As Boolean
    'returns the value of the loosepart toggle
    getLoosePart = loosePart
End Function

Public Sub setPunch(pun As Boolean)
    'sets the punch toggle
    punch = pun
End Sub

Public Function getPunch() As Boolean
    'returns the value of the punch toggle
    getPunch = punch
End Function

Public Sub setPowdered(powd As Boolean)
    'sets the powd toggle
    powdered = powd
End Sub

Public Function getPowdered() As Boolean
    'returns the value of the powd toggle
    getPowdered = powdered
End Function

Public Sub setShipLoose(loose As Boolean)
    'sets the shiploose toggle
    shiploose = loose
End Sub

Public Function getShipLoose() As Boolean
    'returns the value of the shiploose toggle
    getShipLoose = shiploose
End Function

Public Sub orderAxis()
    'orders the axis from largest to smallest
    If Me.getXAxis < Me.getYAxis Then
        Dim tmp As Double
        tmp = Me.getXAxis
        Me.setXAxis (Me.getYAxis)
        Me.setYAxis (tmp)
    End If
    
End Sub

Public Function getIsTemplate() As Boolean
    'returns the value of the istemplate toggle
    getIsTemplate = isTemplate
End Function

Public Sub setIsTemplate(template As Boolean)
    'sets the istemplate toggle
    isTemplate = template
End Sub

Public Function getCutTemplate() As Boolean
    'returns the value of the cuttemplate toggle
    getCutTemplate = cutTemplate
End Function

Public Sub setCutTemplate(template As Boolean)
    'sets the cuttemplate toggle
    cutTemplate = template
End Sub

Public Function getHardwareLot() As Boolean
    'returns the value of the hardwarelot toggle
    getHardwareLot = hardwareLot
End Function

Public Sub setHardwareLot(hw As Boolean)
    'sets the hardwarelot toggle
    hardwareLot = hw
End Sub

Public Function getAssemblyCut() As Boolean
    'returns the value of the assemblycut toggle
    getAssemblyCut = assemblyCut
End Function

Public Sub setAssemblyCut(assCut As Boolean)
    'sets the assemblycut toggle
    assemblyCut = assCut
End Sub

Public Function getMatIsMissing() As Boolean
    'return the value of the matismissing toggle
    getMatIsMissing = matIsMissing
End Function

Public Sub setMatIsMissing(isMiss As Boolean)
    'set the matismissing toggle
    matIsMissing = isMiss
End Sub

Public Sub addPart(part As partInfo)
    'adds a part to the partcol collection
    partCol.add item:=part
End Sub

Public Function getParts() As Collection
    'returns partcol
    Set getParts = partCol
End Function

Public Function partCount() As Integer
    'counts the number of parts in the partcol
    If partCol Is Nothing Then
        partCount = 0
    Else
        partCount = partCol.count
    End If
    
End Function

Public Sub setParts(parts As Collection)
    'sets the partcol equal to the given parts collection
    Set partCol = parts
End Sub

Public Function everythingName() As String
    'returns the name of the program
    everythingName = prog
End Function

Public Function setShear(sh As Boolean)
    'sets the shear toggle
    shear = sh
End Function

Public Function getShear() As Boolean
    'returns the value of the shear toggle
    getShear = shear
End Function

Public Sub setExt(extension As String)
    'sets the value of ext (the file extension)
    ext = extension
End Sub

Public Function getExt() As String
    'returns the value of ext (the file extension)
    getExt = ext
End Function

Public Function getExtCut() As Boolean
    'returns true if, based on the extension of the filename alone, the part should be cut
    If StrComp(ext, "psm", vbTextCompare) And StrComp(ext, "sldprt", vbTextCompare) Then
        getExtCut = False
    Else
        getExtCut = True
    End If
End Function

Public Function formList() As String
    'returns the program name if the makeXML variable stays "True"
    Dim makeXML As Boolean
    makeXML = True
    
    If StrComp(prog, "", vbTextCompare) = 0 Then    'check that the program variable is not empty
        makeXML = False
    ElseIf IsNull(rot) Then     'check that the rotation allowance has a value
        rot = 90
    ElseIf IsNull(qty) Then     'check that qty is not empty
        makeXML = False
    ElseIf qty = 0 Then     'check that qty is not zero
        makeXML = False
    ElseIf shear Then
        makeXML = False
    ElseIf assemblyCut Then
        makeXML = False
    ElseIf loosePart Then
        makeXML = False
    ElseIf hardwareLot Then
        makeXML = False
    ElseIf shiploose Then
        makeXML = False
    ElseIf make Then
        If StrComp(Left(prog, 2), "LF", vbTextCompare) = 0 Then
        ElseIf StrComp(Left(prog, 1), "L", vbTextCompare) = 0 Then
        ElseIf StrComp(Left(prog, 1), "F", vbTextCompare) = 0 Then
        Else
            makeXML = False
        End If
    End If
    
    If makeXML Then
        If xAxis <> 0 Then
            formList = prog
        End If
    End If

End Function

Public Function verDims() As Boolean
    'verify that both x and y dimensions have some positive value
    If xAxis = 0 Or yAxis = 0 Then
        verDims = True
    Else
        verDims = False
    End If
End Function

Public Function verOneDim() As Boolean
    'verify that at least one x or y dimensions have a positive value
    If xAxis = 0 And yAxis = 0 Then
        verOneDim = True
    Else
        verOneDim = False
    End If
End Function

Public Function getArea() As Double
    'returns the area of this part
    If xAxis <> 0 And yAxis <> 0 Then
        getArea = xAxis * yAxis
    Else
        getArea = 0
    End If
End Function
