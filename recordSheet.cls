VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "recordSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private dict As Object
Private row As Integer
Private sheetName As String
Private rows As Integer

Private Sub Class_Initialize()
    'initializes the dictionary to and empty set
    'sets row equal to 1
    Set dict = CreateObject("Scripting.Dictionary")
    row = 1
End Sub

Private Sub add(k As String, value As Integer)
    'adds items to the dictionary as a key value pair
    dict.add key:=k, item:=value
End Sub

Private Function colNum(k As String) As Integer
    'finds the column number for a particular field
    If dict.exists(k) Then
        colNum = dict(k)
    Else
        colNum = 0
    End If
End Function

Private Sub moveNext()
    'moves to the next row
    row = row + 1
End Sub

Public Function field(k As String) As Variant
    'returns the value of the given field at the current row
    If dict.exists(k) Then
        field = Worksheets(sheetName).Range(dict(k) & row).value
        If IsEmpty(field) Then
            field = "  "
        End If
    Else
        field = "  "
    End If
End Function

Public Sub setsheet(sheet As String)
    'initializes the sheet name and sets the key value pairs for the fields in the dictionary
    sheetName = sheet
    Dim colIndex As Integer
    
    For colIndex = 1 To Worksheets(sheetName).Cells(1, Columns.count).End(xlToLeft).column
        Dim index As String
        Dim vArr
        vArr = Split(Cells(1, colIndex).Address(True, False), "$")
        index = Worksheets(sheetName).Range(vArr(0) & 1).value
        dict(index) = vArr(0)
    Next
    rows = Worksheets(sheetName).Cells(Worksheets(sheetName).rows.count, 1).End(xlUp).row
End Sub

Public Function hasNext() As Boolean
    'checks if there are more rows to process and moves to the next row if there is
    If row >= rows Then
        hasNext = False
    Else
        hasNext = True
        Call moveNext
    End If
End Function

Public Function emptyQuery() As Boolean
    'checks that there was at least one entry returned by the query
    If noValues() Then
        emptyQuery = True
    Else
        emptyQuery = False
    End If
End Function

Private Function noValues() As Boolean
    'check the first row for values, returns true if there are values present
    Dim hasValue As Boolean, key As Variant
    hasValue = False
    For Each key In dict
        If StrComp(Worksheets(sheetName).Range(dict(key) & row + 1), "", vbTextCompare) <> 0 Then
            hasValue = True
        End If
    Next
    noValues = hasValue
End Function

Public Function entryExists(myfield As String, value As String) As Boolean
    'this function looks through all the remaining entries looking for the given value in the given field
    Do While hasNext()
        Dim tmpValue As String
        tmpValue = field(myfield)
        If StrComp(tmpValue, value, vbTextCompare) = 0 Then
            entryExists = True
            row = 1
            Exit Function
        End If
    
    Loop
    entryExists = False
    row = 1
End Function

Public Sub resetRow()
    'resets the row counter to 1
    row = 1
End Sub

Public Function getInfo(from As String, where As String, isa As String) As String
    'looks for the value in "isa" in the column "where",
    'when the value is found this function returns current row`s value in the "from" field
    'if nothing is found it returns zero
    Dim tmpValue As Double
    tmpValue = 0
    Do While hasNext()
        If StrComp(field(where), isa, vbTextCompare) = 0 Then
            tmpValue = CDbl(field(from))
            Exit Do
        End If
    Loop
    row = 1
    getInfo = tmpValue
End Function

Public Function goToRow(searchCriteria As String, column As String) As Boolean
    row = 2
    Do While StrComp(searchCriteria, CStr(field(column)), vbTextCompare)
        If Not hasNext Then
            goToRow = False
            Exit Function
        End If
    Loop
    
    goToRow = True
End Function
