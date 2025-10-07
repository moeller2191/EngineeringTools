VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Sub Main()

    Call connectDatabase                        'create database connection
    
    Dim i As Integer
    For i = 1 To Worksheets("CheckedSOs").Cells(rows.count, "A").End(xlUp).row
        If StrComp(Worksheets("CheckedSOs").Range("A" & i), "", vbTextCompare) Then
            Call insertSO(Worksheets("CheckedSOs").Range("A" & i))
            Worksheets("CheckedSOs").Range("A" & i) = ""
        End If
    Next

    Call closeDatabase                      'close database and record sets
End Sub
