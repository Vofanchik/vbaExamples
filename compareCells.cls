VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Text

Function fileName() As String
fileName = Application.GetOpenFilename(filefilter:="excel files(*.xls*), *.xls*")
End Function

Sub retfilename()
MsgBox (fileName())
'Set wb = Workbooks.Open(fileName())
' Worksheets(1).Activate
' MsgBox (ActiveCell.Value)

'MsgBox (wb.Worksheets(1).Range("B1"))
End Sub

Sub openfile()
Dim filetoopen As String
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim RowNum As String
Dim activerow As String

Dim stateEtapmain As String
Dim stateEtapcompar As String

activerow = ActiveCell.Row

RowNum = ActiveCell.Value

Set ws2 = Workbooks(1).Worksheets(1)
Set ws1 = Workbooks(2).Worksheets(1)

For Each c In ws1.Range("p2:p212").Cells
stateEtapmain = ws2.Range("J" & c.Value).Value
stateEtapcompar = ws1.Range("I" & c.Row).Value
If (stateEtapmain = stateEtapcompar) Then
ws1.Range("q" & c.Row).Value = True
Else: ws1.Range("q" & c.Row).Value = falce
End If
On Error Resume Next
Next

'MsgBox (a)
    
'stateEtapmain = ws2.Range("J" & RowNum).Value
'
'stateEtapcompar = ws1.Range("I" & activerow).Value
'
'MsgBox (stateEtapmain = stateEtapcompar)
'ws2.Range ("H" & RowNum)


End Sub
