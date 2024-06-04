Attribute VB_Name = "Module1"
Sub InsertCell()

Dim SourceWorkbook As Workbook
Dim TargetWorkbook As Workbook
Dim SourceSheet As Worksheet
Dim TargetSheet As Worksheet
Dim CellToCopy As Range
Dim CellToPaste As Range
Dim active As Range
Dim arrSplitStrings1() As String
Dim SplitStrings1 As String


Set SourceWorkbook = Workbooks(2)
Set TargetWorkbook = Workbooks(1)
Set SourceSheet = SourceWorkbook.Worksheets(1)
Set TargetSheet = TargetWorkbook.Worksheets(1)


Set CellToCopy = ActiveCell

Set CellToPaste = SourceSheet.Cells(CellToCopy.Row, 16)

arrSplitStrings1 = Split(CellToPaste.Value, "!")
SplitStrings1 = arrSplitStrings1(1)

SplitStrings1 = Replace(SplitStrings1, "#", "")
SplitStrings1 = Replace(SplitStrings1, "H", "C")
Set CellToPaste = TargetSheet.Range(SplitStrings1)

'Set CellToPaste = TargetSheet.Cells(TargetSheet.Rows.Count).End(xlUp).Offset(1, 0)
CellToCopy.Copy
CellToPaste.PasteSpecial xlPasteValues
'
Application.CutCopyMode = False
'SourceWorkbook.Close SaveChanges:=True


'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
End Sub
