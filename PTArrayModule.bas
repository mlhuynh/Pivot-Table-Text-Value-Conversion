'---
'PURPOSE: VBA module to assign designated arrays to numerical values within Excel pivot tables
'AUTHOR: Linh Huynh
'DATE: January 17, 2022
'DIRECTIONS: 
'1. Manually assign your arrays list (i.e. words or letter) to each corresponding numerical value within the AnimalGroup section
'2. Select any cell within the pivot table, then run this macro.
'---

Sub ApplyPTArrays()

Dim PTRange As String
Dim pt As PivotTable
Dim FirstCell As String
Dim ColonLocation As Integer
Dim i As Long
Dim AnimalGroup As Variant
Dim AnimalNumber As Variant
Dim CurrentName As String
Dim Quote As String

'Define arrays lists
AnimalGroup = _
  Array("Amphibian", "Mammal", "Fish", "Bird", "Reptile")
AnimalNumber = Array(1, 2, 3, 4, 5)

'Returns a string containing the character associated with the specified character code. In this case, Chr(34) is quotation marks.
Quote = Chr(34)

On Error Resume Next
Set pt = ActiveCell.PivotTable
On Error GoTo 0
If pt Is Nothing Then
  MsgBox "Please select a pivot table" _
     & vbCrLf _
     & "Please select any cell within your pivot table and try again"
  Exit Sub
End If

'Find the location of the top left cell of the pivot table's DataBodyRange
PTRange = pt.DataBodyRange.Address
ColonLocation = InStr(PTRange, ":")
FirstCell = Left(PTRange, ColonLocation - 1)
FirstCell = Replace(FirstCell, "$", "")
  
'Start the conditional format rule
FirstCell = "=" & FirstCell & "="

'Set up arrays
  ReDim Preserve AnimalGroup(1 _
    To UBound(AnimalGroup) + 1)
  ReDim Preserve AnimalNumber(1 _
    To UBound(AnimalNumber) + 1)
  
'Iterate through array list to convert assigned values to associated text
For i = 1 To 5
  CurrentName = "[=" _
    & AnimalNumber(i) & "]" _
        & Quote & AnimalGroup(i) _
        & Quote & ";;"
		
  With Range(PTRange).FormatConditions _
    .Add(Type:=xlExpression, _
      Formula1:=FirstCell _
      & AnimalNumber(i))
    .NumberFormat = CurrentName
  End With
Next i

End Sub