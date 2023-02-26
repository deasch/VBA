Attribute VB_Name = "func_GetNumeric"
'Option Explicit

Function func_GetNumeric(CellRef As String) as Long
  
  Dim StringLength As Integer
  
  StringLength = Len(CellRef)
  
  For i = 1 To StringLength
    If IsNumeric(Mid(CellRef, i, 1)) Then 
      Result = Result & Mid(CellRef, i, 1)
  Next i
  
  func_GetNumeric = Result
End Function
