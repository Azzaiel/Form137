Attribute VB_Name = "mod_grade"
Option Explicit
Public Function getRemark(grade As Double, Optional isKinder As Boolean = False)
  If (isKinder) Then
  
    If (grade >= 90) Then
      getRemark = "O"
    ElseIf (grade >= 85) Then
      getRemark = "VS"
    ElseIf (grade >= 80) Then
      getRemark = "S"
    ElseIf (grade >= 75) Then
      getRemark = "I"
    Else
      getRemark = "NI"
    End If
  
  Else
   If (grade >= 90) Then
      getRemark = "A"
    ElseIf (grade >= 85) Then
      getRemark = "P"
    ElseIf (grade >= 80) Then
      getRemark = "AP"
    ElseIf (grade >= 75) Then
      getRemark = "D"
    Else
      getRemark = "B"
    End If
  End If
End Function

Public Function getFinalGrade(grades() As Double) As Integer
  Const arraySize As Integer = 3
  Dim index As Integer
  Dim divider As Integer
  Dim totalGrade As Integer
  divider = 0
  
  For index = 0 To arraySize
    If (grades(index) > 0) Then
      totalGrade = totalGrade + grades(index)
      divider = divider + 1
    End If
  Next index
  
  If (divider >= 1) Then
   getFinalGrade = totalGrade / divider
  Else
   getFinalGrade = 0
  End If
  
End Function


