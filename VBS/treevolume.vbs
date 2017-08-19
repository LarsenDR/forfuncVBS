Function volume(dbh As Single, mht As Single, Optional vtype As String = "boardfeet") As Double
'Function to calculate volume from Beers 1964
'By David R. Larsen, Copyright October 8, 2012

  
If (mht > 0#) Then
  a = ((dbh ^ 2 * (dbh + 190#)) / 100000#)
  b = (1# / 100#) * (((mht * (168# - mht)) / 64#) + (32# / mht))
  c = 475# + ((3# * mht ^ 2) / 128#)

  If (vtype = "cords") Then
    volume = a * b
  ElseIf (vtype = "cubic") Then
    volume = a * b * 76
  ElseIf (vtype = "cubicbark") Then
    volume = a * b * 92
  ElseIf (vtype = "boardfeet") Then
    volume = a * b * c
  Else
    volume = 0#
    MsgBox (" vtype must be cords, cubic, cubicbark, or boardfoot")
  End If
Else
  volume = 0#
End If
End Function


