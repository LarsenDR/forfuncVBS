Function mesavage(dbh As Single, mht As Single, Optional volumetype As String = "Doyle", Optional girard As Single = 78#)
' Function to calculate the Mesavage and Girard 1946 volume. using the equations by H.V. Wiant, Jr., 1986, 
' Formula’s for Mesavage and Girard’s Volume Tables, Northern Journal of Applied Forestry 3:124.
' Coded by David R. Larsen, June 20, 2015
L = mht / 16#
cor = (1# + ((girard - 78#) * 0.03))

If (volumetype = "Int1/4") Then
  a = Array(-13.35212, 9.58615, 1.52968)
  b = Array(1.79620, -2.59995, -0.27465)
  c = Array(0.04482, 0.45997, -0.00961)
ElseIf (volumetype = "Scribner") Then
  a = Array(-22.50365, 17.53508, -0.59242)
  b = Array(3.02888, -4.34381, -0.02302)
  c = Array(-0.01969, 0.51593, -0.02035)
ElseIf (volumetype = "Doyle") Then
  a = Array(-29.37337, 41.51275, 0.55743)
  b = Array(2.78043, -8.77272, -0.04516)
  c = Array(0.04177, 0.59042,  -0.01578)
Else
  mesavage = 0
  MsgBox (" volumetype must be Int1/4, Scribner, or Doyle")
End If

v1 = (a(0) + a(1) * L + a(2) * L^2) 
v2 = (b(0) + b(1) * L + b(2) * L^2) * dbh 
v3 = (c(0) + c(1) * L + c(2) * L^2) * dbh^2 
mesavage = (v1 + v2 + v3) * cor
End Function


