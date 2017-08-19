Function polartocart(dist As Double, azi As Double, coord As String) As Double
' Function to calcalate the cartesian coordinate from polar coordinates
' by David R. Larsen, Copyright November 7, 2016
Pi = 3.1415926535


If (dist > 0#) Then
  If coord = "x" Then
    polartocart = Math.Sin(azi * Pi / 180#) * dist
    Else
    polartocart = Math.Cos(azi * Pi / 180#) * dist
   End If
End If
    

End Function
