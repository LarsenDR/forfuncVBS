Function percentStocking(ba As Single, tpa As Single, Optional foresttype As String = "upland.oak") As Double
' Function to calculate quadratic mean diameter from basal area per acre and trees per acre
' by David R. Larsen, Copyright November 16, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/


dq = Math.Sqr((ba / tpa) / 0.005454154)
amd = (-0.259 + (0.973 * dq))
If (foresttype = "upland.oak") Then
  percentStocking = (tpa * (-0.00507 + 0.01698 * amd + 0.00307 * dq ^ 2))
ElseIf (foresttype = "northern.red.oak") Then
  percentStocking = (tpa * (0.02476 + 0.004182 * amd + 0.00267 * dq ^ 2))
ElseIf (foresttype = "sugar.maple") Then
  percentStocking = (tpa * (-0.003082 + 0.006272 * amd + 0.00469 * dq ^ 2))
ElseIf (foresttype = "black.cherry") Then
  percentStocking = (tpa * (0.02794 + 0.01545 * amd + 0.000871 * dq ^ 2))
ElseIf (foresttype = "red.maple") Then
  percentStocking = (tpa * (-0.01798 + 0.02143 * amd + 0.001711 * dq ^ 2))
ElseIf (foresttype = "black.walnut") Then
  percentStocking = (tpa * (0.01646 + 0.01347 * amd + 0.002757 * dq ^ 2))
ElseIf (foresttype = "shortleaf.pine") Then
  percentStocking = (tpa * (0.008798 + 0.009435 * amd + 0.00253 * dq ^ 2))
ElseIf (foresttype = "cottonwood.silver.maple") Then
  percentStocking = (tpa * (-0.0685724 + 0.0010125 * amd + 0.0023656 * dq ^ 2))
Else
  percentStocking = 0#
  MsgBox ("Unknown forest type, options are: upland.oak, northern.red.oak, sugar.maple, black.cherry, red.maple, black.walnut, shortleaf.pine, cottonwood.silver.maple")
End If
If (percentStocking < 0#) Then
  percentStocking = 0#
End If

End Function


