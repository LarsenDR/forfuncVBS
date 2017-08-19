Function qmd(ba As Single, tpa As Single, Optional unittype As String = "imperial") As Double
' Function to calculate quadratic mean diameter from basal area per acre and trees per acre
' by David R. Larsen, Copyright October 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

If (unittype = "imperial") Then
    qmd = Math.Sqr((ba / tpa) / 0.005454154)
ElseIf (unittype = "metric") Then
    qmd = Math.Sqr((ba / tpa) / 0.00007854)
Else
    qmd = 0#
    MsgBox ("Unknown unit type, options are: imperial or metic")
End If

End Function

