Function sdi(tpa As Single, qmd As Single, Optional unittype as string = "imperial") As Double
' Function to calculate the Stand Density Index as defined by Reineke, 1933
' tpa is trees per acre, qmd is the quadratic mean diameter in inches
' by David R. Larsen, Copyright October, 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/


If (unittype = "imperial") Then
	sdi = tpa * (qmd / 10#) ^ 1.605
ElseIf (unittype = "metric") Then
	sdi = tpa * (qmd / 25.4) ^ 1.605
Else
    sdi = 0#
    MsgBox ("Unknown unit type, options are: imperial or metic")
End If
End Function

