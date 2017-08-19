Function logVolume(sdia As Single, ldia As Single, length As Single, Optional equationtype As String = "smalian", Optional unittype As String = "imperial") As Double
' Function to calculate the Cubic volume of a log.
' You must specify both unittype and equationtype.
' by David R. Larsen, Copyright October 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

Const PI = 3.14159265358979
If (unittype = "none") Then
    sarea = PI / 4# * (sdia) ^ 2
    larea = PI / 4# * (ldia) ^ 2
ElseIf (unittype = "imperial") Then
    sarea = PI / 4# * (sdia / 12#) ^ 2
    larea = PI / 4# * (ldia / 12#) ^ 2
ElseIf (unittype = "metric") Then
    sarea = PI / 4# * (sdia / 100#) ^ 2
    larea = PI / 4# * (ldia / 100#) ^ 2
Else
    logVolume = 0#
    MsgBox ("Unknown unittype, Options are: none, imperial or metric")
End If

If (equationtype = "smalian") Then
    logVolume = length / 2# * (sarea + larea)
ElseIf (equationtype = "cone") Then
    logVolume = length / 3# * (sarea + Math.Sqr(sarea * larea) + larea)
ElseIf (equationtype = "neloid") Then
    logVolume = length / 4# * (sarea + (sarea ^ 2 * larea) ^ (1 / 3) + (sarea * larea ^ 2) ^ (1 / 3) + larea)
Else
    logVolume = 0#
    MsgBox ("Unknown equationtype, Options are: smalian, cone or neloid")
End If


End Function

