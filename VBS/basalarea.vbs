Function basalarea(dbh As Range, w As Range, Optional unittype As String = "imperial") As Double
' Function to calculate basal area per acre from diameter at breast height and expansion factor weights
' by David R. Larsen, Copyright October 10, 2012
' Creative Commons  http://creativecommons.org/licenses/by-nc/3.0/us/



If (unittype = "imperial") Then
    For i = 1 To dbh.Height
         basalarea = basalarea + 0.005454154 * dbh(i).Value ^ 2 * w(i).Value
    Next i
ElseIf (unittype = "metric") Then
    For i = 1 To dbh.Height
         basalarea = basalarea + 0.00007854 * dbh(i).Value ^ 2 * w(i).Value
    Next i
Else
    basalarea = 0#
    MsgBox ("Unknown unit type, options are: imperial or metric")
End If

End Function

