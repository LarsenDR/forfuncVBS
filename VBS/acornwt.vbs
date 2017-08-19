Function acornwt(dbh As Double, species As String) As Double
' Function to calculate mean acorn weight (in lb) from dbh and species and trees per acre
' by David R. Larsen, Copyright October 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

If (species = "black" And dbh > 9.9 And dbh < 36.1) Then
   acornwt = -1.9065 + 0.2973 * dbh
ElseIf (species = "chestnut" And dbh > 9.9 And dbh < 36.1) Then
   acornwt = 0.0008271 * dbh ^ 3 - 0.08157 * dbh ^ 2 + 2.692 * dbh - 18.85
ElseIf (species = "red" And dbh > 9.9 And dbh < 36.1) Then
    acornwt = 0.0004016 * dbh ^ 4 - 0.0349937 * dbh ^ 3 + 0.9864357 * dbh ^ 2 - 9.5233885 * dbh + 27.32720516
ElseIf (species = "scarlet" And dbh > 9.9 And dbh < 36.1) Then
   acornwt = 0.0005975 * dbh ^ 4 - 0.05325 * dbh ^ 3 + 1.651 * dbh ^ 2 - 19.97 * dbh + 84.71
ElseIf (species = "white" And dbh > 9.9 And dbh < 36.1) Then
   acornwt = 0.0001987 * dbh ^ 4 - 0.02027 * dbh ^ 3 + 0.694 * dbh ^ 2 - 8.768 * dbh + 37.36
Else
   acornwt = Empty
End If


End Function

