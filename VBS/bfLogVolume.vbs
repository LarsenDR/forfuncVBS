Function bfLogVolume(sdia As Single, length As Single, Optional voltype As String = "int14") As Double
' Function to calculate the Doyle, scribner and International 1/4 board foot volume of a log
' sdia is in inches and length is in feet
' by David R. Larsen, Copyright October 9, 201
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

If (voltype = "doyle") Then
    bfLogVolume = ((sdia - 4) / 4) ^ 2 * length
ElseIf (voltype = "scribner") Then
    bfLogVolume = (0.79 * sdia ^ 2 - 2# * sdia - 4#) * (length / 16#)
ElseIf (voltype = "int14") Then
    If (length = 4#) Then
        bfLogVolume = 0.22 * sdia ^ 2 - 0.71 * sdia
    ElseIf (length = 8#) Then
        bfLogVolume = 0.44 * sdia ^ 2 - 1.2 * sdia - 0.3
    ElseIf (length = 12#) Then
        bfLogVolume = 0.66 * sdia ^ 2 - 1.47 * sdia - 0.79
    ElseIf (length = 16#) Then
        bfLogVolume = 0.88 * sdia ^ 2 - 1.56 * sdia - 1.36
    ElseIf (length = 20#) Then
        bfLogVolume = 1.1 * sdia ^ 2 - 1.35 * sdia - 1.9
    ElseIf (length = 24#) Then
        bfLogVolume = 1.1 * sdia ^ 2 - 1.35 * sdia - 1.9 + 0.22 * sdia ^ 2 - 0.71 * sdia
    ElseIf (length = 28#) Then
        bfLogVolume = 1.1 * sdia ^ 2 - 1.35 * sdia - 1.9 +  0.44 * sdia ^ 2 - 1.2 * sdia - 0.3
    ElseIf (length = 32#) Then
        bfLogVolume = 1.1 * sdia ^ 2 - 1.35 * sdia - 1.9 + 0.66 * sdia ^ 2 - 1.47 * sdia - 0.79 
    ElseIf (length = 36#) Then
        bfLogVolume = (0.88 * sdia ^ 2 - 1.56 * sdia - 1.36) * 2
    ElseIf (length = 40#) Then
        bfLogVolume = (1.1 * sdia ^ 2 - 1.35 * sdia - 1.9 ) * 2
    Else
        bfLogVolume = 0
        MsgBox ("Unknown log length, options are: 4, 8, 12, 16, 20")
    End If
    
Else
    bfLogVolume = 0
    MsgBox ("Unknown voltype, options are: doyle, scribner, or int14")
End If

End Function

