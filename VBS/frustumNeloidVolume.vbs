Function frustumNeloidVolume(sdia As Single, ldia As Single, length As Single) As Double
' Function to calculate the Smailian volume of a log.
' diameter and length must be in the same units.
' by David R. Larsen, Copyright October 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

Const PI = 3.14159265358979
sarea = PI / 4 * sdia * sdia
larea = PI / 4 * ldia * ldia

frustumNeloidVolume = length / 4# * (sarea + (sarea ^ 2 * larea) ^ (1 / 3) + (sarea * larea ^ 2) ^ (1 / 3) + larea)

End Function

