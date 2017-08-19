Function frustumConeVolume(sdia As Single, ldia As Single, length As Single) As Double
' Function to calculate the Smailian volume of a log.
' diameter and length must be in the same units.
' by David R. Larsen, Copyright October 9, 2012
' Creative Commons http://creativecommons.org/licenses/by-nc/3.0/us/

Const PI = 3.14159265358979
sarea = PI / 4 * sdia * sdia
larea = PI / 4 * ldia * ldia

frustumConeVolume = length / 3# * (sarea + Math.Sqr(sarea * larea) + larea)

End Function

