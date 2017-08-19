Function simpleTaper(h As Double, dbh As Double, ht As Double, htcb As Double, stumpR As Double, stemR As Double, bh As Double) As Double

  ' Function to calculate a simple taper equation
  ' Copyright by David R. Larsen, August 14, 2012
  '
  diameterCrownBase = dbh + stemR * (htcb - bh)
  crownLength = ht - htcb
  topRate = diameterCrownBase / crownLength
  
  simpleTaper = 0#
    
  If (h < bh) Then
     simpleTaper = dbh + stumpR * (h - bh)
  ElseIf ((h >= bh) And (h < htcb)) Then
     simpleTaper = dbh + stemR * (h - bh)
  Else
     simpleTaper = (ht - h) * topRate
  End If
  
End Function

