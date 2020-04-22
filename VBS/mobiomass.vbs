Function mobiomass(dbh As Single, mht As Single, Optional Species As String = "black")
' Function to calculate the biomass of a tree using Goerndt, 2016
' by David R. Larsen, Copyright April 22, 2020
mobio = 0
  If ( mht > 0 ){
   If( species == "black" )  Then
	mobio = 1.67079 + 0.04796 * dbh^2 + 0.81549 * log((dbh^2)*mht)
   ElseIf( species == "post" ) Then
	mobio = -0.50714 + 0.01655 * dbh^2 + 0.81549 * Log((dbh^2)*mht)
   ElseIf( species == "hickory" ) Then
	mobio = 0.70177 + 0.05791 * dbh^2 + 0.60755 * Log((dbh^2)*mht)
   ElseIf( species == "white" ) Then
	mobio = 0.61557 + 0.00373 * dbh^2 + 0.71159 * Log((dbh^2)*mht)
   Else
	mobio = 0
    MsgBox (" Species must be black, post, hickory, white!")
   End If
  }
  mobiomass = mobio
End Function
