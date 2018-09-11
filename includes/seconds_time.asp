<%
    
Function SecondsToTime(ByVal sglScnds)
  Dim hours, minutes, seconds

  ' calculates whole hours (like a div operator)
  hours = sglScnds \ 3600

  ' calculates the remaining number of seconds
  intSeconds = sglScnds Mod 3600

  ' calculates the whole number of minutes in the remaining number of seconds
  minutes = sglScnds \ 60

  ' calculates the remaining number of seconds after taking the number of minutes
  seconds = sglScnds Mod 60

  ' returns as a string
  SecondsToTime = hours & ":" & minutes & ":" & seconds
End Function
%>
