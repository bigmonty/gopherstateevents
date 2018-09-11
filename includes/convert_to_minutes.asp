<%
Private Function ConvertToMinutes(sglScnds)
    Dim iHours, iMinutes, iSeconds, iDecPos, iScnds
    Dim sHours, sMinutes, sSeconds
    Dim sDecimal, strScnds

    If CSng(sglScnds) = Int(CSng(sglScnds)) Then
        sDecimal = "0"
    Else
        'get the decimal part of the time
        strScnds = CStr(sglScnds)
        iDecPos = InStr(strScnds, ".")
        sDecimal = Right(strScnds, Len(strScnds) - iDecPos)
        sDecimal = Left(sDecimal, 1) 'truncate to one decimal place
    End If
    
    'remove decimal from time
    iScnds = Int(CSng(sglScnds))
    
    ' calculates whole hours (like a div operator)
    iHours = iScnds \ 3600

    ' calculates the whole number of minutes in the remaining number of seconds
    iMinutes = (iScnds - (iHours * 3600)) \ 60

    ' calculates the remaining number of seconds after taking the number of minutes
    iSeconds = (iScnds - (iHours * 3600) - (iMinutes * 60))
    
    'convert hours to string
    If iHours > 0 Then sHours = CStr(iHours)

    'convert minutes to string add leading zero to minutes if needed
    If iMinutes < 10 And iHours > 0 Then
        sMinutes = "0" & CStr(iMinutes)
    Else
        sMinutes = CStr(iMinutes)
    End If

    'convert seconds to string and add leading zero to seconds if needed
    If iSeconds < 10 Then
        sSeconds = "0" & CStr(iSeconds)
    Else
        sSeconds = CStr(iSeconds)
    End If

    ' returns as a string
    If iHours = 0 Then
        ConvertToMinutes = sMinutes & ":" & sSeconds & "." & sDecimal
    Else
        ConvertToMinutes = sHours & ":" & sMinutes & ":" & sSeconds & "." & sDecimal
    End If
End Function
%>
