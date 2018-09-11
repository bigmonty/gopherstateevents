<%
Private Function PacePerMile(sTime, iDist, sUnits)
    Dim sglDist
    
    'convert distance to miles
    Select Case sUnits
        Case "miles"
            sglDist = CSng(iDist)
        Case "kms"
            sglDist = CSng(iDist) * 0.6213712
        Case "yds"
            sglDist = CSng(iDist) / 1760
        Case "meters"
            sglDist = CSng(iDist) * 0.000621371192
    End Select
    
    'calculate the pace
    PacePerMile = ConvertToMinutes(Round(CSng(ConvertToSeconds(sTime)) / sglDist, 3))
End Function
%>
