<%
Private Function PacePerKM(sTime, iDist, sUnits)
    Dim sglDist
    
    'convert distance to kms
    Select Case sUnits
        Case "miles"
            sglDist = CSng(iDist) * 1.609344
        Case "kms"
            sglDist = CSng(iDist)
        Case "yds"
            sglDist = CSng(iDist) * 2832.4454
        Case "meters"
            sglDist = CSng(iDist) / 1000
    End Select
    
    'calculate the pace
    PacePerKM = ConvertToMinutes(Round(CSng(ConvertToSeconds(sTime)) / sglDist, 2))
End Function
%>
