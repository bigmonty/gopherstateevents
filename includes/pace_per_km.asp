<%
Private Function PacePerKM(sTime, sThisDist)
    Dim sglDist

    If UCase(Right(sThisDist, 2)) = "MI" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3)) * 1.609344
    ElseIf UCase(Right(sThisDist, 4)) = "MILE" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 5)) * 1.609344
    ElseIf UCase(sThisDist) = "MARATHON" Then
        sglDist = CSng(26.2) * 1.609344
    ElseIf UCase(sThisDist) = "H. MAR" Then
        sglDist = CSng(13.1) * 1.609344
    Else
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3))
    End If
    
    'calculate the pace
    PacePerKM = ConvertToMinutes(Round(CSng(ConvertToSeconds(Round(sTime, 2))) / Round(sglDist, 2), 3))
End Function
%>