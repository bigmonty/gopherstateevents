<%
Private Function ConvertToSeconds(sTime)
    Dim sSubStr(3), Count, j
    Dim sglSeconds(3), k

    'find out how many substrings are needed
    If sTime & "" = "" Then
		ConvertToSeconds = 0
    Else
		For k = 1 To Len(sTime)
		    If Mid(sTime, k, 1) = ":" Then Count = Count + 1
		Next
    
		'break the time into substrings
		For k = 1 To Len(sTime)
		    If Mid(sTime, k, 1) = ":" Then
		        j = j + 1
		    Else
		        sSubStr(j) = sSubStr(j) & Mid(sTime, k, 1)
		    End If
		Next
    
		'do the conversion
		For k = 0 To Count
		    j = Count - k
		    If sSubStr(k) = vbNullString Then
		        sglSeconds(k) = 0
		    Else
                sglSeconds(k) = CSng(sSubStr(k)) * (60 ^ j)
		    End If
		    ConvertToSeconds = ConvertToSeconds + sglSeconds(k)
		Next
	End If
End Function
%>