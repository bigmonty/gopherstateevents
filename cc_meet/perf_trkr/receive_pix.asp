<%@ Language=VBScript %>

<%
Option Explicit

Dim sql, conn, rs
Dim sFileName, sFileToDelete
Dim sErrMsg
Dim fs

Class FileUploader
	Public  Files
	Private mcolFormElem

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")
		Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(Files) Then
			Files.RemoveAll()
			Set Files = Nothing
		End If
		If IsObject(mcolFormElem) Then
			mcolFormElem.RemoveAll()
			Set mcolFormElem = Nothing
		End If
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
	End Property

	Public Default Sub Upload()
		Dim biData, sInputName
		Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
		Dim nPosFile, nPosBound

		biData = Request.BinaryRead(Request.TotalBytes)
		nPosBegin = 1
		nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
		
		If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
		 
		vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
		nDataBoundPos = InstrB(1, biData, vDataBounds)
		
		Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
			
			nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
			nPos = InstrB(nPos, biData, CByteString("name="))
			nPosBegin = nPos + 6
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
			sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
			nPosBound = InstrB(nPosEnd, biData, vDataBounds)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile, sFileName
				Set oUploadFile = New UploadedFile
				
				nPosBegin = nPosFile + 10
				nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

				nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
				nPosBegin = nPos + 14
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
				
				oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				
				nPosBegin = nPosEnd+4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
				
				If oUploadFile.FileSize > 0 Then Files.Add LCase(sInputName), oUploadFile
			Else
				nPos = InstrB(nPos, biData, CByteString(Chr(13)))
				nPosBegin = nPos + 4
				nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
				If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
			End If

			nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
		Loop
	End Sub

	'String to byte string conversion
	Private Function CByteString(sString)
		Dim nIndex
		For nIndex = 1 to Len(sString)
		   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function CWideString(bsString)
		Dim nIndex
		CWideString =""
		For nIndex = 1 to LenB(bsString)
		   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public FileName
	Public FileData
	
	Public Property Get FileSize()
		FileSize = LenB(FileData)
	End Property

	Public Sub SaveToDisk(sPath)
		Dim oFS, oFile
		Dim nIndex
	
		If sPath = "" Or FileName = "" Then Exit Sub
		If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
	
		Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		If Not oFS.FolderExists(sPath) Then Exit Sub
		
		Set oFile = oFS.CreateTextFile(sPath & FileName, True)
		
		For nIndex = 1 to LenB(FileData)
		    oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
		Next

		oFile.Close
	End Sub
	
	Public Sub SaveToDatabase(ByRef oField)
		If LenB(FileData) = 0 Then Exit Sub
		
		If IsObject(oField) Then
			oField.AppendChunk FileData
		End If
	End Sub
End Class

Dim Uploader, File
Set Uploader = New FileUploader

Uploader.Upload()

If Uploader.Files.Count = 0 Then
	Response.Write "File(s) not uploaded."
Else
	Response.Buffer = true		'Turn buffering on
	Response.Expires = -1		'Page expires immediately
												
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
		
	For Each File In Uploader.Files.Items
		File.FileName = Session("this_user") & Right(File.FileName, 4)
		sFileName = File.FileName
				
		If File.FileSize > 2000000 Then
			sErrMsg = "You may not upload any files greater than 2MB.  The file was not uploaded."
			Exit For
		ElseIf Not (Right(LCase(sFileName), 4) = ".jpg" Or Right(LCase(sFileName), 4) = ".bmp" Or Right(LCase(sFileName), 4) = ".gif" Or Right(LCase(sFileName), 4) = ".png")  Then 
			sErrMsg = "You can only uploads files with .jpg, .bmp, .png, or .gif extensions."
			Exit For
		Else
			'delete existing pix
			Set fs=Server.CreateObject("Scripting.FileSystemObject") 
			If fs.FileExists("c:\inetpub\h51web\gopherstateevents\cc_meet\perf_trkr\images\" & sFileName) = True Then
				fs.DeleteFile("c:\inetpub\h51web\gopherstateevents\cc_meet\perf_trkr\images\" & sFileName)
			End If
			Set fs=Nothing
			
			'now insert this pix
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT Image FROM PerfTrkr WHERE PerfTrkrID = " & Session("this_user")
			rs.Open sql, conn, 1, 2
			rs(0).Value = sFileName
			rs.Update
			rs.Close
			Set rs = Nothing	
			
			File.SaveToDisk "c:\inetpub\h51web\gopherstateevents\cc_meet\perf_trkr\images\"
		End If
	Next
				
	conn.Close
	Set conn = Nothing
	
	If sErrMsg = vbNullString Then
		Session("this_user") = vbNullString
		Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
	Else
		Response.Write(sErrMsg)
	End If
End If
%>