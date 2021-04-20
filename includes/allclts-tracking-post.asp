<%
' ADD/UPDATE BILL TRACKING FOR A SINGLE CLIENT
	If Request.Form("UpdateBillTracking")="True" Then
		ClientID = DeCrypt(Request.Form("UpdateClientID"))
		If ClientID = 0 Then Response.End

		BillNumber = CInt(Request.Form("Bill"))
		If BillNumber < 1000 Or BillNumber > 9999 Then Response.End

		strSQL = _
			"SELECT * FROM [Client Specific Bill Info] " & _
			"WHERE [ClientID]=" & ClientID & " AND [Bill Number]=" & BillNumber
		Set rsBillInfo = Server.CreateObject("ADOR.Recordset")
		rsBillInfo.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic

		If Request.Form("Delete") <> "Delete" Then
			' If bill number doesn't already exist, add it
			If rsBillInfo.EOF Then
				rsBillInfo.AddNew
				rsBillInfo("ClientID")=ClientID
				rsBillInfo("Bill Number")=BillNumber
			End If 
			' Add tracking information for this bill
			Select Case Request.Form("Pri")
				Case "High"  :rsBillInfo("PriorityNum") = 1
				Case "Medium":rsBillInfo("PriorityNum") = 2
				Case "Low"   :rsBillInfo("PriorityNum") = 3
				Case "TBD"   :rsBillInfo("PriorityNum") = 4
			End Select
			Select Case Request.Form("Pos")
				Case "Support" :rsBillInfo("PositionNum") = 1
				Case "Oppose"  :rsBillInfo("PositionNum") = 2
				Case "Concerns":rsBillInfo("PositionNum") = 3
				Case "Neutral" :rsBillInfo("PositionNum") = 4
				Case "Monitor" :rsBillInfo("PositionNum") = 5
				Case ""        :rsBillInfo("PositionNum") = 6
			End Select
			rsBillInfo("Dead") = -CInt(Request.Form("Dead")="True")
			rsBillInfo("Notes") = Trim(Request.Form("Notes"))
			rsBillInfo("Comments") = Request.Form("Comments")
			rsBillInfo.Update
		Else
			' If bill number exists, delete it
			If Not rsBillInfo.EOF Then
				rsBillInfo.Delete
				rsBillInfo.Update
				rsBillInfo.MoveFirst
			End If
		End If
		rsBillInfo.Close
		Set rsBillInfo = Nothing
	Else
		Response.Cookies("LegiTrak")("scrollTop") = ""
	End If
	
' UPDATE MULTIPLE BILL DETAILS
	If Request.Form("UpdateMultipleBills")="True" Then
		Set cxnSQL = CreateObject("ADODB.Connection")
		cxnSQL.Open strConnection
		
		Clients = Split(Request.Form("BillsToUpdate"),";")
		For j = 1 to UBound(Clients)
			Bills = Split(Clients(j),",")
			BillCnt = UBound(Bills)

			strWhere = ""
			For i = 1 to BillCnt ' Start at 1 because 0 is the ClientID
				strWhere = strWhere & "[Bill Number]=" & CInt(Bills(i))
				If i <> BillCnt Then strWhere = strWhere & " OR "
			Next 'i
			strWhere = " WHERE ClientID=" & DeCrypt(Bills(0)) & " AND (" & strWhere & ")"
			
			If Request.Form("Delete") = "Delete" Then
				strSQL = "DELETE FROM [Client Specific Bill Info]" & strWhere
			Else
				strSet = ""
				If Request.Form("Pos") <> "0" Then strSet = ",PositionNum=" & CInt(Request.Form("Pos"))
				If Request.Form("Pri") <> "0" Then strSet = strSet & ",PriorityNum=" & CInt(Request.Form("Pri"))
				If Request.Form("Dead") <> "-1" Then strSet = strSet & ",Dead=" & CInt(Request.Form("Dead"))
				strNotes = TweakQuote(Trim(Request.Form("Notes")))
				If strNotes <> "" Then strSet = strSet & ",Notes='" & strNotes & "'"
				strSQL = "UPDATE [Client Specific Bill Info] SET " & Right(strSet,Len(strSet)-1) & strWhere
			End If
			cxnSQL.Execute strSQL, , adExecuteNoRecords
		Next 'j

		cxnSQL.Close
		Set cxnSQL = Nothing
	End If
%>