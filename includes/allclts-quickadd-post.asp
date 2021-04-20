<%
' ADD/UPDATE BILL TRACKING FOR ALL CLIENTS
	If Request.Form("QuickAdd") = "True" Then
		BillNumber = CInt("0" & Request.Form("Bill"))
        If BillNumber < 1000 Or BillNumber > 9999 Then Response.End
		CompBill = CInt("0" & Request.Form("Companion"))
        If CompBill <> 0 And (CompBill < 1000 Or CompBill > 9999) Then Response.End
		
		c = Trim(Request.Form("Comments"))
		If Len(c) > 0 Then
			Do Until Asc(Left(c,1)) > 13
				c = Right(c,Len(c)-1)
				If Len(c)=0 Then Exit Do
			Loop
		End If
		If Len(c) > 0 Then
			Do Until Asc(Right(c,1)) > 13
				c = Left(c,Len(c)-1)
				If Len(c)=0 Then Exit Do
			Loop
		End If
		Do Until BillNumber = 0
			strSQLJoin = "[Client Specific Bill Info]"
			strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Client List]  ON [Client Specific Bill Info].[ClientID] = [Client List].[ClientID])"
			strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer Clients] ON [Client List].[ClientID] = [Customer Clients].[ClientID])"
			strSQL = _
				"SELECT [Client Specific Bill Info].*" & _
				" FROM " & strSQLJoin & _
				" WHERE [Customer Clients].CustomerID=" & CustomerID & _
				" ORDER BY [Client Specific Bill Info].[ClientID], [Client Specific Bill Info].[Bill Number]"
			Set rsBillInfo=Server.CreateObject("ADOR.Recordset")
			 ' Use a local cursor so I can use the .Find method starting
			 ' back at the top of the Recordset each time through the loop
			rsBillInfo.CursorLocation = adUseClient
			rsBillInfo.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
			For i = 0 to Request.Form("ClientCount")
				If Request.Form("Clt_" & i) = "True" Then
					ClientID = Decrypt(Request.Form("CltNum_" & i))
					If ClientID = 0 Then Response.End

					' If there are no tracked bills, then we don't need to look for a match
					If Not rsBillInfo.EOF Then
						' Since you can't do a multi-column find using ADO, I have to
						' do two sequential finds
						rsBillInfo.MoveFirst
						rsBillInfo.Find "[ClientID] = " & ClientID
						If Not rsBillInfo.EOF Then rsBillInfo.Find "[Bill Number] = " & BillNumber
						' If we've reach the end of the table, then we need to add a new record
						If rsBillInfo.EOF Then
							rsBillInfo.AddNew
						' If we didn't reach the end, but did move to a different client,
						' then we also need to add a new record
						ElseIf rsBillInfo("ClientID") <> ClientID Then
							rsBillInfo.AddNew
						End If
					Else
						rsBillInfo.AddNew
					End If

					intPri = CInt(Request.Form("Pri_" & i))
					If intPri < 1 Or intPri > 4 Then
					    intPri = 1
					End If
					intPos = CInt(Request.Form("Pos_" & i))
					If intPos < 1 Or intPos > 6 Then
					    intPri = 1
					End If

					' Add tracking information for this bill
					rsBillInfo("ClientID") = ClientID
					rsBillInfo("Bill Number") = BillNumber
					rsBillInfo("PriorityNum") = intPri
					rsBillInfo("PositionNum") = intPos
					n = Trim(Request.Form("Notes_" & i))
					If Len(n) <> 0 Then rsBillInfo("Notes") = n
					If Len(c) <> 0 Then rsBillInfo("Comments") = c
					rsBillInfo.Update
				End If
			Next
			rsBillInfo.Close
			Set rsBillInfo = Nothing

			' Repeat loop for companion bill if selected
			If CompBill <> 0 Then
				BillNumber = CompBill
				CompBill = 0
			Else
				BillNumber = 0
			End If

		Loop
		BillNumber = ""
	End If
%>