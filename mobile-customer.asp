<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	' These variables come from the sign-on form
	CustomerUsername = TweakQuote(Request.Form("CustomerUsername"))
	CustomerPassword = TweakQuote(Request.Form("CustomerPassword"))

	' If we've got a username, try to log in
	If Len(CustomerUsername) <> 0 Then
		strSQL= _
			"SELECT * FROM [Customer List] C INNER JOIN [Organization List] O" & _
			" ON C.OrganizationID=O.OrganizationID " & _
			"WHERE" & _
			" C.[Username]='" & CustomerUsername & "' AND" & _
			" C.[Password]='" & CustomerPassword & "' AND" & _
			" O.[Billing Clients] > 0"
		Set rsCustomer=Server.CreateObject("ADOR.Recordset")
		rsCustomer.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		
		SignonOK = True
		If rsCustomer.EOF Then
			SignonOK = False ' If the record wasn't found, then bad password
		Else
			' The SQL call is case-insensitive, so we check that here
			If CustomerPassword <> rsCustomer("Password") Then
				SignonOK = False
			End If
		End If

		If Not SignonOK Then
			Response.Redirect "errors/mobile-signon-error.htm"
		Else
			CustomerID = rsCustomer("CustomerID")
			ClientID = 0
			Response.Cookies("LegiTrak")("CustomerID") = Encrypt(CustomerID)
			CustomerName = rsCustomer("Customer Company Name")
			Response.Cookies("LegiTrak")("CustomerName") = CustomerName
			rsCustomer("In Use") = 1
			rsCustomer("Session Started") = Now
			rsCustomer("IP Address") = Request.ServerVariables("REMOTE_ADDR")
			rsCustomer.Update
		End If
		rsCustomer.Close
		Set rsCustomer = Nothing
	Else
%>
<!--#include virtual="includes/security-check.asp"-->
<%
		If Request.Form("ChangeClient") = "Change Client" Then
			Response.Cookies("LegiTrak")("ClientID") = Request.Form("ClientID")
			ClientID = Decrypt(Request.Form("ClientID"))
		End If
	End If
%>
<html>
<head>
<meta name=HandheldFriendly content=true>
<meta name=PalmComputingPlatform content=true>
<link rel=stylesheet href="mobile-styles.css" type="text/css">
</head>
<body>
<b>LegiTrak</b> <i>Mobile!</i><br>
<b><%=CustomerName%></b><br><br>
<table width=153 cellspacing=0 cellpadding=0 border=0>
<tr><td><i>Tracking List</i></td>
<td><a href='mobile-calendar.asp'>Calendar</a></td></tr>
<tr><td><a href='mobile-quickadd.asp'>Quick Add</a></td>
<td><a href='mobile-votecards.asp'>Vote Cards</a></td></tr>
<tr><td colspan=2><a href='mobile.asp'>Logout</a></td></tr>
</table><br>
<form action="mobile-customer.asp" method=post>
<select name=ClientID style='width:150;font-weight:bold'>
<%
' CLIENT LIST
	strSQL = _
		"SELECT [Client List].*" & _
		" FROM [Customer Clients] INNER JOIN [Client List]" & _
		" ON [Customer Clients].ClientID = [Client List].ClientID" & _
		" WHERE [Customer Clients].CustomerID=" & CustomerID & _
		" ORDER BY [Short Company Name]"
	set rsClients=Server.CreateObject("ador.Recordset")
	rsClients.Open strSQL, strConnReadOnly
	If ClientID = 0 Then
		ClientID = rsClients("ClientID")
		Response.Cookies("LegiTrak")("ClientID") = Encrypt(ClientID)
	End If
	Do Until rsClients.EOF
		Response.Write "<option value=" & Encrypt(rsClients("ClientID"))
		If ClientID = rsClients("ClientID") Then
			Response.Write " selected"
		End If
		Response.Write ">" & rsClients("Short Company Name")
		rsClients.MoveNext
	Loop
	rsClients.Close
	set rsClients = Nothing
%>
</select><br>
<input type=submit name=ChangeClient value='Change Client'>
</form>
<%
' BILL TRACKING
	strSQLJoin = "[Client Specific Bill Info] CS"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Client List] CL ON CS.[ClientID] = CL.[ClientID])"
	strSQLJoin = "(" & strSQLJoin & "  LEFT JOIN [Daily Status] DS ON CS.[Bill Number] = DS.[Bill Number])"
	strSQL = "SELECT" & _
		" CL.ClientID, CL.[Client Company Name]," & _
		" CS.[Bill Number], CS.[PositionNum], CS.[Dead]," & _
		" DS.Title, DS.House, DS.Location, DS.Action" & _
		" FROM " & strSQLJoin & _
		" WHERE CL.ClientID=" & ClientID & _
		" ORDER BY CL.[Client Company Name], CS.[Bill Number]"
	set rsBillInfo=Server.CreateObject("ador.Recordset")
	rsBillInfo.Open strSQL, strConnReadOnly

	Response.Write "<table width=153 cellspacing=0 cellpadding=0 border=0>"
	Response.Write "<tr><td colspan=2><b>" & rsBillInfo("Client Company Name") & "</b></td></tr>"

	Do Until rsBillInfo.EOF
		If Trim(rsBillInfo("Title")) <> "" Then
			strTitle = rsBillInfo("Title")
		Else
			strTitle = "(No Title Available)"
		End If
		If Trim(rsBillInfo("House")) <> "" Then
			strHouseLoc = rsBillInfo("House") & ", " & rsBillInfo("Location")
		ElseIf Trim(rsBillInfo("Location")) <> "" Then
			strHouseLoc = rsBillInfo("Location")
		Else
			strHouseLoc = ""
		End If
		If rsBillInfo("Dead") = "True" Then strHouseLoc = strHouseLoc & " (Dead)"
		Response.Write "<tr>"
		Response.Write "<td valign=top><a href='mobile-edit.asp?" & _
			"bill=" & rsBillInfo("Bill Number") & "'>"
		Response.Write rsBillInfo("Bill Number") & "</a>&nbsp; </td>"
		Response.Write "<td>" & strTitle & "</td>"
		Response.Write "</tr><tr><td></td>"
		Response.Write "<td>" & strHouseLoc & "</td>"
		Response.Write "</tr>"
		rsBillInfo.MoveNext
	Loop
	Response.Write "</table>"
	rsBillInfo.Close
	set rsBillInfo = Nothing
%>    
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>