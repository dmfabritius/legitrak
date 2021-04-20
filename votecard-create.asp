<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "votecard-summary.asp"
	End If

	VotecardID = Decrypt(Request.Form("VotecardID"))
	If VotecardID = 0 Then VotecardID = Decrypt(Request.Cookies("LegiTrak")("VotecardID"))

	If Request.Form("DeleteEntry") = "True" Then
' DELETE VOTE CARD
		strCommand = _
			"DELETE FROM [Votecards] " & _
			"WHERE VotecardID=" & VotecardID
		Set cxnSQL = CreateObject("ADODB.Connection")
		With cxnSQL
			.Open strConnection
			.Execute strCommand, , adExecuteNoRecords
			.Close
		End With
		Set cxnSQL = Nothing
		Response.Redirect "votecard-summary.asp"
	End If
	
	If Request.Form("UpdateEntry") = "Submit" Then

		If Request.Form("Chamber") = "H" Then
		    strChamber = "H"
		Else
		    strChamber = "S"
		End If
        strDesc = TweakQuote(Trim(Request.Form("Desc")))

		If  VotecardID = 0 Then

' ADD NEW VOTE CARD
			Set cxnSQL = CreateObject("ADODB.Connection")
			With cxnSQL
				.Open strConnection

		' Determine next Votecard ID
				strCommand = "SELECT MAX(VotecardID) AS MaxID FROM [VoteCards]"
				Set rsResult = .Execute(strCommand)
				If IsNull(rsResult("MaxID")) Then
					VotecardID = 1
				Else
					VotecardID = rsResult("MaxID") + 1
				End If

		' Add Votecard record
				strCommand = _
					"INSERT INTO [Votecards]" & _
					" (VotecardID, Owner, Description, Chamber) VALUES (" & _
					VotecardID & "," & _
					CustomerID & "," & _
					"'" & strDesc & "'," & _
					"'" & strChamber & "')"

				.Execute strCommand, , adExecuteNoRecords

		' Add votecard customer associations
				strCommand = "INSERT INTO [Customer Votecards] VALUES (" & CustomerID & "," & VotecardID & ")"
				.Execute strCommand, , adExecuteNoRecords
				For i = 1 to Request.Form("CustCount")
				    ID = Decrypt(Request.Form("c" & i))
					If ID <> 0 Then
						strCommand = "INSERT INTO [Customer Votecards] VALUES (" & ID & "," & VotecardID & ")"
						.Execute strCommand, , adExecuteNoRecords
					End If
				Next ' i

				.Close
			End With
			Set cxnSQL = Nothing
			Response.Redirect "votecard-summary.asp"

		Else

' UPDATE AN EXISTING VOTE CARD
			Set cxnSQL = CreateObject("ADODB.Connection")
			With cxnSQL
				.Open strConnection

		' Update Votecard record
				strCommand = _
					"UPDATE Votecards SET" & _
					" Description='" & strDesc & "', " & _
					" Chamber='" & strChamber & "' " & _
					"WHERE" & _
					" VotecardID=" & VotecardID
					.Execute strCommand, , adExecuteNoRecords

		' Add votecard customer associations
				strCommand = _
					"DELETE FROM [Customer Votecards] " & _
					"WHERE VotecardID=" & VotecardID & " AND CustomerID <> " & CustomerID
				.Execute strCommand, , adExecuteNoRecords
				For i = 1 to Request.Form("CustCount")
				    ID = Decrypt(Request.Form("c" & i))
					If ID <> 0 Then
						strCommand = "INSERT INTO [Customer Votecards] VALUES (" & ID & "," & VotecardID & ")"
						.Execute strCommand, , adExecuteNoRecords
					End If
				Next ' i

				.Close
			End With
			Set cxnSQL = Nothing
			Response.Redirect "votecard-summary.asp"
		End If
	Else
	
		If Request.Cookies("LegiTrak")("ModifyVotecard") = "True" Then
			bolModifyVotecard = True
			Response.Cookies("LegiTrak")("ModifyVotecard") = ""

' LOAD AN EXISTING VOTE CARD'S INFORMATION
			VotecardID = Decrypt(Request.Cookies("LegiTrak")("VotecardID"))
			If VotecardID <> 0 Then
				Set cxnSQL = CreateObject("ADODB.Connection")
				With cxnSQL
					.Open strConnReadOnly
					strCommand = "SELECT * FROM [VoteCards] WHERE VotecardID=" & VotecardID
					Set rsResult = .Execute(strCommand)
					If IsNull(rsResult("VotecardID")) Then
						Response.Redirect "customer-votecards.asp"
					Else
						Desc = TweakQuote(rsResult("Description"))
						If rsResult("Chamber") = "S" Then
							Seneate = " selected"
						Else
							House = " selected"
						End If
					End If
					.Close
				End With
				Set cxnSQL = Nothing
			End If
		Else
			VotecardID = 0
			Response.Cookies("LegiTrak")("VotecardID") = ""
		End If
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg04 style='margin:10'>
<%
' DETERMINE VOTE CARD CREATION AUTHORIZATION

	strCommand = _
		"SELECT CreateVotecards FROM [Customer List] " & _
		"WHERE CreateVotecards=1 AND CustomerID=" & CustomerID
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnReadOnly
	Set rsResult = cxnSQL.Execute(strCommand)
	
	If Not rsResult.EOF Then
%>
<form id=Votecards action="votecard-create.asp" method=post>
<input type=hidden name=VotecardID value=0<%=Encrypt(VotecardID)%>>

<span class=hdg24 style='width:125'>Description: </span><input type=text name=Desc maxlength=24 size=32 value='<%=Desc%>'><br>
<span class=hdg24 style='width:125'>Chamber: </span><select name=Chamber>
<option value=S<%=Senate%>>Senate
<option value=H<%=House%>>House
</select><br>
<%
		If bolModifyVotecard Then
			Response.Write _
				"<span class=hdg24 style='width:121'>Delete vote card:</span>" & _
				"<input type=checkbox name=DeleteEntry value=True>"
		End If
%>
<center>
<span style='width:150'><input type=submit name=UpdateEntry value=Submit></span>
<span style='width:150'><input type=submit name=UpdateEntry value=Cancel></span>
</center>

<br><span class=hdg24>Invite Other LegiTrak Users:</span>

<table width=100% border=0 cellpadding=0 cellspacing=0 class=det00><col span=3 width=33%>
<tr valign=top><td>
<%
'
'
'	SPECIAL CASE FOR THE DEPT. OF NATURAL RESOURCES, ORGANIZATIONID = 19
'	They don't want to show up on this list
'
'
' CUSTOMER LIST
		If bolModifyVotecard Then
			strSQL = "(SELECT * FROM [Customer Votecards] WHERE VotecardID=" & VotecardID & ") AS V"
			SQLJoin = "([Customer List] C INNER JOIN [Organization List] O ON C.OrganizationID = O.OrganizationID)"
			SQLJoin = "(" & SQLJoin & " LEFT JOIN " & strSQL & " ON C.CustomerID = V.CustomerID) "
			strSQL = _
				"SELECT C.*, V.VotecardID " & _
				"FROM " & SQLJoin & _
				"WHERE" & _
				" O.OrganizationID <> 19 AND" & _
				" O.[Billing Clients] > 0 AND" & _
				" O.[Billing Type] <> 2 AND" & _
				" C.CustomerID <>" & CustomerID & _
				" ORDER BY C.[Contact First Name]+C.[Contact Last Name]"
		Else
			strSQL = _
				"SELECT * FROM [Customer List] C INNER JOIN [Organization List] O" & _
				" ON C.OrganizationID = O.OrganizationID " & _
				"WHERE" & _
				" O.OrganizationID <> 19 AND" & _
				" O.[Billing Clients] > 0 AND" & _
				" O.[Billing Type] <> 2 AND" & _
				" C.CustomerID <>" & CustomerID & _
				" ORDER BY C.[Contact First Name]+C.[Contact Last Name]"
		End If

		Set rsCustomers=Server.CreateObject("ADOR.Recordset")
		rsCustomers.CursorLocation = adUseClient ' so I can get the Recordcount
		rsCustomers.Open strSQL, strConnReadOnly
		Response.Write "<input type=hidden name=CustCount value=" & rsCustomers.RecordCount & ">"
		
		third = Int((rsCustomers.RecordCount+2)/3)
		b = 1
		e = third
		For j = 1 to 3
			For i = b to e
				Response.Write "<input type=checkbox"
				If VotecardID <> 0 Then
					If Not IsNull(rsCustomers("VotecardID")) Then
						Response.Write " checked"
					End If
				End If
				Response.Write " name=c" & i & " value=" & Encrypt(rsCustomers("CustomerID")) & "> &nbsp;"
				Response.Write rsCustomers("Contact First Name") & " " & rsCustomers("Contact Last Name") & "<br>"
				rsCustomers.MoveNext
			Next ' i
			Response.Write "</td><td>"
			b = j*third+1
			If j = 2 Then
				e = rsCustomers.RecordCount
			Else
				e = (j+1)*third
			End If
		Next ' j

		rsCustomers.Close
		Set rsCustomers = Nothing
%>
</td></tr>
</table>

<br>
<center>
<span style='width:150'><input type=submit name=UpdateEntry value=Submit></span>
<span style='width:150'><input type=submit name=UpdateEntry value=Cancel></span>
</center>
</form>

<%
	Else
' VOTE CARD CREATION AUTHORIZATION DENIED
		Response.Write "<div class=hdg24><br><br>"
		Response.Write "You are not currently authorized to create vote cards."
		Response.Write "<br><br>"
		Response.Write "For information on gaining access to this feature, please contact "
		Response.Write "<a href='mailto:Brad Tower <bhtower@towerltd.org>?Subject=BTS: Votecard Creation'>"
		Response.Write "Brad Tower</a>.</div>"
	End If

	cxnSQL.Close
	Set cxnSQL = Nothing
%>

</body>
</html>
