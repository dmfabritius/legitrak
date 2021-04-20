<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "votecard-summary.asp"
	End If

' QUEUE UP SOLICITATION REQUEST
	If Request.Form("UpdateEntry") = "Submit" Then

		VotecardID = CInt(Decrypt(Request.Form("VotecardID")))
		If VotecardID = 0 Then Response.End

		Set cxnSQL = CreateObject("ADODB.Connection")
		With cxnSQL
			.Open strConnection

		' Determine VoteRequestID
			strCommand = "SELECT MAX(VoteRequestID) AS ID FROM [Vote Request Queue]"
			Set rsResult = .Execute(strCommand)
			If rsResult("ID") <> "" Then
				VoteRequestID = rsResult("ID")+1
			Else
				VoteRequestID = 1
			End If
			Set rsResult = Nothing
			
		' Create parent Request record
			strCommand = "INSERT INTO [Vote Request Queue] VALUES (" & _
				VoteRequestID & "," & _
				VotecardID & "," & _
				"'REQUESTED'," & _
				"'" & Now & "'," & _
				"NULL," & _
				"'" & TweakQuote(Request.Form("theSubject")) & "'," & _
				"'" & TweakQuote(Request.Form("theMessage")) & "')"
			.Execute strCommand, , adExecuteNoRecords

		' Create detail records
			For i = 1 to Request.Form("LegCount")
				If Request.Form("Solicit_" & i) = "True" Then
				' Create security key
					strKey=""
					Randomize
					For k = 1 to 24
						strKey = strKey + Chr(Int(Rnd(1)*26)+97)
					Next 'k
					strCommand = "INSERT INTO [Vote Request Details] VALUES (" & _
						VoteRequestID & "," & _
						CInt(Request.Form("LegID_" & i)) & "," & _
						"'" & strKey & "')"
					.Execute strCommand, , adExecuteNoRecords
				End If
			Next ' i

			.Close
		End With
		Set cxnSQL = Nothing
		Response.Redirect "votecard-summary.asp"
	Else
		VotecardID = Decrypt(Request.Cookies("LegiTrak")("VotecardID"))
	End If

' LOAD SELECTED VOTE CARD INFORMATION
	strSQLJoin = "[Votecards]"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer List] ON [Votecards].Owner=[Customer List].CustomerID)"
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Contact First Name]" & _
		" FROM " & strSQLJoin & _
		" WHERE [Votecards].VotecardID=" & VotecardID
	Set rsCustVcards=Server.CreateObject("ADOR.Recordset")
	rsCustVcards.Open strSQL, strConnReadOnly

	strDesc = rsCustVcards("Description")
	If rsCustVcards("Chamber") = "S" Then
		strChamber = "Senate"
		strSeat = "Seat=0"
	Else
		strChamber = "House"
		strSeat = "Seat<>0"
	End If

	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg04 scroll=yes>

<form id=VotecardLegislators action="votecard-legislators.asp" method=post>
<input type=hidden name=VotecardID value=<%=Encrypt(VotecardID)%>>

<table width=100% border=0 cellpadding=5 cellspacing=0 class=hdg29>
<col width=80 align=right>
<tr><td>Vote Card: &nbsp;</td><td class=shd29><%=strDesc%> &nbsp;(<%=strChamber%>)</td></tr>
<tr><td>Subject: &nbsp;</td>
<td><input name=theSubject type=text size=98></td></tr>
<tr valign=top><td>Message: &nbsp;</td>
<td><textarea name=theMessage rows=7 cols=100></textarea><br><br></td></tr>
</table>
<br>

<!-- VOTE CARD LEGISLATOR DETAILS TABLE -->
<table width=840 border=0 cellpadding=4 cellspacing=0>
<colgroup span=7 width=120>
<tr valign=top><td>
<%
' LOAD LEGISLATORS FOR THIS VOTE CARD
	strSQL = 	_
		"SELECT" & _
		" LegislatorID, [Rollcall Name] " & _
		"FROM [Legislators] " & _
		"WHERE" & _
		" EndDate = '12/31/2299' AND " & strSeat & _
		" ORDER BY [Rollcall Name]"
	Set rsLeg=Server.CreateObject("ADOR.Recordset")
	rsLeg.CursorLocation = adUseClient ' so I can get the Recordcount
	rsLeg.Open strSQL, strConnReadOnly
	Response.Write "<input type=hidden name=LegCount value=" & rsLeg.RecordCount & ">"
	
	intNumParts = 7
	intPartSize = Int((rsLeg.RecordCount+(intNumParts-1))/intNumParts)
	For j = 1 to intNumParts
		b = (j-1)*intPartSize+1
		If j < intNumParts Then
			e = j*intPartSize
		Else
			e = rsLeg.RecordCount
		End If

		Response.Write "<table width=120 border=0 cellpadding=0 cellspacing=0 class=det00>"
		Response.Write "<col width=25><col width=95>"

		For i = b to e
			Response.Write "<tr valign=middle><td>"
			Response.Write "<input type=hidden name=LegID_" & i
			Response.Write " value=" & rsLeg("LegislatorID") & ">"
			Response.Write "<input type=checkbox name=Solicit_" & i & " value='True'>"
			Response.Write "</td><td>"
			Response.Write rsLeg("Rollcall Name")
			Response.Write "</td></tr>"
			rsLeg.MoveNext
		Next 'i

		Response.Write "</table>"
		If j < intNumParts Then Response.Write "</td><td>"
	Next 'j

	rsLeg.Close
	Set rsLeg = Nothing
%>
</td></tr>
</table>
<br>

<center>
<input type=submit name=UpdateEntry value='Submit'>
<span style='margin-left:50'>
<input type=submit name=UpdateEntry value='Cancel'>
</span>
</center>
</form>


</body>
</html>
