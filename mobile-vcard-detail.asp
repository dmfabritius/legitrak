<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "mobile-vcard-summary.asp?card=" & Request.Form("vcard")
	End If

	VotecardID = Decrypt(Request.Form("vcard"))
	LegislatorID = CInt(Request.Form("leg"))

' UPDATE LEGISLATOR'S VOTE
	If Request.Form("UpdateEntry") = "Submit" And VotecardID <> 0 Then
		Set cxnSQL = CreateObject("ADODB.Connection")
		With cxnSQL
			.Open strConnection
			strSQL = _
				"DELETE FROM [Votecard Details] " & _
				"WHERE VotecardID=" & VotecardID & " AND LegislatorID=" & LegislatorID
			.Execute strSQL, , adExecuteNoRecords
			strSQL = "INSERT INTO [Votecard Details] VALUES (" & _
				VotecardID & "," & _
				LegislatorID & "," & _
				CInt(Request.Form("vote")) & "," & _
				"'" & Now & "'," & _
				"NULL," & _
				CustomerID & ")"
			.Execute strSQL, , adExecuteNoRecords
			.Close
		End With
		Set cxnSQL = Nothing
		Response.Redirect "mobile-vcard-summary.asp?card=" & Request.Form("vcard")
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
<form action="mobile-vcard-detail.asp" method=post>
<input type=hidden name=vcard value=<%=Request.Form("vcard")%>>
<input type=hidden name=leg value=<%=LegislatorID%>>
<input type=submit name=UpdateEntry value=Submit>
<input type=submit name=UpdateEntry value=Cancel><table width=153 cellspacing=0 cellpadding=0 border=0>
<br>
<%
' LOAD VOTE CARD GENERAL INFORMATION
	strSQLJoin = "[Votecards]"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Customer List] ON [Votecards].Owner=[Customer List].CustomerID)"
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Customer Company Name]" & _
		" FROM " & strSQLJoin & _
		" WHERE [Votecards].VotecardID=" & VotecardID
	Set rsCustVcards=Server.CreateObject("ADOR.Recordset")
	rsCustVcards.Open strSQL, strConnReadOnly

' LOAD VOTE CARD DETAILS FOR LEGISLATOR
	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "[Legislators]"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS VDet ON [Legislators].LegislatorID = [VDet].LegislatorID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] ON [VDet].CustUpdate = [Customer List].CustomerID) "
	strSQL = 	"SELECT" & _
		" [VDet].*," & _
		" [Legislators].LastName, [Legislators].FirstName," & _
		" [Customer List].[Contact First Name]," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Customer Company Name] " & _
		"FROM " & strSQLJoin & _
		"WHERE [Legislators].LegislatorID=" & LegislatorID
	Set rsVDet=Server.CreateObject("ADOR.Recordset")
	rsVDet.Open strSQL, strConnReadOnly

	If rsCustVcards("Chamber") = "S" Then
		strChamber = "Senate"
	Else
		strChamber = "House"
	End If
	If IsNull(rsVDet("Updated")) Or IsNull(rsVDet("Vote")) Then
		strDate = "n/a"
	Else
		If Hour(rsVDet("Updated")) < 13 Then
			strTime = Hour(rsVDet("Updated")) & ":" & Right("0" & Minute(rsVDet("Updated")),2) & " am"
		Else
			strTime = Hour(rsVDet("Updated"))-12 & ":" & Right("0" & Minute(rsVDet("Updated")),2) & " pm"
		End If
		strDate = 	MonthName(Month(rsVDet("Updated")),True) & "-" & Day(rsVDet("Updated")) & " " & strTime
	End If
	If IsNull(rsVDet("LegUpdate")) And IsNull(rsVDet("Customer Company Name")) Then
		strWho = "n/a"
	Else
		If rsVDet("LegUpdate") Then
			strWho = "<b>" & rsVDet("FirstName") & " " & rsVDet("LastName") & "</b>"
		Else
			strWho = rsVDet("Contact First Name") & " " & rsVDet("Contact Last Name")
		End If
	End If

	Response.Write "<tr><td>Description:&nbsp;</td><td>" & rsCustVcards("Description") & "</td></tr>"
	Response.Write "<tr><td>Owner:&nbsp;</td><td>" & rsCustVcards("Customer Company Name") & "</td></tr>"
	Response.Write "<tr valign=top><td>Chamber:&nbsp;</td><td>" & strChamber & "<br><br></td></tr>"

	Response.Write "<tr><td colspan=2><b>"
	Response.Write rsVDet("FirstName") & " " & rsVDet("LastName")
	Response.Write "<b></td></tr>"
	Response.Write "<tr><td>Last&nbsp;Update:&nbsp;</td><td>" & strDate & "</td></tr>"
	Response.Write "<tr valign=top><td>Updated&nbsp;by:&nbsp;</td><td>" & strWho & "<br><br></td></tr>"

	Dim strVote(5)
	If IsNull(rsVDet("Vote")) Then
		strVote(0) = " selected"
	Else
		strVote(rsVDet("Vote")) = " selected"
	End If
	Response.Write "<tr><td colspan=2><select name=vote>"
	Response.Write "<option value=0" & strVote(0) & ">No Response"
	Response.Write "<option value=1" & strVote(1) & ">Yes"
	Response.Write "<option value=2" & strVote(2) & ">Leaning Yes"
	Response.Write "<option value=3" & strVote(3) & ">Undecided"
	Response.Write "<option value=4" & strVote(4) & ">Leaning No"
	Response.Write "<option value=5" & strVote(5) & ">No"
	Response.Write "</select></td></tr>"

	rsVDet.Close
	Set rsVDet = Nothing
	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>    
</table>
</form>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>