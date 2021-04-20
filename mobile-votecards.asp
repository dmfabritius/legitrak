<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
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
<tr><td><a href='mobile-customer.asp'>Tracking List</a></td>
<td><a href='mobile-calendar.asp'>Calendar</a></td></tr>
<tr><td><a href='mobile-quickadd.asp'>Quick Add</a></td>
<td><i>Vote Cards</i></td></tr>
<tr><td colspan=2><a href='mobile.asp'>Logout</a></td></tr>
</table><br>
<%
	Response.Cookies("LegiTrak")("VotecardID") = ""
	Response.Cookies("LegiTrak")("LegislatorID") = ""
	strSQLJoin = "[Customer Votecards]"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Votecards] ON [Customer Votecards].VotecardID = Votecards.VotecardID)"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer List] ON [Votecards].Owner = [Customer List].CustomerID)"
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Customer Company Name]" & _
		" FROM " & strSQLJoin & _
		" WHERE [Customer Votecards].CustomerID=" & CustomerID & _
		" ORDER BY [Votecards].[Description]"
	set rsCustVcards=Server.CreateObject("ador.Recordset")
	rsCustVcards.Open strSQL, strConnReadOnly

	Response.Write "<table width=153 cellspacing=0 cellpadding=0 border=0>"
	If rsCustVcards.EOF Then
		Response.Write "<tr><td><br>"
		Response.Write "No vote cards have been created."
		Response.Write "</td></tr>"
	End If

	Do Until rsCustVcards.EOF
		If rsCustVcards("Chamber") = "S" Then
			strChamber = "Senate"
		Else
			strChamber = "House"
		End If
		Response.Write "<tr>"
		Response.Write "<td><a href='mobile-vcard-summary.asp?" & _
			"card=" & Encrypt(rsCustVcards("VotecardID")) & "'>"
		Response.Write rsCustVcards("Description") & "</a></td>"
		Response.Write "<td>" & strChamber & "</td>"
		Response.Write "</tr>"
		rsCustVcards.MoveNext
	Loop

	Response.Write "</table>"
	rsCustVcards.Close
	set rsCustVcards = Nothing
%>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>