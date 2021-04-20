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
<td><a href='mobile-votecards.asp'>Vote Cards</a></td></tr>
<tr><td colspan=2><a href='mobile.asp'>Logout</a></td></tr>
</table><br>
<form action="mobile-vcard-detail.asp" method=post>
<table width=153 cellspacing=0 cellpadding=0 border=0>
<col width=83><col width=70>
<%
' DETERMINE VOTE CARD ID
	VotecardID = Decrypt(Request.QueryString("card"))

' LOAD VOTE CARD GENERAL INFORMATION
	strSQLJoin = "[Votecards]"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer List] ON [Votecards].Owner = [Customer List].CustomerID)"
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Customer Company Name]" & _
		" FROM " & strSQLJoin & _
		" WHERE [Votecards].VotecardID=" & VotecardID
	Set rsCustVcards=Server.CreateObject("ADOR.Recordset")
	rsCustVcards.Open strSQL, strConnReadOnly

	If rsCustVcards("Chamber") = "S" Then
		strChamber = "Senate"
		strSeat = "Seat=0"
	Else
		strChamber = "House"
		strSeat = "Seat<>0"
	End If

	Response.Write "<tr><td>Description:&nbsp;</td><td>" & rsCustVcards("Description") & "</td></tr>"
	Response.Write "<tr><td>Owner:&nbsp;</td><td>" & rsCustVcards("Customer Company Name") & "</td></tr>"
	Response.Write "<tr><td>Chamber:&nbsp;</td><td>" & strChamber & "</td></tr>"

' VOTE CARD STATISTICS SUMMARY
	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "[Legislators]"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS VDet ON [Legislators].LegislatorID = [VDet].LegislatorID) "
	strSQL = "SELECT" & _
		" [NR]=SUM(CASE WHEN Vote=0 OR VOTE IS NULL THEN 1 ELSE 0 END)," & _
		" [YS]=SUM(CASE WHEN Vote=1 THEN 1 ELSE 0 END)," & _
		" [LY]=SUM(CASE WHEN Vote=2 THEN 1 ELSE 0 END)," & _
		" [UN]=SUM(CASE WHEN Vote=3 THEN 1 ELSE 0 END)," & _
		" [LN]=SUM(CASE WHEN Vote=4 THEN 1 ELSE 0 END)," & _
		" [NO]=SUM(CASE WHEN Vote=5 THEN 1 ELSE 0 END)," & _
		" [LU]=MAX(Updated) " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" [Legislators].EndDate = '12/31/2299' AND" & _
		" [Legislators]." & strSeat
	Set rsDetail=Server.CreateObject("ADOR.Recordset")
	rsDetail.Open strSQL, strConnReadOnly

	If IsNull(rsDetail("LU")) Then
		strDate = ""
	Else
		If Hour(rsDetail("LU")) < 13 Then
			strTime = Hour(rsDetail("LU")) & ":" & Right("0" & Minute(rsDetail("LU")),2) & " am"
		Else
			strTime = Hour(rsDetail("LU"))-12 & ":" & Right("0" & Minute(rsDetail("LU")),2) & " pm"
		End If
		strDate = 	MonthName(Month(rsDetail("LU")),True) & "-" & Day(rsDetail("LU")) & " " & strTime
	End If

	Response.Write "<tr valign=top><td>Last&nbsp;Update:&nbsp;</td><td>" & strDate & "<br><br></td></tr>"
	Response.Write "<tr><td>No&nbsp;Response:&nbsp;</td><td>" & rsDetail("NR") & "</td></tr>"
	Response.Write "<tr><td>Yes:&nbsp;</td><td><b>" & rsDetail("YS") & "</b></td></tr>"
	Response.Write "<tr><td>Leaning&nbsp;Yes:&nbsp;</td><td><b>" & rsDetail("LY") & "</b></td></tr>"
	Response.Write "<tr><td>Undecided:&nbsp;</td><td><b>" & rsDetail("UN") & "</b></td></tr>"
	Response.Write "<tr><td>Leaning&nbsp;No:&nbsp;</td><td><b>" & rsDetail("LN") & "</b></td></tr>"
	Response.Write "<tr><td>No:&nbsp;</td><td><b>" & rsDetail("NO") & "</b></td></tr>"
	rsDetail.Close
	Set rsDetail = Nothing

	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>    
</table>
<br>
<input type=hidden name=vcard value=<%=Encrypt(VotecardID)%>>
<select name=leg>
<%
' LEGISLATOR LIST
	strSQL = _
		"SELECT * FROM [Legislators] " & _
		"WHERE EndDate='12/31/2299' AND " & strSeat & " ORDER BY LastName"
	Set rsLegs=Server.CreateObject("ADOR.Recordset")
	rsLegs.Open strSQL, strConnReadOnly

	Do Until rsLegs.EOF
		Response.Write "<option value=" & rsLegs("LegislatorID")
		Response.Write ">" & rsLegs("LastName") & ", " & rsLegs("FirstName")
		rsLegs.MoveNext
	Loop
	rsLegs.Close
	Set rsLegs = Nothing
%>
</select><br>
<input type=submit name=UpdateLeg value="Update Legislator's Vote">
</form>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>