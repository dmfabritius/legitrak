<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function gotoCard(v) {
	setCookie("VotecardID",v)
	setCookie("ModifyVotecard","True")
	window.location.href="votecard-create.asp"
}
function voteDetails(v) {
	setCookie("VotecardID",v)
	window.location.href="votecard-details.asp"
}
function voteLeg(v) {
	setCookie("VotecardID",v)
	window.location.href="votecard-legislators.asp"
}
function sortBy(f) {
	if (f=="Owner")
		f='[Customer List].[Contact Last Name], '
	else if (f=="Chamber")
		f='Votecards.Chamber, '
	setCookie("vOrderField",f+'Votecards.Description')
	window.location.href="votecard-summary.asp"
}
</script>
</head>
<body class=bkg03>

<table width=100% border=0 cellpadding=0 cellspacing=4 class=det00 style='padding:0 3'>
<col span=2><col span=7 align=center>
<tr class=hdg29>
<td align=left style='cursor:pointer' onclick='sortBy("")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Description</td>
<td align=left style='cursor:pointer' onclick='sortBy("Owner")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Owner</td>
<td style='cursor:pointer' onclick='sortBy("Chamber")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Chamber</td>
<td>No<br>Response</td>
<td>Yes</td>
<td>Leaning<br>Yes</td>
<td>Not<br>Decided</td>
<td>Leaning<br>No</td>
<td>No</td>
<td>Last<br>Update</td>
</tr>
<%
	Set rsCustVcards=Server.CreateObject("ADOR.Recordset")
	Set rsDetail=Server.CreateObject("ADOR.Recordset")
	
' TEST FOR VOTE CARD CREATION PRIVILEGES
	strSQL = "SELECT CreateVotecards FROM [Customer List] WHERE CreateVotecards=1 AND CustomerID=" & CustomerID
	rsCustVcards.Open strSQL, strConnReadOnly
	If Not rsCustVcards.EOF Then
		Response.Write _
			"<tr class=shd29><td class=lnk70 onclick='window.location.href=""votecard-create.asp""'>" & _
			"Create vote card</td><td colspan=9></td></tr>"
	End If
	rsCustVcards.Close

' CUSTOMER VOTE CARDS
	OrderField = Request.Cookies("LegiTrak")("vOrderField")
	If Len(OrderField) = 0 Then OrderField = "[Votecards].[Description]"
	strSQLJoin = "[Customer Votecards]"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Votecards] ON [Customer Votecards].VotecardID=Votecards.VotecardID)"
	strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer List] ON [Votecards].Owner=[Customer List].CustomerID)"
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Contact First Name]" & _
		" FROM " & strSQLJoin & _
		" WHERE [Customer Votecards].CustomerID=" & CustomerID & _
		" ORDER BY " & OrderField
	rsCustVcards.Open strSQL, strConnReadOnly

	If rsCustVcards.EOF Then
		Response.Write _
			"<tr><td colspan=10 align=center class=hdg24 style='padding:20'>" & _
			"You are not currently participating in any vote cards." & _
			"</td></tr>"
	End If

	Do Until rsCustVcards.EOF
		vcard = Encrypt(rsCustVcards("VotecardID"))
		Response.Write _
			"<tr class=bkg04 valign=top>" & _
			"<td class=lnk70 onclick='voteDetails(""" & vcard & """)'>" & _
			rsCustVcards("Description") & "</td>"
		If rsCustVcards("Owner") = CustomerID Then
			Response.Write "<td class=lnk70 onclick='gotoCard(""" & vcard & """)'>"
		Else
			Response.Write "<td>"
		End If
		Response.Write _
			rsCustVcards("Contact First Name") & " " & _
			rsCustVcards("Contact Last Name") & "</td>"
		If rsCustVcards("Chamber") = "S" Then
			Response.Write "<td>" & "Senate" & "</td>"
			strSeat = "Seat=0"
		Else
			Response.Write "<td>" & "House" & "</td>"
			strSeat = "Seat<>0"
		End If

		strSQL = _
			"SELECT * FROM [Votecard Details] " & _
			"WHERE [Votecard Details].VotecardID=" & rsCustVcards("VotecardID")
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
		rsDetail.Open strSQL, strConnReadOnly
		If IsNull(rsDetail("LU")) Then
			strDate = ""
		Else
			If Hour(rsDetail("LU")) < 13 Then
				strTime = Hour(rsDetail("LU")) & ":" & Right("0" & Minute(rsDetail("LU")),2) & " am"
			Else
				strTime = Hour(rsDetail("LU"))-12 & ":" & Right("0" & Minute(rsDetail("LU")),2) & " pm"
			End If
			strDate = 	MonthName(Month(rsDetail("LU")),True) & " " & Day(rsDetail("LU")) & ", " & strTime
		End If
		If CustomerID = 1 Or CustomerID = 24 Or CustomerID = 41 Then
			Response.Write "<td class=lnk70 onclick='voteLeg(""" & vcard & """)'>"
		Else
			Response.Write "<td>"
		End If
		Response.Write rsDetail("NR") & "</td>"
		Response.Write _
			"<td>" & rsDetail("YS") & "</td>" & _
			"<td>" & rsDetail("LY") & "</td>" & _
			"<td>" & rsDetail("UN") & "</td>" & _
			"<td>" & rsDetail("LN") & "</td>" & _
			"<td>" & rsDetail("NO") & "</td>" & _
			"<td>" & strDate & "</td></tr>"
		rsDetail.Close
		rsCustVcards.MoveNext
	Loop

	Set rsDetail = Nothing
	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>
</table>
<div class=bkg04 style='position:relative;height:100%;margin:0 4'></div>

</body>
</html>
