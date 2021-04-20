<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "votecard-summary.asp"
	End If

	If Request.Form("UpdateEntry") = "Submit" Then
		VotecardID = CInt(Decrypt(Request.Form("VotecardID")))
		If VotecardID = 0 Then Response.End
		Set cxnSQL = CreateObject("ADODB.Connection")
		With cxnSQL
			.Open strConnection

	' Update votecard details
			For i = 1 to Request.Form("LegCount")
			    LegID = CInt(Request.Form("L" & i))
				If Request.Form("U" & i) = "True" Then
					strCommand = _
						"DELETE FROM [Votecard Details] " & _
						"WHERE VotecardID=" & VotecardID & " AND LegislatorID=" & LegID
					.Execute strCommand, , adExecuteNoRecords
					strCommand = "INSERT INTO [Votecard Details] VALUES (" & _
						VotecardID & "," & _
						LegID & "," & _
						CInt(Request.Form("V" & i)) & "," & _
						"'" & Now & "'," & _
						"NULL," & _
						CustomerID & ")"
					.Execute strCommand, , adExecuteNoRecords
				End If
			Next ' i

			.Close
		End With
		Set cxnSQL = Nothing
	Else
		VotecardID = Decrypt(Request.Cookies("LegiTrak")("VotecardID"))
	End If

' LOAD SELECTED VOTE CARD
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

	If rsCustVcards("Chamber") = "S" Then
		strChamber = "Senate"
		strSeat = "Seat=0"
	Else
		strChamber = "House"
		strSeat = "Seat<>0"
	End If

	Dim TypeSel(4), CommSel(1000)
	intFilterType = CInt("0" & Request.Cookies("LegiTrak")("FilterType"))
	TypeSel(intFilterType) = " selected"
	intFilterCmtte = CInt("0" & Request.Cookies("LegiTrak")("FilterCmtte"))
	CommSel(intFilterCmtte) = " selected"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>select{font:8.5pt Arial}</style>
<script src="js/bts.js"></script>
<script>
function mark2(i) {
	document.getElementsByName("V"+i)[0].style.backgroundColor=myStyles[".hdg29"].backgroundColor
	document.getElementsByName("U"+i)[0].value="True"
}
function updateFilter(){
	vcd=document.getElementById("VotecardDetails")
	setCookie("FilterType",vcd.FilterType.value)
	setCookie("FilterCmtte",vcd.FilterCmtte.value)
	window.location.href=window.location.href
}
function printable(){
	window.open("votecard-printable.asp","print")
}
function filler(){
	if (f=document.getElementById("Filler")){
		f.style.height=0
		f.style.height=document.getElementById("Leg1").offsetHeight-document.getElementById("Leg2").offsetHeight+14
	}
}
</script>
</head>

<body onload='filler()' onresize='filler()' class=bkg03>

<form id=VotecardDetails method=post action="votecard-details.asp">
<input type=hidden name=VotecardID value=<%=Encrypt(VotecardID)%>>

<!-- VOTE CARD SUMMARY HEADER TABLE -->
<table width=100% border=0 cellpadding=0 cellspacing=4 class=det00 style='padding:3'>
<col span=2><col span=7 align=center>
<tr class=hdg29>
<td>Description</td><td>Owner</td><td>Chamber</td><td>No<br>Response</td>
<td>Yes</td><td>Leaning<br>Yes</td>
<td>Not<br>Decided</td>
<td>Leaning<br>No</td><td>No</td>
<td>Last<br>Update</td>
</tr>
<%
' VOTE CARD SUMMARY
	Response.Write _
		"<tr class=bkg04 valign=top>" & _
		"<td>" & rsCustVcards("Description") & "</td>" & _
		"<td>" & rsCustVcards("Contact First Name") & " " & _
		rsCustVcards("Contact Last Name") & "</td>" & _
		"<td>" & strChamber & "</td>"

	Select Case intFilterType
		Case 0 : strSQLWhere = ""
		Case 1 : strSQLWhere = " AND [Legislators].Party='D'"
		Case 2 : strSQLWhere = " AND [Legislators].Party='R'"
	End Select
	
	strSQLJoin = "[Legislators]"
	If intFilterCmtte <> 0 Then
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Committee Membership] ON Legislators.LegislatorID = [Committee Membership].LegislatorID)"
		strSQLWhere = strSQLWhere & _
			" AND [Committee Membership].Year=" & Year(Date) & _
			" AND [Committee Membership].CommitteeID=" & intFilterCmtte
	End If
	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS VDet ON [Legislators].LegislatorID = [VDet].LegislatorID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] ON [VDet].CustUpdate = [Customer List].CustomerID) "
	strSQL = 	"SELECT" & _
		" [NR]=SUM(CASE WHEN Vote=0 OR VOTE IS NULL THEN 1 ELSE 0 END)," & _
		" [YS]=SUM(CASE WHEN Vote=1 THEN 1 ELSE 0 END)," & _
		" [LY]=SUM(CASE WHEN Vote=2 THEN 1 ELSE 0 END)," & _
		" [UN]=SUM(CASE WHEN Vote=3 THEN 1 ELSE 0 END)," & _
		" [LN]=SUM(CASE WHEN Vote=4 THEN 1 ELSE 0 END)," & _
		" [NO]=SUM(CASE WHEN Vote=5 THEN 1 ELSE 0 END)," & _
		" [LU]=MAX(Updated) " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" [Legislators].EndDate = '12/31/2299' AND " & strSeat & strSQLWhere

'response.Write strsql

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
		strDate = 	MonthName(Month(rsDetail("LU")),True) & " " & Day(rsDetail("LU")) & ", " & strTime
	End If

	Response.Write _
		"<td>" & rsDetail("NR") & "</td>" & _
		"<td>" & rsDetail("YS") & "</td>" & _
		"<td>" & rsDetail("LY") & "</td>" & _
		"<td>" & rsDetail("UN") & "</td>" & _
		"<td>" & rsDetail("LN") & "</td>" & _
		"<td>" & rsDetail("NO") & "</td>" & _
		"<td>" & strDate & "</td></tr>"
	rsDetail.Close
	Set rsDetail = Nothing

	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>
</table>

<table width=100% border=0 cellpadding=0 cellspacing=0>

<tr class=bkg04 height=35 valign=middle><td colspan=3 align=center style='border:4px solid white;border-width:0 4px'>
<span style='width:150'><input type=submit name=UpdateEntry value=Submit></span>
<span style='width:150'><input type=submit name=UpdateEntry value=Cancel></span>
<span style='width:150'><input type=button value=Print onclick='printable()' style='width:55'></span>
</td></tr>

<tr class=hdg29 height=35 valign=middle><td colspan=3 align=center style='border:4px solid white;border-bottom-width:0'>
Filter by: &nbsp;
<select name=FilterType onchange='updateFilter()'>
  <option value=0<%=TypeSel(0)%>>-Both Parties-
  <option value=1<%=TypeSel(1)%>>Democrats
  <option value=2<%=TypeSel(2)%>>Republicans
</select> &nbsp;<select name=FilterCmtte onchange='updateFilter()'>
<option value=0>-All Committees- <%
' COMMITTEE LIST
	strSQL = _
		"SELECT * FROM [Committees] " & _
		"WHERE House='" & strChamber & "' " & _
		"ORDER BY [Committee Name]"
	Set rsComms=Server.CreateObject("ADOR.Recordset")
	rsComms.Open strSQL, strConnReadOnly
	i = 1
	Do Until rsComms.EOF
		Response.Write _
			"<option value=" & rsComms("CommitteeID") & CommSel(rsComms("CommitteeID")) & ">" & _
			rsComms("Committee Name")
		rsComms.MoveNext
		i=i+1
	Loop
	rsComms.Close
	Set rsComms = Nothing
%>
</select></td></tr>

<tr valign=top><td>
<%
' LOAD VOTE CARD DETAILS
	strSQLJoin = "[Legislators]"

	If intFilterCmtte <> 0 Then _
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Committee Membership] ON Legislators.LegislatorID = [Committee Membership].LegislatorID)"

	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS VDet ON [Legislators].LegislatorID = [VDet].LegislatorID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] ON [VDet].CustUpdate = [Customer List].CustomerID) "
	strSQL = 	"SELECT" & _
		" [VDet].*," & _
		" [Legislators].[Rollcall Name], [Legislators].LegislatorID," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Contact First Name] " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" [Legislators].EndDate = '12/31/2299' AND " & strSeat & strSQLWhere & _
		" ORDER BY [Rollcall Name]"
	Set rsLeg=Server.CreateObject("ADOR.Recordset")
	rsLeg.CursorLocation = adUseClient ' so I can get the Recordcount
	rsLeg.Open strSQL, strConnReadOnly
	Response.Write "<input type=hidden name=LegCount value=" & rsLeg.RecordCount & ">"
	
	Dim strVote(5)
	half = Int((rsLeg.RecordCount+1)/2)
	b = 1
	e = half
	For j = 1 to 2
		Response.Write "<table id=Leg" & j & " width=100% border=0 cellpadding=0 cellspacing=4 class=det00 style='padding:0 3'>"
		Response.Write "<tr class=hdg29><td colspan=2></td>"
		Response.Write "<td align=left>Last Update</td></tr>" ' colspan=2

		For i = b to e

			For k = 0 to 5
				strVote(k) = ""
			Next 'k
			If IsNull(rsLeg("Vote")) Then
				strVote(0) = " selected"
			Else
				strVote(rsLeg("Vote")) = " selected"
			End If

			If IsNull(rsLeg("Updated")) Or IsNull(rsLeg("Vote")) Then
				strDate = ""
			Else
				If Hour(rsLeg("Updated")) < 13 Then
					strTime = Hour(rsLeg("Updated")) & ":" & Right("0" & Minute(rsLeg("Updated")),2) & " am"
				Else
					strTime = Hour(rsLeg("Updated"))-12 & ":" & Right("0" & Minute(rsLeg("Updated")),2) & " pm"
				End If
				strDate = 	MonthName(Month(rsLeg("Updated")),True) & " " & Day(rsLeg("Updated")) & ", " & strTime
			End If

			If IsNull(rsLeg("LegUpdate")) And IsNull(rsLeg("Contact Last Name")) Then
				strWho = ""
			Else
				If rsLeg("LegUpdate") Then
					strWho = ", <b>" & rsLeg("Rollcall Name") & "</b>"
				Else
					strWho = ", " & _
						rsLeg("Contact Last Name") & ", " & _
						Left(rsLeg("Contact First Name"),1) & "."
				End If
			End If
		
			Response.Write "<tr class=bkg04 valign=middle><td>"
			Response.Write rsLeg("Rollcall Name")
			Response.Write "</td><td>"
			Response.Write "<input type=hidden name=U" & i & " value=False>"
			Response.Write "<input type=hidden name=L" & i
			Response.Write " value=" & rsLeg("LegislatorID") & ">"
			Response.Write "<select class=bt1 onchange='mark2(" & i & ")'"
			Response.Write " name=V" & i & ">"
			Response.Write "<option value=0" & strVote(0) & ">No Response"
			Response.Write "<option value=1" & strVote(1) & ">Yes"
			Response.Write "<option value=2" & strVote(2) & ">Leaning Yes"
			Response.Write "<option value=3" & strVote(3) & ">Undecided"
			Response.Write "<option value=4" & strVote(4) & ">Leaning No"
			Response.Write "<option value=5" & strVote(5) & ">No"
			Response.Write "</select></td><td>"
			Response.Write strDate & strWho & "</td></tr>"
'			Response.Write strDate & "</td><td>"
'			Response.Write strWho & "</td></tr>"
			rsLeg.MoveNext
		Next 'i

		If e = rsLeg.RecordCount And e Mod 2 = 1 Then Response.Write _
			"<tr id=Filler class=bkg04>" & _
			"<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"


		Response.Write "</table>"
		If j=1 Then Response.Write "</td><td>"
		b = half+1
		e = rsLeg.RecordCount
	Next 'j
	rsLeg.Close
	Set rsLeg = Nothing
%>
</td></tr>

<tr class=bkg04 height=35 valign=middle><td colspan=3 align=center style='border:4px solid white;border-width:0 4px'>
<span style='width:150'><input type=submit name=UpdateEntry value=Submit></span>
<span style='width:150'><input type=submit name=UpdateEntry value=Cancel></span>
<span style='width:150'><input type=button value=Print onclick='printable()' style='width:55' id="Button1" name="Button1"></span>
</td></tr>

</table>
<div class=bkg04 style='position:relative;height:100%;margin:0 4'></div>

</form>

</body>
</html>
