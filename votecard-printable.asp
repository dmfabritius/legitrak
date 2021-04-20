<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' LOAD SELECTED VOTE CARD
	VotecardID = Decrypt(Request.Cookies("LegiTrak")("VotecardID"))
	strSQL = "SELECT" & _
		" Votecards.*," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Contact First Name]" & _
		" FROM [Votecards] INNER JOIN [Customer List] ON [Votecards].Owner = [Customer List].CustomerID" & _
		" WHERE [Votecards].VotecardID=" & VotecardID
	Set rsCustVcards=Server.CreateObject("ADOR.Recordset")
	rsCustVcards.Open strSQL, strConnReadOnly

	strDesc = rsCustVcards("Description")

	Dim strVote(5)
	strVote(0)="&nbsp;"
	strVote(1)="Yes"
	strVote(2)="Leaning Yes"
	strVote(3)="Undecided"
	strVote(4)="Leaning No"
	strVote(5)="No"
%>
<html>
<head>
<title><% =CustomerName %> - Votecard Report</title>
<style><!--
@page Section1{size:8.5in 11in;margin:.75in .5in .75in .5in}
div.Section1{page:Section1}
.nameTitle{font:8pt Tahoma}
.nameSubTitle{font-family:Book Antiqua;font-variant:small-caps;font-size:14pt}
.reportTitle{font-family:Book Antiqua;font-size:14pt;font-weight:bold}
.reportSubTitle{font-family:Book Antiqua;font-size:12pt}
.subReportTitle{font-family:Book Antiqua;font-variant:small-caps;font-size:12pt;
    text-decoration:underline;font-weight:bold;color:#000080;letter-spacing:.2em}
.subReportHeader{font-family:Book Antiqua;font-variant:small-caps;font-size:11pt;
    font-weight:bold;letter-spacing:.2em}
.subReportHeader4{font-family:Book Antiqua;font-variant:small-caps;font-size:11pt;
    font-weight:bold;letter-spacing:.2em;color:#000080}
.subReportHeader2{font-family:Book Antiqua;font-variant:small-caps;font-size:10pt;
    font-weight:bold;text-decoration:underline}
.subReportHeader3{font-family:Book Antiqua;font-variant:small-caps;font-size:10pt;
    font-weight:bold;text-decoration:underline;color:#000080}
.subReportSmallHeader{font-family:Book Antiqua;font-size:10pt;font-weight:bold}
.subReportSmallHeader2{font-family:Book Antiqua;font-size:10pt;font-weight:bold;color:#000080}
.subReportSmallHeader3{font-family:Book Antiqua;font-size:10pt;color:#000080;
	vertical-align:bottom;border-bottom:2px solid #808080}
.subReportDetail{font:5pt Tahoma;vertical-align:top}
--></style>
</head>
<body><div class=Section1>
<!--
'------------------------------------------------
' VOTECARD - REPORT HEADER
'------------------------------------------------
-->
<table width=360 border=0 cellspacing=0 cellpadding=0 class=nameTitle>
<tr><td>Prepared by LegiTrak for <% =CustomerName %>
 on <% Response.Write FormatDateTime(Date, vbShortDate) %>
 for <% =strDesc %></td></tr>
<tr><td><br>Yays ___; Nays ___; Abstain ___; Absent ___</td></tr>
</table><br>

<!--
'------------------------------------------------
' VOTECARD - DETAILS - HOUSE
'------------------------------------------------
-->
<table width=360 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=180><col width=180>
<tr valign=top><td>
<%
' LOAD VOTE CARD DETAILS
	strSQL = "SELECT * FROM [Votecard Details] WHERE VotecardID=" & VotecardID
	strSQLJoin = "[Legislators] L"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS V ON L.LegislatorID = V.LegislatorID)"
	strSQL = 	"SELECT" & _
		" V.*," & _
		" L.[Rollcall Name], L.LegislatorID, L.Party, L.DistrictID, " & _
		" L.[BusinessStreet] St, RIGHT(L.[BusinessPhone],4) Ph " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" L.EndDate = '12/31/2299' AND L.Seat<>0" & _
		" ORDER BY L.[Rollcall Name]"

	Set rsLeg=Server.CreateObject("ADOR.Recordset")
	rsLeg.CursorLocation = adUseClient ' so I can get the Recordcount
	rsLeg.Open strSQL, strConnReadOnly
	
	half = Int((rsLeg.RecordCount+1)/2)
	b = 1
	e = half
	For j = 1 to 2
		Response.Write "<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>"
		Response.Write "<col width=120><col width=60>"

		For i = b to e
			If strBackground = "#F4F4F4" Then
				strBackground = "#FFFFFF"
			Else
				strBackground = "#F4F4F4"
			End If
		
			Response.Write "<tr valign=center bgcolor=" & strBackground & "><td>"
			Response.Write "___ " & rsLeg("Ph")
			Response.Write " &nbsp;"
			Response.Write rsLeg("Rollcall Name")
			Response.Write " (" & rsLeg("Party") & "-" & rsLeg("DistrictID") & ")"
			Response.Write "</td><td>"
			Response.Write rsLeg("St")
			Response.Write "</td></tr>"
			rsLeg.MoveNext
		Next 'i
		Response.Write "</table>"
		If j=1 Then Response.Write "</td><td>"
		b = half+1
		e = rsLeg.RecordCount
	Next 'j
	rsLeg.Close
	Set rsLeg = Nothing
%>
</td></tr>
</table>

<br><br>

<!--
'------------------------------------------------
' VOTECARD - DETAILS - SENATE
'------------------------------------------------
-->
<table width=360 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=180><col width=180>
<tr valign=top><td>
<%
' LOAD VOTE CARD DETAILS
	strSQL = "SELECT * FROM [Votecard Details] WHERE VotecardID=" & VotecardID
	strSQLJoin = "[Legislators] L"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS V ON L.LegislatorID = V.LegislatorID)"
	strSQL = 	"SELECT" & _
		" V.*," & _
		" L.[Rollcall Name], L.LegislatorID, L.Party, L.DistrictID, " & _
		" L.[BusinessStreet] St, RIGHT(L.[BusinessPhone],4) Ph " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" L.EndDate = '12/31/2299' AND L.Seat=0" & _
		" ORDER BY L.[Rollcall Name]"

	Set rsLeg=Server.CreateObject("ADOR.Recordset")
	rsLeg.CursorLocation = adUseClient ' so I can get the Recordcount
	rsLeg.Open strSQL, strConnReadOnly
	
	half = Int((rsLeg.RecordCount+1)/2)
	b = 1
	e = half
	For j = 1 to 2
		Response.Write "<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>"
		Response.Write "<col width=120><col width=60>"

		For i = b to e
			If strBackground = "#F4F4F4" Then
				strBackground = "#FFFFFF"
			Else
				strBackground = "#F4F4F4"
			End If
		
			Response.Write "<tr valign=center bgcolor=" & strBackground & "><td>"
			Response.Write "___ " & rsLeg("Ph")
			Response.Write " &nbsp;"
			Response.Write rsLeg("Rollcall Name")
			Response.Write " (" & rsLeg("Party") & "-" & rsLeg("DistrictID") & ")"
			Response.Write "</td><td>"
			Response.Write rsLeg("St")
			Response.Write "</td></tr>"
			rsLeg.MoveNext
		Next 'i
		Response.Write "</table>"
		If j=1 Then Response.Write "</td><td>"
		b = half+1
		e = rsLeg.RecordCount
	Next 'j
	rsLeg.Close
	Set rsLeg = Nothing
%>
</td></tr>
</table>

<br><br>

<!--
'------------------------------------------------
' COMMITTEES
'------------------------------------------------
-->
<table width=360 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=180><col width=180>

<tr valign=top><td>
<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<tr><td><b>SENATE</b></td></tr>
<%
	bolChamber = False

' LOAD COMMITTEES
	strSQL = "SELECT * FROM [Committees] WHERE ISNULL([House],'')<>'' ORDER BY [House] DESC, [Committee Name]"

	Set rsComm=Server.CreateObject("ADOR.Recordset")
	rsComm.CursorLocation = adUseClient ' so I can get the Recordcount
	rsComm.Open strSQL, strConnReadOnly

	Do
		If rsComm("House") = "House" AND bolChamber = False Then
			Response.Write "</table></td><td>"
			Response.Write "<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail style='text-align:right'>"
			Response.Write "<tr><td><b>HOUSE</b></td></tr>"
			bolChamber = True
		End If

		Response.Write "<tr><td>"
		If rsComm("House")= "Senate" Then
			Response.Write Right("xxxx" & rsComm("Telephone"),4) & " " & rsComm("Committee Name")
		Else
			Response.Write rsComm("Committee Name") & " " & Right("xxxx" & rsComm("Telephone"),4)
		End If
		Response.Write "</td></tr>"

		rsComm.MoveNext
	Loop Until rsComm.EOF
	rsComm.Close
	Set rsComm = Nothing
%>
</table>
</td></tr>
</table>

<!--
'------------------------------------------------
' Other Info
'------------------------------------------------
-->
<br>
<table width=360 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=180><col width=180>
<tr valign=top><td>
<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<tr><td>7550 Secretary/Clerk</td></tr>
<tr><td>7541 Sgt. At Arms</td></tr>
<tr><td>7350 Democratic Caucus</td></tr>
<tr><td>7517 Republican Caucus</td></tr>
<tr><td>7593 Workroom</td></tr>
<tr><td>6777 Code Reviser</td></tr>
<tr><td>7573 Leg. Info Center</td></tr>
<tr><td>7540 Leg. Ethics Board</td></tr>
<tr><td>753-5000 State Info</td></tr>
<tr><td>754-3290 Ulcer Gulch</td></tr>
</table></td><td>
<table width=180 cellpadding=0 cellspacing=0 border=0 class=subReportDetail style='text-align:right'>
<tr><td>Secretary/Clerk 7750</td></tr>
<tr><td>Sgt. At Arms 7760/7771</td></tr>
<tr><td>Democratic Caucus 7222</td></tr>
<tr><td>Republican Caucus 7791</td></tr>
<tr><td>Workroom 7780</td></tr>
<tr><td>Tours 902-8880</td></tr>
<tr><td>Gov. Office 902-4111</td></tr>
<tr><td>Exec. Ethics 664-0871</td></tr>
<tr><td>PDC 753-1111</td></tr>
<tr><td>Leg Hot Line 800-562-6000</td></tr>
</table>
</td></tr>
</table>
</div>
</body>
</html>