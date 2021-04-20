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

	If rsCustVcards("Chamber") = "S" Then
		strChamber = "Senate"
		strSeat = "Seat=0"
	Else
		strChamber = "House"
		strSeat = "Seat<>0"
	End If
%>
<html>
<head>
<title><% =CustomerName %> - Votecard Report</title>
<style><!--
@page Section1{size:8.5in 11in;margin:.75in .5in .75in .5in}
div.Section1{page:Section1}
.nameTitle{font-family:Book Antiqua;font-variant:small-caps;font-size:18pt;font-weight:bold;
	color:#000080;letter-spacing:.1em}
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
.subReportDetail{font-family:Book Antiqua;font-size:9pt;vertical-align:top}
--></style>
</head>
<body><div class=Section1>
<!--
'------------------------------------------------
' VOTECARD - REPORT HEADER
'------------------------------------------------
-->
<table width=720 border=0 cellspacing=0 cellpadding=0>
<tr><td width=450 class=nameTitle><% = CustomerName %></td>
<td width=270 align=right class=reportTitle>Votecard</td></tr>
<tr><td width=720 align=right colspan=2 class=reportSubTitle>
<% Response.Write FormatDateTime(Date, vbLongDate) %></td></tr>
</table><br>

<table width=720 border=0 cellspacing=0 cellpadding=0><tr><td>
<table width=265 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=75 align=right><col width=190 class=bt0><tr>
<td>Description: &nbsp;</td>
<td class=subReportSmallHeader><%=rsCustVcards("Description")%></td></tr><tr>
<td>Owner: &nbsp;</td>
<td class=subReportSmallHeader><%=rsCustVcards("Contact First Name") & " " & rsCustVcards("Contact Last Name") %></td></tr><tr>
<td>Chamber: &nbsp;</td>
<td class=subReportSmallHeader><%=strChamber%></td></tr></table></td>
<td valign=center>
<!--
'------------------------------------------------
' VOTECARD - SUMMARY STATISTICS
'------------------------------------------------
-->
<table width=455 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=60 align=center><col width=60 align=center>
<col width=60 align=center><col width=60 align=center>
<col width=60 align=center><col width=60 align=center>
<col width=95>
<tr>
<th class=subReportSmallHeader3>No<br>Response</th>
<th class=subReportSmallHeader3>Yes</th>
<th class=subReportSmallHeader3>Leaning<br>Yes</th>
<th class=subReportSmallHeader3>Not<br>Decided</th>
<th class=subReportSmallHeader3>Leaning<br>No</th>
<th class=subReportSmallHeader3>No</th>
<th class=subReportSmallHeader3>Last<br>Update</th>
</tr>
<%
' VOTE CARD SUMMARY
	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "[Legislators]"
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
		" Legislators.EndDate = '12/31/2299' AND " & strSeat
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

	Response.Write "<td>" & rsDetail("NR") & "</td>"
	Response.Write "<td>" & rsDetail("YS") & "</td>"
	Response.Write "<td>" & rsDetail("LY") & "</td>"
	Response.Write "<td>" & rsDetail("UN") & "</td>"
	Response.Write "<td>" & rsDetail("LN") & "</td>"
	Response.Write "<td>" & rsDetail("NO") & "</td>"
	Response.Write "<td>" & strDate & "</td>"
	response.write "</tr>"
	rsDetail.Close
	Set rsDetail = Nothing

	rsCustVcards.Close
	Set rsCustVcards = Nothing
%>
</table>
</td></tr></table>
</table><br>


<!--
'------------------------------------------------
' VOTECARD - DETAILS
'------------------------------------------------
-->
<table width=720 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>
<col width=360><col width=360>
<tr valign=top><td>
<%
' LOAD VOTE CARD DETAILS
	strSQL = "SELECT * FROM [Votecard Details] WHERE [Votecard Details].VotecardID=" & VotecardID
	strSQLJoin = "[Legislators]"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN (" & strSQL & ") AS VDet ON [Legislators].LegislatorID = [VDet].LegislatorID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] ON [VDet].CustUpdate = [Customer List].CustomerID) "
	strSQL = 	"SELECT" & _
		" [VDet].*," & _
		" [Legislators].[Rollcall Name], [Legislators].LegislatorID," & _
		" [Customer List].[Contact Last Name]," & _
		" [Customer List].[Contact First Name] " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		" [Legislators].EndDate = '12/31/2299' AND " & strSeat & _
		" ORDER BY [Rollcall Name]"
	Set rsLeg=Server.CreateObject("ADOR.Recordset")
	rsLeg.CursorLocation = adUseClient ' so I can get the Recordcount
	rsLeg.Open strSQL, strConnReadOnly
	
	Dim strVote(5)
	strVote(0)="&nbsp;"
	strVote(1)="Yes"
	strVote(2)="Leaning Yes"
	strVote(3)="Undecided"
	strVote(4)="Leaning No"
	strVote(5)="No"

	half = Int((rsLeg.RecordCount+1)/2)
	b = 1
	e = half
	For j = 1 to 2
		Response.Write "<table width=360 cellpadding=0 cellspacing=0 border=0 class=subReportDetail>"
		Response.Write "<col width=85><col width=80><col width=100><col width=95>"
		Response.Write "<tr><th class=subReportSmallHeader3></th><th class=subReportSmallHeader3></th>"
		Response.Write "<th class=subReportSmallHeader3 align=left>Updated On</th>"
		Response.Write "<th class=subReportSmallHeader3 align=left>Updated By</th></tr>"

		For i = b to e
           If strBackground = "#F4F4F4" Then
               strBackground = "#FFFFFF"
           Else
               strBackground = "#F4F4F4"
           End If

			If IsNull(rsLeg("Updated")) Or IsNull(rsLeg("Vote")) Then
				strDate = ""
			Else
				If Hour(rsLeg("Updated")) < 13 Then
					strTime = Hour(rsLeg("Updated")) & ":" & Right("0" & Minute(rsLeg("Updated")),2) & " am"
				Else
					strTime = Hour(rsLeg("Updated"))-12 & ":" & Right("0" & Minute(rsLeg("Updated")),2) & " pm"
				End If
				strDate = 	MonthName(Month(rsLeg("Updated")),True) & "-" & Day(rsLeg("Updated")) & " " & strTime
			End If

			If IsNull(rsLeg("LegUpdate")) And IsNull(rsLeg("Contact Last Name")) Then
				strWho = ""
			Else
				If rsLeg("LegUpdate") Then
					strWho = "<b>" & rsLeg("Rollcall Name") & "</b>"
				Else
					strWho = _
						rsLeg("Contact Last Name") & ", " & _
						Left(rsLeg("Contact First Name"),1) & "."
				End If
			End If
		
			Response.Write "<tr valign=center bgcolor=" & strBackground & "><td>"
			Response.Write rsLeg("Rollcall Name")
			Response.Write "</td><td>"
			If IsNull(rsLeg("Vote")) Then
				Response.Write strVote(0)
			Else
				Response.Write strVote(rsLeg("Vote"))
			End If

			Response.Write "</td><td>"
			Response.Write strDate & "</td><td>"
			Response.Write strWho & "</td><td>"
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

</div>
</body>
</html>