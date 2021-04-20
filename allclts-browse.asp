<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set rsResults = Server.CreateObject("ADOR.Recordset")
	With rsResults

' SPONSOR LIST
	strSQL = "SELECT DISTINCT [Prime Sponsor] FROM Supplements ORDER BY [Prime Sponsor]"
	.Open strSQL, strConnReadOnly
	aSponsors = .GetRows()
	.Close
	intNumSponsors = UBound(aSponsors,2)+1

' DAILY STATUS LOCATION (COMMITTEE) LIST
	strSQL = _
		"SELECT CommitteeID, [Committee Abbr] FROM [Committees] " & _
		"WHERE LocLevelID IS NOT NULL " & _
		"ORDER BY [Committee Abbr]"
	.Open strSQL, strConnReadOnly
	aLocations = .GetRows()
	.Close
	intNumLocations = UBound(aLocations,2)+1

' LOCATION LEVEL LIST
	strSQL = "SELECT LocLevelID, Level FROM [Location Levels] ORDER BY LocLevelID"
	.Open strSQL, strConnReadOnly
	aLevels = .GetRows()
	.Close
	intNumLevels = UBound(aLevels,2)+1

' LOAD BILL NUMBER AND SUPPLEMENT RANGE
	strSQL = _
		"SELECT MIN([Bill Number]) AS MinBill, MAX([Bill Number]) AS MaxBill FROM [Daily Status]"
	.Open strSQL, strConnReadOnly
	minBill = rsResults("MinBill")
	maxBill = rsResults("MaxBill")
	.Close

	strSQL = _
		"SELECT MIN(S.[Supplement]) AS MinSup, MAX(S.[Supplement]) AS MaxSup " & _
		"FROM [Supplements] S INNER JOIN [System Status] Y ON S.Edition = Y.Edition"
	.Open strSQL, strConnReadOnly
	If Not IsNull(rsResults("MinSup")) Then
		minSup = rsResults("MinSup")
		maxSup = rsResults("MaxSup")
	Else
		minSup = 0
		maxSup = 0
	End If
	.Close

	End With
	Set rsResults = Nothing

' Attempt to load saved values
	BillStart = CInt("0" & Request.Cookies("LegiTrak")("BillStart"))
	BillEnd = CInt("0" & Request.Cookies("LegiTrak")("BillEnd"))
	DigestStart = CInt("0" & Request.Cookies("LegiTrak")("DigestStart"))
	DigestEnd = CInt("0" & Request.Cookies("LegiTrak")("DigestEnd"))
	Source = CInt("0" & Request.Cookies("LegiTrak")("Source"))

	Dim SponsorSel(200), CommSel(100), LevelSel(20)
	intFilterSponsor = CInt("0" & Request.Cookies("LegiTrak")("FilterSponsor"))
	SponsorSel(intFilterSponsor) = " selected"
	intFilterComm = CInt("0" & Request.Cookies("LegiTrak")("FilterComm"))
	CommSel(intFilterComm) = " selected"
	intFilterLevel = CInt("0" & Request.Cookies("LegiTrak")("FilterLevel"))
	LevelSel(intFilterLevel) = " selected"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>u{color:blue;cursor:pointer}</style>
<script src="js/bts.js"></script>
<script src="js/allclts-browse.js"></script>
</head>
<body class=bkg03 onload='init()'>

<form id=BillsForm method=post action='allclts-browse.asp' onsubmit='submitFilters()' style='margin:0'>
<input type=hidden name=minBill value=<%=minBill%>>
<input type=hidden name=maxBill value=<%=maxBill%>>
<input type=hidden name=minSup value=<%=minSup%>>
<input type=hidden name=maxSup value=<%=maxSup%>>
<table id=Bills width=100% border=0 cellspacing=4 cellpadding=0>
<tr align=center class=hdg29>
<td id=BillHead style='cursor:pointer' onclick='selBills()' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Bill Range</td>
<td id=DigestHead style='cursor:pointer' onclick='selDigests()' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Digest Range</td>
<td>Sponsor</td><td>Location</td><td>Level</td></tr>

<tr align=center class=hdg29>
<td><input type=button value=All style='position:relative;top:-2;height:18;width:25;font-size:10;cursor:pointer' onclick='allBills()'>&nbsp;
<input type=text name=BillStart style='width:40' onchange='return isBill(this)'
value=<%=BillStart%>>&nbsp; <input type=text name=BillEnd style='width:40' onchange='return isBill(this)'
value=<%=BillEnd%>></td>
<td><input type=button value=All style='position:relative;top:-2;height:18;width:25;font-size:10;cursor:pointer' onclick='allDigests()'>&nbsp;
<input type=text name=DigestStart style='width:40'
value=<%=DigestStart%>>&nbsp; <input type=text name=DigestEnd style='width:40'
value=<%=DigestEnd%>></td>
<td><select name=FilterSponsor onchange='mark(this)'>
<option value=0<%=SponsorSel(0)%>>-All-
<%
	For i = 1 to intNumSponsors
		Response.Write _
			"<option value=" & i & SponsorSel(i) & ">" & _
			aSponsors(0,i-1)
	Next ' i
%>
</select></td>
<td><select name=FilterComm onchange='mark(this)'>
<option value=0<%=CommSel(0)%>>-All-
<%
	For i = 1 to intNumLocations
		Response.Write _
			"<option value=" & i & CommSel(i) & ">" & _
			aLocations(1,i-1)
	Next ' i
%>
</select></td>
<td><select name=FilterLevel onchange='mark(this)'>
<option value=0<%=LevelSel(0)%>>-All-
<%
	For i = 1 to intNumLevels
		Response.Write _
			"<option value=" & i & LevelSel(i) & ">" & _
			aLevels(1,i-1)
	Next ' i
%>
</select></td>
</tr>
<tr align=center class=shd29><td colspan=5>
<span style='position:relative;top:-2'>To apply filter settings, click </span>
<input type=submit value=Submit></td></tr>
</table>
</form>

<form id=BrowseBillsForm style='margin:0'>
<%
' LOAD SELECTED BILLS FOR BROWSING
	
	If Source = 0 Then
		strSQLWhere = _
			" D.[Bill Number] >=" & BillStart & " AND" & _
			" D.[Bill Number] <=" & BillEnd
	Else
		strSQLWhere = _
			" S.Supplement >=" & DigestStart & " AND" & _
			" S.Supplement <=" & DigestEnd
	End If
	If intFilterSponsor <> 0 Then
		strSQLWhere = strSQLWhere & _
			" AND S.[Prime Sponsor]='" & aSponsors(0,intFilterSponsor-1) & "'"
	End If
	If intFilterComm <> 0 Then
		strSQLWhere = strSQLWhere & _
			" AND D.[CommitteeID]='" & aLocations(0,intFilterComm-1) & "'"
	End If
	If intFilterLevel <> 0 Then
		strSQLWhere = strSQLWhere & _
			" AND C.[LocLevelID]=" & aLevels(0,intFilterLevel-1)
	End If

	strSQLJoin = "[Daily Status] D"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Committees] C ON D.CommitteeID = C.CommitteeID) "
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Supplements with Unique Bill Numbers] S ON D.[Bill Number] = S.[Bill Number])"
	If Source = 1 Then
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN [System Status] Y ON S.Edition = Y.Edition) "
	End If
	strSQL = "SELECT" & _
		" D.Status, D.[Bill Number], D.House, D.Location," & _
		" S.[Prime Sponsor], S.[Long Title] " & _
		"FROM " & strSQLJoin & _
		"WHERE" & strSQLWhere & _
		" ORDER BY D.[Bill Number]"
	Set rsBills=Server.CreateObject("ADOR.Recordset")
	rsBills.Open strSQL, strConnReadOnly
	If rsBills.EOF Then
		strMessage = "Enter a range of bill numbers or digests to get a list of titles and links to their documentation."
	Else
		strMessage = ""
		Response.Write _
			"<table id=Bills border=0 cellspacing=4 cellpadding=0 width=100% class=det00>"
	
		Do Until rsBills.EOF
			If IsNull(rsBills("House")) Or Trim(rsBills("House")) = "" Then
				strHouse=""
			Else
				strHouse=rsBills("House") & ", "
			End If
			Response.Write _
				"<tr class=bkg04 valign=top><td class=lnk70 align=right onclick='quickAdd(" & rsBills("Bill Number") & ")'>" & _
				Replace(rsBills("Status") & rsBills("Bill Number")," ","") & "</td>" & _
				"<td>" & rsBills("Long Title") & "</td>" & _
				"<td><div style='width:80'>" & rsBills("Prime Sponsor") & "</div></td>" & _
				"<td><div style='width:110'>" & strHouse & rsBills("Location") & "</div></td>" & _
				"<td class=lnk40 onclick='lnk(arguments[0],""" & rsBills("Status") & """)'>" & _
				"<u>D</u>_<u>F</u>_<u>A</u></td></tr>"
			rsBills.MoveNext
		Loop
		Response.Write "</table>"
	End If
	rsBills.Close
	Set rsBills = Nothing

	Response.Write _
		"<div id=Bottom class=hdg14 style='position:relative;height:100%;margin:0 4;padding:40'>" & _
		strMessage & "</div>"
%>    
</form>
</body>
</html>
