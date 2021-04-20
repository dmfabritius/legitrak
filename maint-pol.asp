<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"
	Set cmdSQL = CreateObject("ADODB.Connection")
	cmdSQL.Open strConnection

	PolID = CLng("0" & Request.Form("PolID"))
	LegID = CLng("0" & Request.Form("LegID"))
	CandID = CLng("0" & Request.Form("CandID"))
	Dim SeatSel(3)

' ADD/UPDATE POLITICIAN
	If Request.Form("UpdatePol") = "True" Then
		If Request.Form("Birthday") <> "" Then
			strBD = "'" & Request.Form("Birthday") & "'"
		Else
			strBD = "NULL"
		End If
		If PolID = 9999999 Then
			strCommand = _
				"INSERT INTO [Politicians] (LastName) VALUES (" & _
				"'" & TweakQuote(Request.Form("PolLast")) & "')"
			cmdSQL.Execute strCommand, , adExecuteNoRecords
			strCommand = "SELECT MAX(PoliticianID) AS MaxID FROM [Politicians]"
			Set rsResult = cmdSQL.Execute(strCommand)
			PolID = rsResult("MaxID")
			LegID = 0
			CandID = 9999999
			Set rsResult = Nothing
		End If
		strCommand = "UPDATE [Politicians] SET"
		strCommand = strCommand & _	
			" FirstName='" & TweakQuote(Request.Form("PolFirst")) & "'," & _
			" LastName='" & TweakQuote(Request.Form("PolLast")) & "'," & _
			" TaxpayerID='" & TweakQuote(Request.Form("TaxID")) & "'," & _
			" [Rollcall Name]='" & TweakQuote(Request.Form("Rollcall")) & "'," & _
			" [PDCName]='" & TweakQuote(Request.Form("PDCName")) & "'," & _
			" HomeStreet='" & TweakQuote(Request.Form("HomeStreet")) & "'," & _
			" HomeCity='" & TweakQuote(Request.Form("HomeCity")) & "'," & _
			" HomeState='" & Request.Form("HomeState") & "'," & _
			" HomePostalCode='" & Request.Form("HomeZip") & "'," & _
			" HomePhone='" & Request.Form("HomePhone1") & "'," & _
			" HomePhone2='" & Request.Form("HomePhone2") & "'," & _
			" Gender='" & UCase(Left(Trim(Request.Form("Gender")),1)) & "'," & _
			" Birthday=" & strBD & ","
		strCommand = strCommand & _	
			" EmailAddress='" & Request.Form("BusEmail") & "'," & _
			" BusinessStreet='" & TweakQuote(Request.Form("BusStreet1")) & "'," & _
			" BusinessStreet2='" & TweakQuote(Request.Form("BusStreet2")) & "'," & _
			" BusinessCity='" & TweakQuote(Request.Form("BusCity")) & "'," & _
			" BusinessState='" & Request.Form("BusState") & "'," & _
			" BusinessPostalCode='" & Request.Form("BusZip") & "'," & _
			" BusinessPhone='" & Request.Form("BusPhone1") & "'," & _
			" BusinessPhone2='" & Request.Form("BusPhone2") & "'," & _
			" AssistantsName='" & TweakQuote(Request.Form("Assistant")) & "'," & _
			" AideName='" & TweakQuote(Request.Form("Aide")) & "',"
		strCommand = strCommand & _	
			" [Campaign Name]='" & TweakQuote(Request.Form("CampName")) & "'," & _
			" [Campaign URL]='" & Request.Form("CampWeb") & "'," & _
			" [Campaign Email]='" & Request.Form("CampEmail") & "'," & _
			" CampaignStreet='" & TweakQuote(Request.Form("CampStreet1")) & "'," & _
			" CampaignStreet2='" & TweakQuote(Request.Form("CampStreet2")) & "'," & _
			" CampaignCity='" & TweakQuote(Request.Form("CampCity")) & "'," & _
			" CampaignState='" & Request.Form("CampState") & "'," & _
			" CampaignPostalCode='" & Request.Form("CampZip") & "'," & _
			" CampaignPhone='" & Request.Form("CampPhone1") & "'," & _
			" CampaignFax='" & Request.Form("CampPhone2") & "' " & _
			"WHERE PoliticianID=" & PolID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
	End If

' DELETE POLITICIAN
	If Request.Form("UpdatePol") = "Delete" Then
		strCommand = _
			"DELETE FROM [Politicians] " & _
			"WHERE PoliticianID=" & PolID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		PolID = 0
	End If

' ADD/UPDATE LEGISLATOR DETAIL RECORD
	If Request.Form("UpdateLeg") = "True" Then
		If LegID <> 9999999 Then
			strCommand = _
				"UPDATE [Legislator Details] SET" & _
				" Party='" & Request.Form("Party") & "'," & _
				" DistrictID=" & Request.Form("District") & "," & _
				" Seat=" & Request.Form("Seat") & "," & _
				" BeginDate='" & Request.Form("BeginDate") & "'," & _
				" EndDate='" & Request.Form("EndDate") & "'," & _
				" [Leadership Position]='" & TweakQuote(Request.Form("Position")) & "' " & _
				"WHERE LegislatorID=" & LegID
			cmdSQL.Execute strCommand, , adExecuteNoRecords
		Else
			strCommand = _
				"INSERT INTO [Legislator Details] " & _
				"(PoliticianID, Party, DistrictID, Seat," & _
				" BeginDate, EndDate, [Leadership Position]) " & _
				"VALUES (" & _
				PolID & "," & _
				"'" & Request.Form("Party") & "'," & _
				Request.Form("District") & "," & _
				Request.Form("Seat") & "," & _
				"'" & Request.Form("BeginDate") & "'," & _
				"'" & Request.Form("EndDate") & "'," & _
				"'" & TweakQuote(Request.Form("Position")) & "')"
			cmdSQL.Execute strCommand, , adExecuteNoRecords
		End If
		LegID = 0
	End If

' DELETE LEGISLATOR DETAIL RECORD
	If Request.Form("UpdateLeg") = "Delete" Then
		strCommand = _
			"DELETE FROM [Legislator Details] " & _
			"WHERE LegislatorID=" & LegID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		LegID = 0
	End If

' ADD/UPDATE CANDIDATE DETAIL RECORD
	If Request.Form("UpdateCand") = "True" Then

		strPriVotes = Request.Form("PriVotes")
		If strPriVotes = "" Then strPriVotes = "NULL"
		strPriPct = Request.Form("PriPct")
		If strPriPct = "" Then strPriPct = "NULL"
		strDist = Request.Form("District")
		If strDist = "0" Or strDist = "" Then strDist = "NULL"
		
		intPassed = CInt("0" & Request.Form("Passed"))
		intWD = CInt("0" & Request.Form("Withdrawn"))
		intIncumb = CInt("0" & Request.Form("Incumbent"))

		If CandID <> 9999999 Then
'				" Position='" & TweakQuote(Request.Form("Position")) & "' " & _

			strCommand = _
				"UPDATE [Candidate Details] SET" & _
				" Party='" & Request.Form("Party") & "'," & _
				" DistrictID=" & strDist & "," & _
				" Seat=" & Request.Form("Seat") & "," & _
				" Year=" & Request.Form("Year") & "," & _
				" PrimaryCount=" & strPriVotes & "," & _
				" PrimaryPct=" & strPriPct & "," & _
				" PassedPrimary=" & intPassed & "," & _
				" Withdrawn=" & intWD & "," & _
				" Incumbent=" & intIncumb & "," & _
				" SWRaceID=" & CInt(Request.Form("Position")) & _
				" WHERE CandidateID=" & CandID
			cmdSQL.Execute strCommand, , adExecuteNoRecords
		Else
			strCommand = _
				"INSERT INTO [Candidate Details] " & _
				"(PoliticianID, Party, DistrictID, Seat," & _
				" Year, PrimaryCount, PrimaryPct, PassedPrimary, Withdrawn, Incumbent, SWRaceID) " & _
				"VALUES (" & _
				PolID & "," & _
				"'" & Request.Form("Party") & "'," & _
				strDist & "," & _
				Request.Form("Seat") & "," & _
				Request.Form("Year") & "," & _
				strPriVotes & "," & _
				strPriPct & "," & _
				intPassed & "," & _
				intWD & "," & _
				intIncumb & "," & _
				Request.Form("Position") & ")"
			cmdSQL.Execute strCommand, , adExecuteNoRecords
		End If
		CandID = 0
	End If

' DELETE CANDIDATE DETAIL RECORD
	If Request.Form("UpdateCand") = "Delete" Then
		strCommand = _
			"DELETE FROM [Candidate Details] " & _
			"WHERE CandidateID=" & CandID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		CandID = 0
	End If

	intLegCount = -1
	intCandCount = -1
	If PolID <> 0 And PolID <> 9999999 Then
		Set rs=Server.CreateObject("ADOR.Recordset")

' LOAD POLITICIAN LEGISLATOR DETAILS
		strSQL = 	_
			"SELECT LegislatorID, PoliticianID, Party, DistrictID, Seat," & _
			" BeginDate, EndDate, [Leadership Position] " & _
			"FROM [Legislator Details] " & _
			"WHERE PoliticianID=" & PolID & " ORDER BY BeginDate"
		rs.Open strSQL, strConnection
		If Not rs.EOF Then
			aLeg = rs.GetRows()
			intLegCount = UBound(aLeg,2)
			For i = 0 to intLegCount
				If aLeg(0,i) = LegID Then LegIndex = i
			Next 'i
		End If
		rs.Close
' LOAD POLITICIAN CANDIDATE DETAILS
		strSQL = 	_
			"SELECT CandidateID, PoliticianID, Year, DistrictID, Seat, Party," & _
			" PrimaryCount, PrimaryPct, PassedPrimary, GeneralCount, GeneralPct," & _
			" SWRaceID, Withdrawn, Incumbent " & _
			"FROM [Candidate Details] " & _
			"WHERE PoliticianID=" & PolID & " ORDER BY [Year]"
		rs.Open strSQL, strConnection
		If Not rs.EOF Then
			aCand = rs.GetRows()
			intCandCount = UBound(aCand,2)
			For i = 0 to intCandCount
				If aCand(0,i) = CandID Then CandIndex = i
			Next 'i
		End If
		rs.Close
' LOAD STATE-WIDE RACE DESCRIPTIONS
		strSQL = "SELECT SWRaceID, Race, Abbr FROM [State-Wide Races]"
		rs.Open strSQL, strConnection
		aSWR = rs.GetRows()
		intSWR = UBound(aSWR,2)
		rs.Close
		Set rs = Nothing
	End If

	cmdSQL.Close
	Set cmdSQL = Nothing
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var pdf,ldf,cdf
function init(){
	pdf=document.getElementById("PolDetailForm")
	ldf=document.getElementById("LegDetailForm")
	cdf=document.getElementById("CandDetailForm")
	selectTab(1)

	if (0<%=LegID%> != 0) ldf.BeginDate.focus()
	else if (<%=CandID%> != 0)	cdf.Year.focus()
	else if (<%=PolID%>  != 0) pdf.PolFirst.focus()
}
function PolSelect(){
	pdf.UpdatePol.value="False"
	pdf.submit()
}
function PolCancel(){
	pdf.UpdatePol.value="False"
	pdf.PolID.disabled=true
	pdf.submit()
}
function PolDelete(){
	if (confirm("Click OK to confirm delete for this Politician and all their associated detail records.")){
		pdf.UpdatePol.value="Delete"
		pdf.submit()
	}
}
function legSelect(c){
	if (c != -1) ldf.LegID.value=c
	ldf.UpdateLeg.value="False"
	ldf.submit()
}
function legCancel(){
	ldf.UpdateLeg.value="False"
	ldf.LegID.disabled=true
	ldf.submit()
}
function legDelete(){
	if (confirm("Click OK to confirm delete for this Legislative detail record.")){
		ldf.UpdateLeg.value="Delete"
		ldf.submit()
	}
}
function candSelect(c){
	if (c != -1) cdf.CandID.value=c
	cdf.UpdateCand.value="False"
	cdf.submit()
}
function candCancel(){
	cdf.UpdateCand.value="False"
	cdf.CandID.disabled=true
	cdf.submit()
}
function candDelete(){
	if (confirm("Click OK to confirm delete for this Candidate detail record.")){
		cdf.UpdateCand.value="Delete"
		cdf.submit()
	}
}
function lastnameChange(){
	pdf.Rollcall.value=pdf.PolLast.value
}
function begindateChange(){
	ldf.EndDate.value='12/31/2299'
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:10'>


<!--
'------------------------------------------------
' POLITICIAN LIST
'------------------------------------------------
-->
<form id=PolDetailForm action="maint-pol.asp" method=post>
<input type=hidden name=UpdatePol value=True>
<span class=hdg24>Politician Details:</span>
<select name=PolID style='width:250' onchange='pdf.pdfSubmit.disabled=true'>
<option value=9999999>-Add New-
<%
	strSQL = "SELECT PoliticianID, LastName, ISNULL(FirstName,'') FirstName FROM [Politicians] ORDER BY LastName, FirstName"
	Set rsPols=Server.CreateObject("ADOR.Recordset")
	rsPols.Open strSQL, strConnection
	Do Until rsPols.EOF
		Response.Write "<option value=" & rsPols("PoliticianID")
		If PolID=rsPols("PoliticianID") Then Response.Write " selected"
		Response.Write ">" & rsPols("LastName")
		If rsPols("FirstName") <> "" Then Response.Write ", " & rsPols("FirstName")
		Response.Write " (ID:" & rsPols("PoliticianID") & ")"
		rsPols.MoveNext
	Loop
	rsPols.Close
	Set rsPols = Nothing
%>
</select>
&nbsp; <input type=button value="Select" onclick='PolSelect()'>

<!--
'------------------------------------------------
' POLITICIAN DETAILS
'------------------------------------------------
-->
<%
	If PolID <> 0 Then
		If PolID <> 9999999 Then
			strSQL = 	_
				"SELECT" & _
				" FirstName, LastName, [Rollcall Name]," & _
				" HomeStreet, HomeCity, HomeState, HomePostalCode," & _
				" HomePhone, HomePhone2, Gender, Birthday," & _
				" EmailAddress, BusinessStreet, BusinessStreet2," & _
				" BusinessCity, BusinessState, BusinessPostalCode," & _
				" BusinessPhone, BusinessPhone2, AssistantsName, AideName," & _
				" [Campaign Name], [Campaign URL], [Campaign Email]," & _
				" CampaignStreet, CampaignStreet2, CampaignCity, CampaignState," & _
				" CampaignPostalCode, CampaignPhone, CampaignFax, PDCName, TaxpayerID " & _
				"FROM [Politicians] WHERE PoliticianID=" & PolID
			Set rsPol=Server.CreateObject("ADOR.Recordset")
			rsPol.Open strSQL, strConnection
			aPol = rsPol.GetRows()
			rsPol.Close
			Set rsPol = Nothing

			' Set the roll call name equal to the last name if it is blank
			If Len(aPol(2,0)) = 0 Then aPol(2,0)=aPol(1,0)
			aPol(9,0) = Trim(aPol(9,0))
		Else
			Dim aPol(32,1)
		End If
%>
<div class=box20 style='padding:3'>
<table border=0 cellpadding=0 cellspacing=3 width=842><tr>
<!--Personal/Home-->
<td><table border=0 cellpadding=0 cellspacing=0 width=278 class=det00 style='padding-left:3'>
<col width=60 align=right><col width=220>
<tr><td></td><td class=shd24>Personal/Home Contact Information</td></tr>
<tr><td>Name:</td><td>
<input name=PolFirst type=text style='width:95' value="<%=aPol(0,0)%>">
<input name=PolLast type=text style='width:110' value="<%=aPol(1,0)%>" onchange='lastnameChange()'>
</td></tr>
<tr><td>TaxID#:</td><td><input name=TaxID type=text style='width:208' value="<%=aPol(32,0)%>"></td></tr>
<tr><td>PDC:</td><td><input name=PDCName type=text style='width:208' value="<%=aPol(31,0)%>"></td></tr>
<tr><td>Roll call:</td><td><span style='width:98'></span>
<input name=Rollcall type=text style='width:110' value="<%=aPol(2,0)%>"></td></tr>
<tr valign=top><td>Address:</td><td>
<input name=HomeStreet type=text style='width:208' value="<%=aPol(3,0)%>"><br>
<input name=HomeCity type=text style='width:95' value="<%=aPol(4,0)%>">
<input name=HomeState type=text style='width:30' value="<%=aPol(5,0)%>">
<input name=HomeZip type=text style='width:75' value="<%=aPol(6,0)%>">
</td></tr>
<tr><td>Olympia#:</td><td><input name=HomePhone1 type=text style='width:105' value="<%=aPol(7,0)%>"></td></tr>
<tr><td>District#:</td><td><input name=HomePhone2 type=text style='width:105' value="<%=aPol(8,0)%>"></td></tr>
<tr><td>Gender:</td><td><input name=Gender type=text style='width:20' value="<%=aPol(9,0)%>"></td></tr>
</table></td>

<!--Business-->
<td><table border=0 cellpadding=0 cellspacing=0 width=278 class=det00 style='padding-left:3'>
<col width=60 align=right><col width=220>
<tr><td></td><td class=shd24>Business Contact Information</td></tr>
<tr><td>E-Mail:</td><td><input name=BusEmail type=text style='width:208' value="<%=aPol(11,0)%>"></td></tr>
<tr valign=top><td>Address:</td><td>
<input name=BusStreet1 type=text style='width:208' value="<%=aPol(12,0)%>"><br>
<input name=BusStreet2 type=text style='width:208' value="<%=aPol(13,0)%>"><br>
<input name=BusCity type=text style='width:95' value="<%=aPol(14,0)%>">
<input name=BusState type=text style='width:30' value="<%=aPol(15,0)%>">
<input name=BusZip type=text style='width:75' value="<%=aPol(16,0)%>">
</td></tr>
<tr><td>Olympia#:</td><td><input name=BusPhone1 type=text style='width:105' value="<%=aPol(17,0)%>"></td></tr>
<tr><td>District#:</td><td><input name=BusPhone2 type=text style='width:105' value="<%=aPol(18,0)%>"></td></tr>
<tr><td>Assistant:</td><td><input name=Assistant type=text style='width:138' value="<%=aPol(19,0)%>"></td></tr>
<tr><td>Aide:</td><td><input name=Aide type=text style='width:138' value="<%=aPol(20,0)%>"></td></tr>
</table></td>

<!--Campaign-->
<td><table border=0 cellpadding=0 cellspacing=0 width=278 class=det00 style='padding-left:3'>
<col width=60 align=right><col width=220>
<tr><td></td><td class=shd24>Campaign Contact Information</td></tr>
<tr><td>Name:</td><td><input name=CampName type=text style='width:208' value="<%=aPol(21,0)%>"></td></tr>
<tr><td>Web Site:</td><td><input name=CampWeb type=text style='width:208' value="<%=aPol(22,0)%>"></td></tr>
<tr><td>E-Mail:</td><td><input name=CampEmail type=text style='width:208' value="<%=aPol(23,0)%>"></td></tr>
<tr valign=top><td>Address:</td><td>
<input name=CampStreet1 type=text style='width:208' value="<%=aPol(24,0)%>"><br>
<input name=CampStreet2 type=text style='width:208' value="<%=aPol(25,0)%>"><br>
<input name=CampCity type=text style='width:95' value="<%=aPol(26,0)%>">
<input name=CampState type=text style='width:30' value="<%=aPol(27,0)%>">
<input name=CampZip type=text style='width:75' value="<%=aPol(28,0)%>">
</td></tr>
<tr><td>Phone:</td><td><input name=CampPhone1 type=text style='width:105' value="<%=aPol(29,0)%>"></td></tr>
<tr><td>Fax:</td><td><input name=CampPhone2 type=text style='width:105' value="<%=aPol(30,0)%>"></td></tr>
</table></td>

<!--End of Politician details-->
</tr></table>

<center>
<span style='height:25'></span>
<input id=pdfSubmit type=submit value=Submit>
<span style='width:200'></span><input type=button onclick='PolCancel()' value=Cancel>
<span style='width:200'></span><input type=button onclick='PolDelete()' value=Delete>
</center>

</div>
<%
	End If
%>
</form>

<!--
'------------------------------------------------
' POLITICIAN SUMMARY
'------------------------------------------------
-->
<%
	If PolID <> 0 And PolID <> 9999999 And intLegCount+intCandCount <> -2 Then
		Response.Write _
			"<span class=hdg24>Politician Summary</span>" & _
			"<div class=box20 style='padding:10'>"
	End If

' LEGISLATOR SUMMARY
	If 	intLegCount <> -1 Then
		Response.Write _
			"<table border=0 cellpadding=0 cellspacing=0 width=800 class=det00>" & _
			"<col width=120><col width=100><col width=100>" & _
			"<col width=40><col width=50>" & _
			"<col width=60><col width=50><col width=210>"
		Response.Write _
			"<tr align=center class=shd24 style='text-decoration:underline'><td align=right><b>Legislator</b></td>" & _
			"<td>Begin Date</td>" & _
			"<td>End Date</td>" & _
			"<td></td>" & _
			"<td>Party</td>" & _
			"<td>District</td>" & _
			"<td>Seat</td>" & _
			"<td align=left>Leadership Position</td></tr>"
		For i = 0 to intLegCount
			Response.Write "<tr align=center><td></td>"
			Response.Write _
				"<td><div onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
				" style='cursor:hand' onclick='legSelect(" & aLeg(0,i) & ")'>" & _
				aLeg(5,i) & "</div></td>"
			Response.Write "<td>" & aLeg(6,i) & "</td>"
			Response.Write "<td></td>"
			Response.Write "<td>" & aLeg(2,i) & "</td>"
			Response.Write "<td>" & aLeg(3,i) & "</td>"
			Response.Write "<td>" & aLeg(4,i) & "</td>"
			Response.Write "<td align=left>" & aLeg(7,i) & "</td>"
			Response.Write "</tr>"
		Next 'i
		Response.Write "</table><br>"
	End If

' CANDIDATE SUMMARY
	If 	intCandCount <> -1 Then
		Response.Write _
			"<table border=0 cellpadding=0 cellspacing=0 width=800 class=det00>" & _
			"<col width=120><col width=50><col width=70><col width=70>" & _
			"<col width=50><col width=50>" & _
			"<col width=60><col width=50><col width=210>"
		Response.Write _
			"<tr align=center class=shd24 style='text-decoration:underline'><td align=right><b>Candidate</b></td>" & _
			"<td>Year</td>" & _
			"<td>Pri Votes</td>" & _
			"<td>Pri Pct</td>" & _
			"<td>Passed</td>" & _
			"<td>Party</td>" & _
			"<td>District</td>" & _
			"<td>Seat</td>" & _
			"<td align=left>State-Wide Race</td></tr>"
		For i = 0 to intCandCount
			If IsNumeric(aCand(7,i)) Then
				strPct = FormatPercent(aCand(7,i),1)
			Else
				strPct = ""
			End If
			If aCand(12,i) = "True" Then
				strStyle= " style='color:#C0C0C0'"
			Else
				strStyle = ""
			End If
			Response.Write "<tr align=center><td></td>"
			Response.Write _
				"<td" & strStyle & "><div onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
				" style='cursor:hand' onclick='candSelect(" & aCand(0,i) & ")'>" & _
				aCand(2,i) & "</div></td>"
			Response.Write "<td" & strStyle & ">" & aCand(6,i) & "</td>"
			Response.Write "<td" & strStyle & ">" & strPct & "</td>"
			Response.Write "<td" & strStyle & ">" & aCand(8,i) & "</td>"
			Response.Write "<td" & strStyle & ">" & aCand(5,i) & "</td>"
			Response.Write "<td" & strStyle & ">" & aCand(3,i) & "</td>"
			Response.Write "<td" & strStyle & ">" & aCand(4,i) & "</td>"
			Response.Write "<td" & strStyle & " align=left>" & aSWR(1,aCand(11,i)) & "</td>"
			Response.Write "</tr>"
		Next 'i
		Response.Write "</table>"


	End If

	If PolID <> 0 And PolID <> 9999999 And intLegCount+intCandCount <> -2 Then
		Response.Write "</div>"
	End If
%>

<!--
'------------------------------------------------
' LEGISLATOR DETAILS
'------------------------------------------------
-->
<form id=LegDetailForm action="maint-pol.asp" method=post>
<input type=hidden name=PolID value=<%=PolID%>>
<input type=hidden name=UpdateLeg value=True>
<%
	If PolID <> 0 and PolID <> 9999999 Then
		Response.Write _
			"<span class=hdg24 style='width:140'>Legislator Details:</span>" & _
			"<select name=LegID style='width:150'>" & _
			"<option value=9999999>-Add New-"

		If 	intLegCount <> -1 Then
			For i = 0 to intLegCount
				Response.Write "<option value=" & aLeg(0,i)
				If LegID = aLeg(0,i) Then Response.Write " selected"
				Response.Write ">Term starting " & aLeg(5,i)
			Next 'i
		End If

		Response.Write _
			"</select>" & _
			"&nbsp; <input type=button value='Select' onclick='legSelect(-1)'>"
	End If
	
	If LegID <> 0 Then
		If LegID <> 9999999 Then
			strParty = aLeg(2,LegIndex)
			strDistrict = aLeg(3,LegIndex)
			strSeat = aLeg(4,LegIndex)
			strBeginDate = aLeg(5,LegIndex)
			strEndDate = aLeg(6,LegIndex)
			strPosition = aLeg(7,LegIndex)
		Else
			If intLegCount <> -1 Then
				strDistrict = aCand(3,intLegCount)
				strSeat = aCand(4,intLegCount)
				strParty = aCand(5,intLegCount)
			End If
		End If
		If IsNull(strSeat) Then
			SeatSel(3) = " selected"
		Else
			SeatSel(strSeat) = " selected"
		End If
%>
<div class=box20 style='padding:7'>
<table border=0 cellpadding=0 cellspacing=0 width=800 class=det00 style='padding-left:3'>
<col width=80 align=right><col width=100>
<col width=70 align=right><col width=100>
<col width=60 align=right><col width=100>
<col width=70 align=right><col width=220>

<tr><td>Begin Date:</td><td><input name=BeginDate type=text tabindex=1 style='width:80' value="<%=strBeginDate%>" onchange='begindateChange()'></td>
<td>Party:</td><td><input name=Party type=text tabindex=3 style='width:40' value="<%=strParty%>"></td>
<td>Seat:</td><td>
<select name=Seat tabindex=8 style='width:200'>
<option value=0<%=SeatSel(0)%>>Senate
<option value=1<%=SeatSel(1)%>>House Position 1
<option value=2<%=SeatSel(2)%>>House Position 2
</select></td></tr>

<tr><td>End Date:</td><td><input name=EndDate type=text tabindex=2 style='width:80' value="<%=strEndDate%>"></td>
<td>District:</td><td><input name=District type=text tabindex=5 style='width:40' value="<%=strDistrict%>"></td>
<td>Position:</td><td><input name=Position type=text tabindex=7 style='width:200' value="<%=strPosition%>"></td></tr>

<tr style='height:25'><td colspan=8 align=center valign=bottom>
<input type=submit value=Submit>
<span style='width:200'></span><input type=button onclick='legCancel()' value=Cancel>
<span style='width:200'></span><input type=button onclick='legDelete()' value=Delete>
</td></tr>
</table>
</div>
<%
	End If
%>
</form>

<!--
'------------------------------------------------
' CANDIDATE DETAILS
'------------------------------------------------
-->
<form id=CandDetailForm action="maint-pol.asp" method=post>
<input type=hidden name=PolID value=<%=PolID%>>
<input type=hidden name=UpdateCand value=True>
<%
	If PolID <> 0 and PolID <> 9999999 Then
		Response.Write _
			"<span class=hdg24 style='width:140'>Candidate Details:</span>" & _
			"<select name=CandID style='width:150'>" & _
			"<option value=9999999>-Add New-"

		If 	intCandCount <> -1 Then
			For i = 0 to intCandCount
				Response.Write "<option value=" & aCand(0,i)
				If CandID = aCand(0,i) Then Response.Write " selected"
				Response.Write ">Election Year " & aCand(2,i)
			Next 'i
		End If

		Response.Write _
			"</select>" & _
			"&nbsp; <input type=button value='Select' onclick='candSelect(-1)'>"
	End If
	
	If CandID <> 0 Then
		If CandID <> 9999999 Then
			strYear = aCand(2,CandIndex)
			strDistrict = aCand(3,CandIndex)
			strSeat = aCand(4,CandIndex)
			strParty = aCand(5,CandIndex)
			strPriVotes = aCand(6,CandIndex)
			strPriPct = aCand(7,CandIndex)
			If aCand(8,CandIndex) = "True" Then
				strPassed = " checked"
			Else
				strPassed = ""
			End If
			intSWRace = aCand(11,CandIndex)
			If aCand(12,CandIndex) = "True" Then
				strWD = " checked"
			Else
				strWD = ""
			End If
			If aCand(13,CandIndex) = 1 Then
				strIncumb = " checked"
			Else
				strIncumb = ""
			End If
		Else
			If intCandCount <> -1 Then
				strDistrict = aCand(3,intCandCount)
				strSeat = aCand(4,intCandCount)
				strParty = aCand(5,intCandCount)
			End If
		End If
		If IsNull(strSeat) Then
			SeatSel(3) = " selected"
		Else
			SeatSel(strSeat) = " selected"
		End If
%>
<div class=box20 style='padding:7'>
<table border=0 cellpadding=0 cellspacing=0 width=800 class=det00 style='padding-left:3'>
<col width=80 align=right><col width=100>
<col width=70 align=right><col width=100>
<col width=60 align=right><col width=100>
<col width=70 align=right><col width=220>

<tr><td>Year:</td><td><input name=Year type=text tabindex=1 style='width:80' value="<%=strYear%>"></td>
<td>Party:</td><td><input name=Party type=text tabindex=5 style='width:40' value="<%=strParty%>"></td>
<td>Seat:</td><td>
<select name=Seat tabindex=8 style='width:250'>
<option value=0<%=SeatSel(0)%>>Senate
<option value=1<%=SeatSel(1)%>>House Position 1
<option value=2<%=SeatSel(2)%>>House Position 2
<option value=NULL<%=SeatSel(3)%>>N/A (State-Wide)
</select></td></tr>

<tr><td>Pri Votes:</td><td><input name=PriVotes type=text tabindex=2 style='width:80' value="<%=strPriVotes%>"></td>
<td>District:</td><td><input name=District type=text tabindex=7 style='width:40' value="<%=strDistrict%>" id="Text1"></td>
<td>State Race:</td><td><select name=Position tabindex=9 style='width:250'>
<%
	For i = 0 to intSWR
		Response.Write "<option value=" & aSWR(0,i)
		If intSWRace = i Then Response.Write " selected"
		Response.Write ">" & aSWR(1,i)
	Next 'i
%>
</select></td></tr>

<tr><td>Pri Pct:</td><td><input name=PriPct type=text tabindex=3 style='width:80' value="<%=strPriPct%>"></td></tr>
<tr>
<td>Passed:</td><td><input name=Passed type=checkbox tabindex=4 value=1 <%=strPassed%>></td>
<td>Withdrawn:</td><td><input name=Withdrawn type=checkbox tabindex=10 value=1 <%=strWD%>></td>
<td>Incumbent:</td><td><input name=Incumbent type=checkbox tabindex=11 value=1 <%=strIncumb%>></td>
</tr>

<tr style='height:25'><td colspan=8 align=center valign=bottom>
<input type=submit value=Submit>
<span style='width:200'></span><input type=button onclick='candCancel()' value=Cancel>
<span style='width:200'></span><input type=button onclick='candDelete()' value=Delete>
</td></tr>
</table>
</div>
<%
	End If
%>
</form>

</body>
</html>
