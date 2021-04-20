<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnReadOnly

    Dim strOrdinal(49)
    For i = 1 to 49
        If i >= 11 And i <= 13 Then
            strOrdinal(i) = i & "th"
        Else
            Select Case (i Mod 10)
                Case 1: strOrdinal(i) = i & "st"
                Case 2: strOrdinal(i) = i & "nd"
                Case 3: strOrdinal(i) = i & "rd"
                Case Else: strOrdinal(i) = i & "th"
            End Select
        End If
    Next 'i

' LOAD CONTRIBUTION GROUPS LIST
	strSQL = "SELECT DISTINCT [Group] FROM [Client Politician Comments] WHERE ClientID=" & ClientID & " ORDER BY [Group]"
	Set rsResults = cxnSQL.Execute(strSQL)
	If Not rsResults.EOF Then
		aGroups = rsResults.GetRows()
	Else
		Dim aGroups(0,0)
	End If
	intNumGroups = UBound(aGroups,2)

' LOAD CAMPAIGN STATE-WIDE RACES LIST
	strSQL = "SELECT SWRaceID, Race, Abbr FROM [State-Wide Races] WHERE SWRaceID <> 0"
	Set rsResults = cxnSQL.Execute(strSQL)
	aSWRaces = rsResults.GetRows()
	intNumSWRaces = UBound(aSWRaces,2)+1

' LOAD COMMITTEE LIST
	strSQL = _
		"SELECT CommitteeID, [Committee Name], LEFT([House],1) FROM [Committees] " & _
		"WHERE ISNULL(House, '') <> '' ORDER BY [Committee Name]"
	Set rsResults = cxnSQL.Execute(strSQL)
	aComms = rsResults.GetRows()
	intNumComms = UBound(aComms,2)+1

	Set rsResults = Nothing
	Set cxnSQL = Nothing

	
' SET FILTERS
	Dim YearSel(6), RaceSel(200), StatusSel(4), PartySel(3)
	Dim NameSel(11), DstSel(50), GrpSel(100), RecSel(6), ActSel(6)

	intYear=DatePart("yyyy",Date)
	intFilterYear = CInt("0" & Request.Cookies("LegiTrak")("FilterYear"))
	YearSel(intFilterYear) = " selected"
	Select Case intFilterYear
		Case 0 : strWhere = "C.[Year]>=" & intYear
		Case 1 : strWhere = "C.[Year]=" & intYear-2
		Case 2 : strWhere = "C.[Year]=" & intYear-1
		Case 3 : strWhere = "C.[Year]=" & intYear
		Case 4 : strWhere = "C.[Year]=" & intYear+1
		Case 5 : strWhere = "C.[Year]=" & intYear+2
		Case 6 : strWhere = "C.[Year]=" & intYear+3
	End Select

	intFilterRace = CInt("0" & Request.Cookies("LegiTrak")("FilterRace"))
	RaceSel(intFilterRace) = " selected"
	'       If intFilterRace >= 1 And intFilterRace <= 49 Then
	'       	strWhere = strWhere & " AND C.DistrictID=" & intFilterRace
	'       Else
	If intFilterRace = 50 Then
		strWhere = strWhere & " AND ISNULL(C.FirstName,'')=''"
	ElseIf intFilterRace > 50 And intFilterRace < 51+intNumSWRaces Then
		strWhere = strWhere & " AND C.SWRaceID='" & aSWRaces(0,intFilterRace-51) & "'"
	ElseIf intFilterRace <> 0 Then
		strWhere = strWhere & " AND CM.CommitteeID=" & aComms(0,intFilterRace-51-intNumSWRaces)
	End If
	
	intFilterStatus = CInt("0" & Request.Cookies("LegiTrak")("FilterStatus"))
	StatusSel(intFilterStatus) = " selected"
	Select Case intFilterStatus
		Case 1 : strWhere = strWhere & " AND C.Incumbent=1"
		Case 2 : strWhere = strWhere & " AND C.Seat=0 AND C.Incumbent=1"
		Case 3 : strWhere = strWhere & " AND (C.Seat=1 OR C.Seat=2) AND C.Incumbent=1"
		Case 4 : strWhere = strWhere & " AND C.Incumbent=0"
	End Select

	intFilterParty = CInt("0" & Request.Cookies("LegiTrak")("FilterParty"))
	PartySel(intFilterParty) = " selected"
	Select Case intFilterParty
		Case 1 : strWhere = strWhere & " AND C.Party='D'"
		Case 2 : strWhere = strWhere & " AND C.Party='R'"
		Case 3 : strWhere = strWhere & " AND C.Party<>'D' AND C.Party<>'R'"
	End Select

	intFilterName = CInt("0" & Request.Cookies("LegiTrak")("FilterName"))
	NameSel(intFilterName) = " selected"
	Select Case intFilterName
		Case 1 : strWhere = strWhere & " AND C.[LastName] > 'A' AND C.[LastName] < 'BZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 2 : strWhere = strWhere & " AND C.[LastName] > 'C' AND C.[LastName] < 'DZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 3 : strWhere = strWhere & " AND C.[LastName] > 'E' AND C.[LastName] < 'GZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 4 : strWhere = strWhere & " AND C.[LastName] > 'H' AND C.[LastName] < 'JZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 5 : strWhere = strWhere & " AND C.[LastName] > 'K' AND C.[LastName] < 'LZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 6 : strWhere = strWhere & " AND C.[LastName] > 'M' AND C.[LastName] < 'OZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 7 : strWhere = strWhere & " AND C.[LastName] > 'P' AND C.[LastName] < 'RZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 8 : strWhere = strWhere & " AND C.[LastName] > 'S' AND C.[LastName] < 'TZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 9 : strWhere = strWhere & " AND C.[LastName] > 'U' AND C.[LastName] < 'ZZZZ' AND ISNULL(C.[FirstName],'')<>''"
		Case 10 : strWhere = strWhere & " AND ISNULL(C.[FirstName],'')=''"
		Case 11 : strWhere = strWhere & " AND ISNULL(C.[FirstName],'')='' AND C.Party <> 'X'"
	End Select






	intFilterDst = CInt("0" & Request.Cookies("LegiTrak")("FilterDst"))
	If intFilterDst > 49 Then intFilterGrp = 0
	DstSel(intFilterDst) = " selected"
	If intFilterDst <> 0 Then strWhere = strWhere & " AND C.[DistrictID]=" & intFilterDst











	intFilterGrp = CInt("0" & Request.Cookies("LegiTrak")("FilterGrp"))
	If intFilterGrp > intNumGroups Then intFilterGrp = 0
	GrpSel(intFilterGrp) = " selected"
	If intFilterGrp <> 0 Then strWhere = strWhere & " AND CC.[Group]=" & aGroups(0,intFilterGrp)

	intFilterRec = CInt("0" & Request.Cookies("LegiTrak")("FilterRec"))
	RecSel(intFilterRec) = " selected"
	Select Case intFilterRec
		Case 1 : strWhere = strWhere & " AND CC.TotRec>0"
		Case 2 : strWhere = strWhere & " AND CC.TotRec>0 AND CC.TotRec<800"
		Case 3 : strWhere = strWhere & " AND CC.TotRec>799 AND CC.TotRec<1600"
		Case 4 : strWhere = strWhere & " AND CC.TotRec=1600"
		Case 5 : strWhere = strWhere & " AND CC.TotRec>1600"
		Case 6 : strWhere = strWhere & " AND CC.TotRec=0"
	End Select

	intFilterAct = CInt("0" & Request.Cookies("LegiTrak")("FilterAct"))
	ActSel(intFilterAct) = " selected"
	Select Case intFilterAct
		Case 1 : strWhere = strWhere & " AND AC.Actual>0"
		Case 2 : strWhere = strWhere & " AND AC.Actual>0 AND AC.Actual<800"
		Case 3 : strWhere = strWhere & " AND AC.Actual>799 AND AC.Actual<1600"
		Case 4 : strWhere = strWhere & " AND AC.Actual=1600"
		Case 5 : strWhere = strWhere & " AND AC.Actual>1600"
		Case 6 : strWhere = strWhere & " AND AC.Actual=0"
	End Select

' SET DEFAULT FILTER
	If Len(strWhere) < 20 Then
		intFilterName = 1
		NameSel(1) = " selected"
		strWhere = strWhere & " AND C.[LastName] > 'A' AND C.[LastName] < 'BZZZ' AND ISNULL(C.[FirstName],'')<>''"
	End If

' DETERMINE SORT ORDER
	strOrder = Request.Cookies("LegiTrak")("CandOrderField")
	If Len(strOrder) = 0 Then strOrder = "C.[LastName], C.[FirstName]"
	If intFilterDst <> 0 Then strOrder = "C.Seat," & strOrder
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>
.c0{overflow:hidden;height:14px;margin:3}
table,td{border:2px solid white}
p{margin:0}p table{border:0px}p td{border:0px}
</style>
<script src="js/bts.js"></script>
<script src="js/client-campaign.js"></script>
</head>
<body onload='init()' onclick='hideDetail(arguments[0])' onscroll='scr()' onmousewheel='scr()' onresize='scr()' class=bkg03>
<iframe name=post style='display:none'></iframe><!-- visible;height:200;width:100% -->
<form id=CandSummaryForm method=post target=post action='client-campaign-post.asp' onsubmit='submitFilters()' style='margin:0'>

<!-- HEADER LEVEL FILTERS -->
<table id=Bills width=100% border=0 cellspacing=0 cellpadding=0 style='border-bottom-width:0px'>
<tr align=center class=hdg29><td>Election Year</td><td>Races & Committees</td><td>Status</td><td>Party</td></tr>
<tr align=center class=hdg29>
<td><select name=FilterYear onchange='updateFilter()'>
  <option value=0<%=YearSel(0)%>>All Future
  <option value=1<%=YearSel(1)%>><%=intYear-2%>
  <option value=2<%=YearSel(2)%>><%=intYear-1%>
  <option value=3<%=YearSel(3)%>><%=intYear%>
  <option value=4<%=YearSel(4)%>><%=intYear+1%>
  <option value=5<%=YearSel(5)%>><%=intYear+2%>
  <option value=6<%=YearSel(6)%>><%=intYear+3%>
</select></td>
<td><select name=FilterRace onchange='updateFilter(1)'>
<optgroup label=General>
<option value=0<%=RaceSel(0)%>>All
<option value=50<%=RaceSel(50)%>>PAC Organizations
</optgroup>
<%
	Response.Write "<optgroup label='State-Wide Races'>"
	For i = 1 to intNumSWRaces
		Response.Write "<option value=" & 50+i & RaceSel(50+i) & ">" & aSWRaces(1,i-1)
	Next ' i
	Response.Write "</optgroup><optgroup label='Committees'>"
	CommitteeID = 0
	For i = 1 to intNumComms
		If intFilterRace = 50+intNumSWRaces+i Then CommitteeID = aComms(0,i-1)
		Response.Write _
			"<option id =" & aComms(0,i-1) & " value=" & 50+intNumSWRaces+i & RaceSel(50+intNumSWRaces+i) & ">" & aComms(1,i-1) & " (" & aComms(2,i-1) & ")"
	Next ' i
	''      Response.Write "</optgroup><optgroup label='Districts'>"
	''      For i = 1 to 49
	''      	Response.Write "<option value=" & i & RaceSel(i) &">" & strOrdinal(i)
	''      Next 'i
	Response.Write "</optgroup>"
%>
</select></td>
<td><select name=FilterStatus onchange='updateFilter()'>
  <option value=0<%=StatusSel(0)%>>All Candidates
  <option value=1<%=StatusSel(1)%>>Incumbents
  <option value=2<%=StatusSel(2)%>>Senators
  <option value=3<%=StatusSel(3)%>>Representatives
  <option value=4<%=StatusSel(4)%>>Challengers
</select></td>
<td><select name=FilterParty onchange='updateFilter()'>
  <option value=0<%=PartySel(0)%>>All
  <option value=1<%=PartySel(1)%>>Democrat
  <option value=2<%=PartySel(2)%>>Republican
  <option value=3<%=PartySel(3)%>>Third Party
</select></td>
</tr>
</table>


<!-- CANDIDATE DETAILS -->
<table width=100% cellspacing=0 cellpadding=0 class=det00 style='cursor:default;border-top-width:0px'>
<col align=center><col><col span=4 align=center><col width=90%>
<tr class=hdg29>
<td id=Sel style='cursor:pointer' onclick='selectCands(this)' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)' title="Select All">Sel</td>
<td style='cursor:pointer' onclick='sortBy("Cnd")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'><div style='width:175'>Candidate</div></td>
<td style='cursor:pointer' onclick='sortBy("Dst")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>District</td>
<td style='cursor:pointer' onclick='sortBy("Grp")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Group</td>
<td style='cursor:pointer' onclick='sortBy("Rec")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Rec'd</td>
<td style='cursor:pointer' onclick='sortBy("Act")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Actual</td>
<td>Comments</td>
</tr>

<!-- DETAIL LEVEL FILTERS -->
<tr class=bkg09>
<td id=checkmark class=lnk70 align=center valign=center onclick='selectMult()'><font face=Wingdings>&#252;</font></td>
<td><div class=box00 onclick='allDet(this)'></div>
<select name=FilterName style='width:150' onchange='updateFilter()'>
  <option value=0<%=NameSel(0)%>>All
  <option value=10<%=NameSel(10)%>>All PACs
  <option value=11<%=NameSel(11)%>>Partisan PACs
  <option value=1<%=NameSel(1)%>>A-B
  <option value=2<%=NameSel(2)%>>C-D
  <option value=3<%=NameSel(3)%>>E-G
  <option value=4<%=NameSel(4)%>>H-J
  <option value=5<%=NameSel(5)%>>K-L
  <option value=6<%=NameSel(6)%>>M-O
  <option value=7<%=NameSel(7)%>>P-R
  <option value=8<%=NameSel(8)%>>S-T
  <option value=9<%=NameSel(9)%>>U-Z
</select></td>


<td><select name=FilterDst onchange='updateFilter()'>
<option value=0<%=DstSel(0)%>>All
<%
	For i = 1 to 49 'intNumDistricts
		Response.Write "<option value=" & i & DstSel(i) & ">" & i
	Next ' i
%>
</select></td>




<td><select name=FilterGrp onchange='updateFilter()'>
<option value=0<%=GrpSel(0)%>>All
<%
	For i = 1 to intNumGroups
		Response.Write "<option value=" & i & GrpSel(i) & ">" & aGroups(0,i)
	Next ' i
%>
</select></td>
<td><select name=FilterRec onchange='updateFilter()'>
<option value=0<%=RecSel(0)%>>All
<option value=6<%=RecSel(6)%>>= 0
<option value=1<%=RecSel(1)%>>> 0
<option value=2<%=RecSel(2)%>>< 800
<option value=3<%=RecSel(3)%>>< 1600
<option value=4<%=RecSel(4)%>>= 1600
<option value=5<%=RecSel(5)%>>> 1600
</select></td>
<td><select name=FilterAct onchange='updateFilter()'>
<option value=0<%=ActSel(0)%>>All
<option value=6<%=ActSel(6)%>>= 0
<option value=1<%=ActSel(1)%>>> 0
<option value=2<%=ActSel(2)%>>< 800
<option value=3<%=ActSel(3)%>>< 1600
<option value=4<%=ActSel(4)%>>= 1600
<option value=5<%=ActSel(5)%>>> 1600
</select></td>
<td class=shd29>Change the selections to filter the list of candidates.</td>
</tr>
<%
' LOAD CLIENT CANDIDATE INFORMATION

	' Link together the tables which contain filter fields
	strSQL = _
		"(SELECT CandidateID, SUM(Amount) AS Actual FROM Contributions " & _
		"GROUP BY CandidateID, ClientID HAVING Contributions.ClientID=" & ClientID & ") AC"
	strJoin = "(Candidates C LEFT JOIN " & strSQL & " ON C.CandidateID=AC.CandidateID)"
	strJoin = "(" & strJoin & " INNER JOIN [State-Wide Races] SW ON C.SWRaceID=SW.SWRaceID)"
	strSQL = _
		"(SELECT PoliticianID, [Group], (ISNULL([Primary Rec],0)+ISNULL([General Rec],0)) TotRec," & _
		" ISNULL([Primary Rec],'') PriRec, ISNULL([General Rec],'') GenRec, Comments " & _
		"FROM [Client Politician Comments] WHERE ClientID=" & ClientID & ") AS CC"
	strJoin = "(" & strJoin & " LEFT JOIN " & strSQL & " ON C.PoliticianID=CC.PoliticianID) "

	' If filtered by committee, then link the candidates to the members of the desired committee
	If CommitteeID <> 0 Then
		strSQL = "(SELECT * FROM [Committee Membership] WHERE Year=DATEPART(YYYY,GETDATE()) AND CommitteeID=" & CommitteeID & ") CM"
		strJoin = "(" & strJoin & " INNER JOIN " & strSQL & " ON C.LegislatorID=CM.LegislatorID) "

	' If not filtered by committee, load an array with each candidate's committee membership information
	Else 'If intFilterRace >=1 And intFilterRace <= 49 Then
		strSQL = _
			"(SELECT CM.LegislatorID, CM.CommitteeID, CM.Position, CO.[Committee Abbr] " & _
			"FROM [Committee Membership] CM INNER JOIN [Committees] CO ON CM.CommitteeID=CO.CommitteeID " & _
			"WHERE CM.Year=DATEPART(YYYY,GETDATE())) CM"
		strSQL = _
			"SELECT C.CandidateID, CM.CommitteeID, CM.Position, CM.[Committee Abbr] " & _
			"FROM " & strJoin & " INNER JOIN " & strSQL & " ON C.LegislatorID=CM.LegislatorID " & _
			"WHERE " & strWhere & " ORDER BY " & strOrder & ",C.Withdrawn,CM.[Committee Abbr]"


'response.Write strsql & "<br><br>"


		Set rsMem=Server.CreateObject("ADOR.Recordset")
		rsMem.Open strSQL, strConnReadOnly
		If Not rsMem.EOF Then
			aMem = rsMem.GetRows()
			intNumMems = UBound(aMem,2)
		Else
			intNumMems = -1
		End If
		rsMem.Close
		Set rsMem = Nothing
	End If

	' Link in the candidates' actual contribution details
	strSQL = _	
		"(SELECT CandidateID, [Date], ISNULL(Amount,0) Amount, [Primary] " & _
		"FROM Contributions WHERE ClientID=" & ClientID & ") AD"
	strJoin = "(" & strJoin & " LEFT JOIN " & strSQL & " ON C.CandidateID=AD.CandidateID) "

	' Load list of candidates
	strSQL = _
		"SELECT C.CandidateID, C.PoliticianID PolID, C.[Year], C.Withdrawn," & _
		" C.Party, ISNULL(C.DistrictID,0) DistrictID, C.Seat, SW.Race, SW.Abbr," & _
		" ISNULL(C.[FirstName],'') FirstName, C.[LastName], C.[TaxpayerID]," & _
		" C.[Campaign Name], C.[Campaign URL], C.[Campaign Email]," & _
		" C.[CampaignStreet], C.[CampaignStreet2], C.[CampaignStreet3]," & _
		" C.[CampaignCity], C.[CampaignState], C.[CampaignPostalCode]," & _
		" C.[CampaignPhone], C.[CampaignFax]," & _
		" AD.[Date], AD.Amount, AD.[Primary]," & _
		" AC.Actual, CC.*," & _
		" C.PDCcontribs, C.PDCCdate, C.PDCexpense, C.PDCEdate " & _
		"FROM " & strJoin & _
		"WHERE " & strWhere & " ORDER BY " & strOrder & ",C.Withdrawn,AD.[Primary] DESC,AD.[Date]"	


'response.Write strsql


'response.End

	Set rsCand=Server.CreateObject("ADOR.Recordset")
	rsCand.Open strSQL, strConnReadOnly, adOpenDynamic, adLockPessimistic

	i = 0
	intMem = 0
	prevSeat = -1
	Do Until rsCand.EOF
		intCandID = rsCand("CandidateID")
		If rsCand("Withdrawn") Then
			strWD = " style='color:#808080'"
		Else
			strWD = ""
		End If

' CANDIDATE NAME, PARTY, AND RACE INFORMATION
		If rsCand("FirstName") = "" Then
			strCandidate = _
				rsCand("LastName") & " (" & _
				rsCand("Party") & ")"
		ElseIf intFilterRace < 51+intNumSWRaces And intFilterRace <> 0 Then
			strCandidate = _
				rsCand("LastName") & ", " & _
				rsCand("FirstName") & " (" & _
				rsCand("Party") & ")"
		ElseIf rsCand("Abbr") <> "" Then
			strCandidate = _
				rsCand("LastName") & ", " & _
				rsCand("FirstName") & " (" & _
				rsCand("Party") & "-" & _
				rsCand("Abbr") & ")"
		Else
			strCandidate = _
				rsCand("LastName") & ", " & _
				rsCand("FirstName") & " (" & _
				rsCand("Party") & ")"           '' "-" & _
				                                '' rsCand("DistrictID") & ")"
		End If

' CAMPAIGN PHONE AND FAX NUMBERS
		strPhone = rsCand("CampaignPhone")
		If Len(strPhone) < 7 Then
			strPhone = ""
		ElseIf Len(strPhone) < 10 Then
			strPhone = Left(strPhone,3) & "-" & Right(strPhone,4)
		Else
			strPhone = _
				"(" & Left(strPhone,3) & ") " & _
				Mid(strPhone,4,3) & "-" & _
				Right(strPhone,4)
		End If
		strFax = rsCand("CampaignFax")
		If IsNull(rsCand("CampaignFax")) Or Len(strFax) < 7 Then
			strFax = ""
		ElseIf Len(strFax) < 10 Then
			strFax = Left(strFax,3) & "-" & Right(strFax,4)
		Else
			strFax = _
				"(" & Left(strFax,3) & ") " & _
				Mid(strFax,4,3) & "-" & _
				Right(strFax,4)
		End If

		strDetails = ""
' LINKS TO CANDIDATES' COMMITTEES

		If CommitteeID = 0 Then
			If intMem <= intNumMems Then
				Do Until aMem(0,intMem) <> intCandID
					strMemPos = ""
					If aMem(2,intMem) <> "" Then strMemPos = " (" & aMem(2,intMem) & ")"
					strDetails = strDetails & _
						"<span class=lnk70 onclick='gotoRace(2," & aMem(1,intMem) & ")'>" & _
						aMem(3,intMem) & strMemPos & "</span>, "
					intMem = intMem + 1
					If intMem > intNumMems Then Exit Do
				Loop
			End If
			k = Len(strDetails)
			If k <> 0 Then strDetails = Left(strDetails,k-2) & "<br>"
		End If

' LINKS TO CANDIDATES' DISTRICT/STATE-WIDE RACES
		''      If intFilterRace = 0 Or intFilterRace > 50+intNumSWRaces Then
		If intFilterDst = 0 Then
			If rsCand("DistrictID") <> 0 And rsCand("DistrictID") <> "" Then
				strDetails = _
					"<span class=lnk70 onclick='gotoDistrict(" & rsCand("DistrictID") & ")'>" & _
					strOrdinal(rsCand("DistrictID")) & " District</span><br>" & _
					strDetails
			Else
				strDetails = _
					"<span class=lnk70 onclick='gotoRace(1,""" & rsCand("Race") & """)'>" & _
					rsCand("Race") & "</span><br>" & _
					strDetails
			End If
		End If

' CANDIDATE DETAIL INFORMATION
		If Len(rsCand("Campaign Name")) Then strDetails = strDetails & rsCand("Campaign Name") & "<br>"
		strDetails = strDetails & "Taxpayer ID: " & rsCand("TaxpayerID") & "<br>"
		If Len(rsCand("Campaign URL")) <> 0 Then strDetails = strDetails & _
			"<a target=campaign href='http://" & rsCand("Campaign URL") & "'>" & rsCand("Campaign URL") & "</a><br>"
		If Len(rsCand("Campaign Email")) <> 0 Then strDetails = strDetails & _
			"<a href='mailto:" & rsCand("Campaign Email") & "'>" & rsCand("Campaign Email") & "</a><br>"
		If Len(rsCand("CampaignStreet")) <> 0 Then strDetails = strDetails & rsCand("CampaignStreet") & "<br>"
		If Len(rsCand("CampaignStreet2")) <> 0 Then strDetails = strDetails & rsCand("CampaignStreet2") & "<br>"
		If Len(rsCand("CampaignStreet3")) <> 0 Then strDetails = strDetails & rsCand("CampaignStreet3") & "<br>"
		strDetails = strDetails & _
			rsCand("CampaignCity") & ", " & _
			rsCand("CampaignState") & " " & _
			rsCand("CampaignPostalCode") & "<br>"
		If Len(strPhone) <> 0 Then strDetails = strDetails & strPhone
		If Len(strFax) <> 0 Then strDetails = strDetails & " / " & strFax & " fax"

        If IsDate(rsCand("PDCCdate")) Then
            strDetails = strDetails & "<br>PDC Contributions: " & FormatCurrency(rsCand("PDCcontribs")) & " as of " & rsCand("PDCCdate")
        End If
        If IsDate(rsCand("PDCEdate")) Then
            strDetails = strDetails & "<br>PDC Expenses: " & FormatCurrency(rsCand("PDCexpense")) & " as of " & rsCand("PDCEdate")
        End If

		If rsCand("DistrictID") = 0 Then
			strDistrict = "&nbsp;"
		Else
			strDistrict = rsCand("DistrictID")
		End If

		If IsNull(rsCand("Group")) Then
			strGroup = "&nbsp;"
		Else
			strGroup = rsCand("Group")
		End If

' CANDIDATE CONTRIBUTION TOTALS
		If IsNull(rsCand("TotRec")) Or rsCand("TotRec") = 0 Or rsCand("Withdrawn") Then
			strRecommend = "&nbsp;"
		Else
			strRecommend = rsCand("TotRec")
		End If
		'intRecTot = intRecTot + rsCand("TotRec")
		If IsNull(rsCand("Actual")) Or rsCand("Actual") = 0 Or rsCand("Withdrawn") Then
			strActual = "&nbsp;"
		Else
			strActual = rsCand("Actual")
		End If
		'intActTot = intActTot + rsCand("Actual")

' DISTRICT SEAT HEADERS
		If intFilterDst <> 0 Then
			intSeat = rsCand("Seat")
			If  intSeat <> prevSeat Then
				Response.Write "<tr class=hdg23><td align=left colspan=6>"
				If intSeat = 0 Then Response.Write "Senate Candidates"
				If intSeat = 1 Then Response.Write "House Position 1 Candidates"
				If intSeat = 2 Then Response.Write "House Position 2 Candidates"
				Response.Write "</td></tr>"
				prevSeat = rsCand("Seat")
			End If
		End If

' DISPLAY CANDIDATE INFORMATION
		Response.Write _
			"<tr id=s" & i & " valign=top class=bkg04>" & _
			"<td><input id=chk" & i & " type=checkbox></td>" & _
			"<td><div class=box00 onclick='toggleDet(" & i & ")'></div>" & _
			"<span id=nam" & i & " class=lnk70 onclick='selectDetail(" & _
				i & "," & intCandID & "," & rsCand("PolID") & ")'" & strWD & ">" & strCandidate & "</span></td>" & _
			"<td id=dst" & i & ">" & strDistrict & "</td>" & _
			"<td id=grp" & i & ">" & strGroup & "</td>" & _
			"<td id=rec" & i & ">" & strRecommend & "</td>" & _
			"<td id=act" & i & ">" & strActual & "</td>" & _
			"<td><div id=com" & i & " class=c0>" & _
			MakeHTML(rsCand("Comments")) & "</div></td></tr>"
		Response.Write _
			"<tr id=d" & i & " valign=top class=bkg04 style='display:none'><td></td>" & _
			"<td colspan=5 style='margin-left:10'>" & strDetails

' CANDIDATE CONTRIBUTION DETAILS
		strMisc = rsCand("PriRec") & "," & rsCand("GenRec") & "," & rsCand("Year")
		strAmt = ""
		For j = 1 to 3 ' Up to three primary election contributions
			If Not rsCand.EOF Then
				If rsCand("CandidateID") = intCandID And rsCand("Primary") Then
					strAmt = strAmt & rsCand("Date") & "," & rsCand("Amount") & ","
					rsCand.MoveNext
				Else
					strAmt = strAmt & ",,"
				End If
			Else
				strAmt = strAmt & ",,"
			End If
		Next 'j
		For j = 4 to 6 ' Up to three general election contributions
			If Not rsCand.EOF Then
				If rsCand("CandidateID") = intCandID Then
					strAmt = strAmt & rsCand("Date") & "," & rsCand("Amount") & ","
					rsCand.MoveNext
				Else
					strAmt = strAmt & ",,"
				End If
			Else
				strAmt = strAmt & ",,"
			End If
		Next
		Response.Write _
			"<span id=amt" & i & " style='display:none'>" & strAmt & strMisc & "</span>"

		Response.Write "</td></tr>"
		i = i + 1
	Loop
	rsCand.Close
	Set rsCand = Nothing
%>
</table>
<input type=hidden name=CandCount value=<%=i%> ID="Hidden1">

<!-- MULTIPLE BILL UPDATE BOX -->
<input type=hidden name=UpdateMult value=True ID="Hidden2">
<input type=hidden name=CandsToUpdate ID="Hidden3">
<div id=MultDetails class=div1A style="z-index:2;display:none;position:absolute;left:15;padding:5 0;height:140;width:95%;overflow:hidden">
<p><table width=100% border=0 cellspacing=0 cellpadding=0 class=hdg1A style='padding:0 2' ID="Table2">
<col align=right width=1><col span=6 align=center><col width=90%>
<tr><td></td><td colspan=2>Primary</td><td>&nbsp;</td><td colspan=2>General</td><td>Group</td><td></td></tr>
<tr><td>Recommend:</td><td colspan=2><input name=PriRec size=4 ID="Text1"></td><td></td><td colspan=2><input name=GenRec size=4 ID="Text2"></td><td> <input name=Group size=3 ID="Text3"></td></tr>
<tr class=shd1A><td></td><td><u>Date</u></td><td><u>Amount</u></td><td></td><td><u>Date</u></td><td><u>Amount</u></td></tr>
<tr><td>Actual:</td><td>
<input name=PriDate size=7 onchange='return isDate(this)' ID="Text4"></td><td><input name=PriAmt size=4 ID="Text5">
</td><td></td><td>
<input name=GenDate size=7 onchange='return isDate(this)' ID="Text6"></td><td><input name=GenAmt size=4 ID="Text7">
</td></tr>
<tr valign=bottom><td colspan=7 align=center style='height:35'>
<input type=button class=btn61 value=Submit onclick='submitMult()' ID="Button1" NAME="Button1"><span style='width:100'></span>
<input type=button class=btn61 value=Cancel onclick='hideDetail(arguments[0],1)' ID="Button2" NAME="Button2">
</td></tr>
</table></p>
</div>
</form>
<div class=hdg14 style='position:relative;height:100%;margin:0 4;padding:20'>
<%
	If i = 0 Then Response.Write "No candidates match the current filter settings."
%>
</div>

<!-- CANDIDATE DETAIL BOX -->
<div id=CandDetails class=div1A style="z-index:2;display:none;position:absolute;left:7;padding:5;height:205;width:97%;overflow:hidden">
<form id=CandDetailForm method=post target=post action='client-campaign-post.asp' onsubmit='submitDetail()'>
<input type=hidden name=UpdateCand value=True ID="Hidden4">
<input type=hidden name=CandID ID="Hidden5">
<input type=hidden name=PolID ID="Hidden6">
<input type=hidden name=Index ID="Hidden7">
<p><table width=100% border=0 cellspacing=0 cellpadding=0 class=hdg1A style='padding:0 5' ID="Table3">
<col align=right width=1><col span=2 align=center width=125>
<tr><td align=left colspan=4><input name=CandName readonly class=hdg1A style='border:0px;width:250' ID="Text8"></td><td align=right>Election Year: <input name=ElecYear readonly class=hdg1A style='width:35;border:0px' ID="Text9"></td></tr>
<tr><td></td><td>Primary</td><td>General</td><td>Comments</td><td align=right>Group: <input size=3 name=Group ID="Text10"></td></tr>
<tr valign=top><td>Rec'd:</td><td><input name=PriRec size=4 ID="Text11"></td><td><input name=GenRec size=4 ID="Text12"></td><td colspan=2 rowspan=3><textarea name=Com rows=7 style='width:100%' ID="Textarea1"></textarea></td></tr>
<tr class=shd1A><td></td><td align=right><u>Date</u> &nbsp; &nbsp; &nbsp; <u>Amount</u>&nbsp;</td><td align=right><u>Date</u> &nbsp; &nbsp; &nbsp; <u>Amount</u>&nbsp;</td></tr>
<tr valign=top><td style='padding-top:3'>Actual:</td><td>
<input size=7 name=C onchange='return isDate(this)' ID="Text13"> <input size=4 name=C ID="Text14"><br>
<input size=7 name=C onchange='return isDate(this)' ID="Text15"> <input size=4 name=C ID="Text16"><br>
<input size=7 name=C onchange='return isDate(this)' ID="Text17"> <input size=4 name=C ID="Text18"><br>
</td><td>
<input size=7 name=C onchange='return isDate(this)' ID="Text19"> <input size=4 name=C ID="Text20"><br>
<input size=7 name=C onchange='return isDate(this)' ID="Text21"> <input size=4 name=C ID="Text22"><br>
<input size=7 name=C onchange='return isDate(this)' ID="Text23"> <input size=4 name=C ID="Text24"><br>
</td></tr>
<tr valign=bottom><td colspan=5 align=center style='height:35'>
<input type=submit class=btn61 value=Submit ID="Submit1" NAME="Submit1">
<input type=checkbox name=ApplyToAll value=All ID="Checkbox1">Apply comments to all tracking lists
&nbsp; <input type=button class=btn61 value=Cancel onclick='hideDetail(arguments[0],1)' ID="Button3" NAME="Button3">
</td></tr>
</table></p>
</form>
</div>

</body>
</html>