<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"
	Set rst=Server.CreateObject("ADOR.Recordset")
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection

	EventID = Request.Form("EventID")
	If IsNumeric(EventID) Then
		EventID = CLng(EventID)
		strSearch = Request.Form("Search")
	End If

' ADD/UPDATE EVENT
	If Request.Form("Update") = "True" Then
		If EventID = -1 Then bolInsert = True
		EvtDate = FormatDateTime(Request.Form("Date"),2)
		If  bolInsert Then
			strCommand = "INSERT INTO [Campaign Events] ([Date]) VALUES ('" & EvtDate & "')"
			cxnSQL.Execute strCommand, , adExecuteNoRecords
			strCommand = "SELECT MAX(EventID) AS MaxID FROM [Campaign Events]"
			Set rsResult = cxnSQL.Execute(strCommand)
			EventID = rsResult("MaxID")
			Set rsResult = Nothing
		End If
		intLen = CInt("0" & Request.Form("Length"))
		If intLen = 0 Then intLen = 1
		strCommand = _
			"UPDATE [Campaign Events] SET" & _
			" [Date]='" & EvtDate & "'," & _
			" [Time]='" & Request.Form("Time") & "'," & _
			" [Length]=" & intLen & "," & _
			" [Title]='" & TweakQuote(Request.Form("Title")) & "'," & _
			" [Address Line 1]='" & TweakQuote(Request.Form("Addr1")) & "'," & _
			" [Address Line 2]='" & TweakQuote(Request.Form("Addr2")) & "'," & _
			" [City]='" & TweakQuote(Request.Form("City")) & "'," & _
			" [State]='" & Request.Form("State") & "'," & _
			" [Zip]='" & Request.Form("Zip") & "'," & _
			" [Comments]='" & TweakQuote(Request.Form("Comments")) & "' " & _
			" WHERE EventID=" & EventID
		cxnSQL.Execute strCommand, , adExecuteNoRecords

		' Add Event Candidates
		If Request.Form("CandID") <> "" Then
			strSearch = ""
			aCandID = Split(Request.Form("CandID"),",")
			For i = 0 to UBound(aCandID)
				strCommand = _
					"INSERT INTO [Campaign Event Candidates] VALUES (" & _
					EventID & "," & aCandID(i) & ")"
				cxnSQL.Execute strCommand, , adExecuteNoRecords
			Next 'i
		End If
	End If

' DELETE EVENT
	If Request.Form("Update") = "Delete" Then
		strCommand = "DELETE FROM [Campaign Events] WHERE EventID=" & EventID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		EventID = 0
	End If

' DELETE EVENT CANDIDATE
	If Request.Form("Update") = "DeleteCand" Then
		strCommand = _
			"DELETE FROM [Campaign Event Candidates] " & _
			"WHERE EventID=" & EventID & _
			" AND CandidateID=" & Request.Form("EvtCand")
		cxnSQL.Execute strCommand, , adExecuteNoRecords
	End If

	cxnSQL.Close
	Set cxnSQL = Nothing

' LOAD ALL EVENTS
	intEvtCount = -1
	Index = 0
	If EventID <> -1 Then
		strSQL = _
			"SELECT EventID, [Date], [Time], Length, Title, [Address Line 1], [Address Line 2]," & _
			" City, State, Zip, Comments " & _
			"FROM [Campaign Events] WHERE [Date] >= '" & Date & "' ORDER BY [Date], Title"
		rst.Open strSQL, strConnection
		If Not rst.EOF Then
			aEvts = rst.GetRows()
			intEvtCount = UBound(aEvts,2)
			'If intEvtCount = 0 Then EventID = aEvts(0,0)
			For i = 0 to intEvtCount
				If aEvts(0,i) = EventID Then Index = i
			Next 'i
			If IsDate(aEvts(2,Index)) Then
				strTime = FormatDateTime(aEvts(2,Index), 3)
				strTime = Mid(strTime,1,Len(strTime)-6) & LCase(Right(strTime,3))
			Else
				strTime = ""
			End If
		End If
		rst.Close
	Else
		Dim aEvts(10,1)
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var df
function init(){
	df=document.getElementById("DetailForm")
	selectTab(4)
	if (<%=EventID%>!=0) df.Date.focus()
}
function EvtDetails(e){
	df.EventID.value=e
	df.Update.value="False"
	df.submit()
}
function Select(){
	df.EventID.value=0
	df.Update.value="False"
	df.submit()
}
function Cancel(){
	df.Update.value="False"
	df.EventID.disabled=true
	df.Search.value=""
	df.submit()
}
function Delete(){
	if(df.EventID.value!=0&&confirm("Click OK to confirm delete for this event.")){
		df.Update.value="Delete"
		df.submit()
	}
}
function DeleteCand(c){
	if(confirm("Click OK to confirm delete for this event candidate.")){
		df.Update.value="DeleteCand"
		df.EvtCand.value=c
		df.submit()
	}
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:15'>

<form id=DetailForm action="maint-events.asp" method=post>
<input type=hidden name=Update value=True>
<input type=hidden name=EventID value=<%=EventID%>>
<input type=hidden name=EvtCand>
<input type=button value="Add New Event" onclick='EvtDetails(-1)'>
<br>

<%
'
'	EVENTS SUMMARY
'
	If EventID = 0 Then
		Response.Write _
			"<br><span class=hdg24>Events Summary</span>" & _
			"<div class=box20 style='padding:10'>" & _
			"<table border=0 cellpadding=0 cellspacing=0 width=700 class=det00>" & _
			"<col width=80><col width=400><col width=220>" & _
			"<tr class=shd24><td><u>Date</u></td><td><u>Title</u></td><td><u>City</u></td></tr>"
		For i = 0 to intEvtCount
			Response.Write _
				"<tr><td>" & aEvts(1,i) & "</td>" & _
				"<td onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
				" style='cursor:hand' onclick='EvtDetails(" & aEvts(0,i) & ")'>" & _
				aEvts(4,i) & "</td><td>" & aEvts(7,i) & "</td></tr>"
		Next 'i
		Response.Write "</table></div>"
	End If

'
'	EVENT DETAILS
'
	If EventID <> 0 Then
%>
<br><span class=hdg24>Event Details</span>
<div class=box20 style='padding:10'>
<table border=0 cellpadding=0 cellspacing=0 width=98% class=shd24 style='padding-right:3'>
<col width=60 align=right><col width=85><col width=60 align=right><col width=250>
<tr><td>Date:</td><td><input name=Date tabindex=1 style='width:75' value="<%=aEvts(1,Index)%>" onchange='return isDate(this)'></td>
<td>Title:</td><td><input name=Title tabindex=4 style='width:240' value="<%=aEvts(4,Index)%>"></td>
<td><u>Comments</u></td></tr>
<tr><td>Time:</td><td><input name=Time tabindex=2 style='width:75' value="<%=strTime%>"></td>
<td>Address:</td><td><input name=Addr1 tabindex=5 style='width:240' value="<%=aEvts(5,Index)%>">
<td rowspan=4 valign=top><textarea name=Comments tabindex=10 style='width=100%;height:67'><%=aEvts(10,Index)%></textarea></td></tr>
<tr><td>Length:</td><td><input name=Length tabindex=3 style='width:75' value="<%=aEvts(3,Index)%>"></td>
<td></td><td><input name=Addr2 tabindex=6 style='width:240' value="<%=aEvts(6,Index)%>"></td></tr>
<tr><td colspan=3></td>
<td><input name=City tabindex=7 style='width:124' value="<%=aEvts(7,Index)%>">
<input name=State tabindex=8 style='width:30' value="<%=aEvts(8,Index)%>">
<input name=Zip tabindex=9 style='width:78' value="<%=aEvts(9,Index)%>"></td></tr>
<tr><td colspan=4>&nbsp;</td></tr>
</table></div>
<%
	End If
'
'	EVENT CANDIDATES
'
	If EventID <> 0 Then
		strSQL = _
			"SELECT C.CandidateID, C.Party, ISNULL(C.DistrictID,0) Dist," & _
			" CASE WHEN ISNULL(C.FirstName, '') = '' THEN C.LastName ELSE C.FirstName + ' ' + C.LastName END Name " & _
			"FROM [Campaign Event Candidates] E" & _
			" INNER JOIN Candidates C ON E.CandidateID=C.CandidateID " & _
			"WHERE E.EventID=" & EventID & " ORDER BY C.LastName"
		rst.Open strSQL, strConnection
		bolCands=False
		If Not rst.EOF Then
			bolCands=True
			Response.Write "<br><span class=hdg24>Sponsoring Candidate(s):</span><div class=box20 style='padding:10'>"
		End If
		Do Until rst.EOF
			Response.Write _
				"<div onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
				" class=det00 style='width:250;cursor:hand' onclick='DeleteCand(" & rst("CandidateID") & ")'>" & _
				rst("Name") & " (" & rst("Party")
			If rst("Dist") <> 0 Then Response.Write "-" & rst("Dist")
			Response.Write ")</div>"
			rst.MoveNext
		Loop
		rst.Close
		If bolCands Then Response.Write "</div>"
'
'	NEW CANDIDATE SELECTIONS
'
		If EventID > 0 Then
			Response.Write _
				"<br><span class=hdg24>Candidate Last Names Starting With:</span> &nbsp;" & _
				"<input name=Search width=100 value='" & strSearch & "'> &nbsp;" & _
				"<input type=submit value=Search onclick='DetailForm.Update.value=""False""'>" & _
				"<div class=shd24>(Enter '*' for all current year candidates.)</div>"
		End If
		If strSearch <> "" Then
			If strSearch = "*" Then strSearch = ""
			strSQL = _
				"SELECT CandidateID, Party, ISNULL(DistrictID,0) Dist," & _
				" CASE WHEN ISNULL(FirstName, '') = '' THEN LastName ELSE FirstName + ' ' + LastName END Name " & _
				"FROM Candidates " & _
 				"WHERE LastName LIKE '" & strSearch & "%'" & _
 				" AND [Year]=DATEPART(yyyy, GETDATE())" & _
 				" AND Withdrawn=0 " & _
 				"ORDER BY LastName"
			rst.Open strSQL, strConnection
			bolCands=False
			If Not rst.EOF Then
				bolCands=True
				Response.Write "<div class=det00 style='padding:10'>"
			End If
			Do Until rst.EOF
				Response.Write _
					"<input type=checkbox name=CandID value=" & rst("CandidateID") & "> " & _
					rst("Name") & " (" & rst("Party")
				If rst("Dist") <> 0 Then Response.Write "-" & rst("Dist")
				Response.Write ")<br>"
				rst.MoveNext
			Loop
			rst.Close
			If bolCands Then Response.Write "</div>"
		End If

	End If ' EventID <> 0
	Set rst = Nothing

	If EventID <> 0 Then
%>
<center><br>
<input type=submit value=Submit><span style='width:200'></span>
<input type=button onclick='Cancel()' value=Cancel><span style='width:200'></span>
<input type=button onclick='Delete()' value=Delete></td></tr>
</center>
<%
	End If
%>
</form>

</body>
</html>

