<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection

	CommID = CLng("0" & Request.Form("CommID"))
	MemID = CLng("0" & Request.Form("MemID"))

' ADD/UPDATE COMMITTEE
	If Request.Form("UpdateComm") = "True" Then
		If CommID = 9999999 Then
			strCommand = _
				"INSERT INTO [Committees] ([Committee Name]) VALUES (" & _
				"'" & TweakQuote(Request.Form("Comm")) & "')"
			cxnSQL.Execute strCommand, , adExecuteNoRecords
			strCommand = "SELECT MAX(CommitteeID) AS MaxID FROM [Committees]"
			Set rsResult = cxnSQL.Execute(strCommand)
			CommID = rsResult("MaxID")
			MemID = 9999999
			Set rsResult = Nothing
		End If

		strCommand = _
			"UPDATE [Committees] SET" & _
			" [Committee Name]='" & TweakQuote(Request.Form("Comm")) & "'," & _
			" [Committee Abbr]='" & TweakQuote(Request.Form("Abbr")) & "'," & _
			" [House]='" & TweakQuote(Request.Form("House")) & "'," & _
			" [Address1]='" & TweakQuote(Request.Form("Addr1")) & "'," & _
			" [Address2]='" & TweakQuote(Request.Form("Addr2")) & "'," & _
			" [City]='" & TweakQuote(Request.Form("City")) & "'," & _
			" [State]='" & TweakQuote(Request.Form("State")) & "'," & _
			" [Zip]='" & TweakQuote(Request.Form("Zip")) & "'," & _
			" [Telephone]='" & TweakQuote(Request.Form("Phone")) & "'," & _
			" [Fax]='" & TweakQuote(Request.Form("Fax")) & "'," & _
			" [LocLevelID]=" & Request.Form("LocLevel") & _
			" WHERE CommitteeID=" & CommID
		cxnSQL.Execute strCommand, , adExecuteNoRecords

		strCommand = "DELETE FROM [Committee Locations] WHERE CommitteeID=" &CommID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		For i = 0 to 5
			If Trim(Request.Form("CLB" & i)) <> "" AND Trim(Request.Form("CLE" & i)) <> "" Then
				strCommand = _
					"INSERT INTO [Committee Locations] VALUES (" & _
					CommID & "," & _
					"'" & Request.Form("CLB" & i) & "'," & _
					"'" & Request.Form("CLE" & i) & "')"
				cxnSQL.Execute strCommand, , adExecuteNoRecords
			End If
		Next 'i

	End If

' DELETE COMMITTEE
	If Request.Form("UpdateComm") = "Delete" Then
		strCommand = "DELETE FROM [Committees] WHERE CommitteeID=" & CommID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		strCommand = "DELETE [Committee Locations] FROM [Committee Locations] L LEFT JOIN Committees C ON L.CommitteeID=C.CommitteeID WHERE C.[Committee Name] IS NULL"
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		CommID = 0
		MemID = 0
	End If

' ADD/UPDATE COMMITTEE MEMBER
	If Request.Form("UpdateMem") = "True" Then
		If MemID <> 9999999 Then
			strCommand = _
				"UPDATE [Committee Membership] SET" & _
				" [Position]='" & TweakQuote(Request.Form("Pos")) & "' " & _
				"WHERE CommitteeID=" & CommID & _
				" AND LegislatorID=" & MemID & _
				" AND [Year]=DATEPART(yyyy,GETDATE())"
			cxnSQL.Execute strCommand, , adExecuteNoRecords
		Else
			For i = 0 to Request.Form("LegCount")
				If Len(Request.Form("Leg" & i)) <> 0 Then
					strCommand = "INSERT INTO [Committee Membership] VALUES (" & _
						CommID & "," & _
						Request.Form("Leg" & i) & "," & _
						DatePart("yyyy",Date) & ",NULL)"
					cxnSQL.Execute strCommand, , adExecuteNoRecords
				End If
			Next ' i
		End If
		MemID = 0
	End If

' DELETE COMMITTEE MEMBER
	If Request.Form("UpdateMem") = "Delete" Then
		strCommand = _
			"DELETE FROM [Committee Membership] " & _
			"WHERE [Year]=DATEPART(yyyy, GETDATE())" & _
			" AND CommitteeID=" & CommID & _
			" AND LegislatorID=" & MemID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		MemID = 0
	End If

' LOAD COMMITTEE MEMBER INFORMATION
	intCommMemCount = -1
	If CommID <> 0 And CommID <> 9999999 Then
		strSQL = 	_
			"SELECT C.LegislatorID, C.[Position]," & _
			" P.LastName + ', ' + P.FirstName + ' (' + L.Party + '-' + CAST(L.DistrictID AS varchar(2)) + ')' AS Name " & _
			"FROM [Committee Membership] C" & _
			" INNER JOIN [Legislator Details] L ON C.LegislatorID = L.LegislatorID " & _
			" INNER JOIN Politicians P ON L.PoliticianID = P.PoliticianID " & _
			"WHERE C.[Year]=DATEPART(yyyy, GETDATE()) AND C.CommitteeID=" & CommID & _
			" ORDER BY P.LastName, P.FirstName"
		Set rsCommMems=Server.CreateObject("ADOR.Recordset")
		rsCommMems.Open strSQL, strConnection
		If Not rsCommMems.EOF Then
			aCommMems = rsCommMems.GetRows()
			intCommMemCount = UBound(aCommMems,2)
			For i = 0 to intCommMemCount
				If aCommMems(0,i) = MemID Then CommMemIndex = i
			Next 'i
		End If
		rsCommMems.Close
		Set rsCommMems = Nothing
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var cdf,mdf
function init(){
	cdf=document.getElementById("CommDetailForm")
	mdf=document.getElementById("MemDetailForm")
	selectTab(2)
	
	 m=0<%=MemID%>
	if (m!=0&&m!=9999999) mdf.Pos.focus()
	else if (<%=CommID%>!=0) cdf.Comm.focus()
}
function commSelect(){
	cdf.UpdateComm.value="False"
	cdf.submit()
}
function commCancel(){
	cdf.UpdateComm.value="False"
	cdf.CommID.disabled=true
	cdf.submit()
}
function commDelete(){
	if (confirm("Click OK to confirm delete for this committee.")){
		cdf.UpdateComm.value="Delete"
		cdf.submit()
	}
}
function memSelect(c){
	if (c!=-1) mdf.MemID.value=c
	mdf.UpdateMem.value="False"
	mdf.submit()
}
function memCancel(){
	mdf.UpdateMem.value="False"
	mdf.MemID.disabled=true
	mdf.submit()
}
function memDelete(){
	if (confirm("Click OK to confirm delete for this committee member.")){
		mdf.UpdateMem.value="Delete"
		mdf.submit()
	}
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:15'>

<%
' UPDATE DAILY STATUS ABBRV LINKS
	strCommand = _
		"UPDATE [Daily Status] " & _
		"SET CommitteeID = C.CommitteeID " & _
		"FROM [Daily Status] D, [Committee Locations] C " & _
		"WHERE D.Location BETWEEN C.[Range Begin] AND C.[Range End]"
	cxnSQL.Execute strCommand, , adExecuteNoRecords

' DISPLAY ANY UNLINKED COMMITTEE ABBRS IN DAILY STATUS
	strSQL = 	"SELECT DISTINCT Location, House FROM [Daily Status] WHERE CommitteeID=94"
	Set rsUnlinked=Server.CreateObject("ADOR.Recordset")
	rsUnlinked.Open strSQL, strConnection
	IsUnlinked = 0
	If rsUnlinked.EOF = False Then
		Response.Write "<div class=hdg24>UNRECOGNIZED ABBRVS FROM DAILY STATUS</div><br>"
  	IsUnlinked = 1
	End If
	Do Until rsUnlinked.EOF
		Response.Write "<span class=det00><span style='width:150px'>" & rsUnlinked("Location") & "</span>"
		If rsUnlinked("House") = "H" Then
			Response.Write "House"
		Else
			Response.Write "Senate"
		End If
		Response.Write "</span><br>"
		rsUnlinked.MoveNext
	Loop
	If isUnlinked = 1 Then
		Response.Write "<div><hr><br></div>"
	End If
	rsUnlinked.Close
	Set rsUnlinked = Nothing
%>


<form id=CommDetailForm action="maint-comm.asp" method=post>
<input type=hidden name=UpdateComm value=True>
<span class=hdg24>Committee:</span>
<select name=CommID style='width:325' onchange='cdf.cdfSubmit.disabled=true'>
<option value=9999999>-Add New-
<%
	strSQL = 	"SELECT CommitteeID,[Committee Name],[House] FROM [Committees] ORDER BY [Committee Name]"
	Set rsComms=Server.CreateObject("ADOR.Recordset")
	rsComms.Open strSQL, strConnection
	Do Until rsComms.EOF
		Response.Write "<option value=" & rsComms("CommitteeID")
		If CommID=rsComms("CommitteeID") Then Response.Write " selected"
		Response.Write ">" & rsComms("Committee Name")
		If Len(rsComms("House")) <> 0 Then
			Response.Write " (" & Left(rsComms("House"),1) & ")"
		End If
		rsComms.MoveNext
	Loop
	rsComms.Close
	Set rsComms = Nothing
%>
</select>
&nbsp; <input type=button value="Select" onclick='commSelect()'>

<!--
'------------------------------------------------
' COMMITTEE DETAILS
'------------------------------------------------
-->
<%
	If CommID <> 0 Then
		If CommID <> 9999999 Then
			strSQL = _
				"SELECT CommitteeID, [Committee Name], [Committee Abbr], House," & _
				" Address1, Address2, City, State, Zip, Telephone, Fax, LocLevelID " & _
				"FROM [Committees] WHERE CommitteeID=" & CommID
			Set rsComm=Server.CreateObject("ADOR.Recordset")
			rsComm.Open strSQL, strConnection
			aComm = rsComm.GetRows()
			rsComm.Close
			Set rsComm = Nothing
		Else
			Dim aComm(11,1)
		End If
		If aComm(3,0) = "House" Then House = " selected"
		If aComm(3,0) = "Senate" Then Senate = " selected"
%>
<div class=box20 style='padding:7'>
<table border=0 cellpadding=0 cellspacing=0 width=800 class=shd24 style='padding-left:3'>
<col width=66 align=right><col width=200>
<col width=66 align=right><col width=200>
<col width=66 align=right><col width=200>

<tr><td>Name:</td><td><input name=Comm type=text tabindex=1 style='width:190' value="<%=aComm(1,0)%>"></td>
<td>Address:</td><td><input name=Addr1 type=text tabindex=4 style='width:190' value="<%=aComm(4,0)%>"></td>
<td>Phone:</td><td><input name=Phone type=text tabindex=9 style='width:190' value="<%=aComm(9,0)%>"></td></tr>

<tr><td>Abbr:</td><td><input name=Abbr type=text tabindex=2 style='width:190' value="<%=aComm(2,0)%>"></td>
<td></td><td><input name=Addr2 type=text tabindex=5 style='width:190' value="<%=aComm(5,0)%>"></td>
<td>Fax:</td><td><input name=Fax type=text tabindex=10 style='width:190' value="<%=aComm(10,0)%>"></td></tr>

<tr><td>Chamber:</td><td><select name=House tabindex=3 style='width:190'>
<option value="" selected><option value=House <%=House%>>House<option value=Senate<%=Senate%>>Senate</select></td>
<td></td><td><input name=City type=text tabindex=6 style='width:74' value="<%=aComm(6,0)%>">
<input name=State type=text tabindex=7 style='width:30' value="<%=aComm(7,0)%>">
<input name=Zip type=text tabindex=8 style='width:78' value="<%=aComm(8,0)%>"></td>
<td>Location:</td><td><select name=LocLevel tabindex=11 style='width:190'>
<%
		strSQL = 	"SELECT LocLevelID, Level FROM [Location Levels]"
		Set rsLoc=Server.CreateObject("ADOR.Recordset")
		rsLoc.Open strSQL, strConnection
		aLoc = rsLoc.GetRows()
		rsLoc.Close
		Set rsLoc = Nothing
		For i = 0 to UBound(aLoc,2)
			Response.Write "<option value=" & aLoc(0,i)
			If aLoc(0,i) = aComm(11,0) Then Response.Write " selected"
			Response.Write ">" & aLoc(1,i)
		Next 'i
%>
</select></td></tr>
<tr><td colspan=6 align=left>
<br><b>Daily Status Abbreviations</b><br><br>
<span style='width=133px'>Range Begin</span>Range End<br>
<%
		strSQL = "SELECT * FROM [Committee Locations] WHERE CommitteeID=" & CommID
		Set rsLoc=Server.CreateObject("ADOR.Recordset")
		rsLoc.Open strSQL, strConnection
		i = 0
		Do Until rsLoc.EOF
			Response.Write _
				"<input name=CLB" & i & " type=text tabindex=" & 12+(i*2) & _
				" style='width:120' value='" & rsLoc("Range Begin") & "'> &nbsp; "
			Response.Write _
				"<input name=CLE" & i & " type=text tabindex=" & 12+(i*2) & _
				" style='width:120' value='" & rsLoc("Range End") & "'><br>"
			i = i + 1
			rsLoc.MoveNext
		Loop
		For j = i to 5
			Response.Write "<input name=CLB" & j & " type=text tabindex=" & 12+(j*2) & " style='width:120'> &nbsp; "
			Response.Write "<input name=CLE" & j & " type=text tabindex=" & 12+(j*2) & " style='width:120'><br>"
		Next 'j
		rsLoc.Close
		Set rsLoc = Nothing
%>
<br><td></tr>

<tr style='height:25'><td colspan=6 align=center valign=bottom>
<input id=cdfSubmit type=submit value=Submit><span style='width:200'></span>
<input type=button onclick='commCancel()' value=Cancel><span style='width:200'></span>
<input type=button onclick='commDelete()' value=Delete>
</td></tr>

</table>

</div>
<%
	End If
%>
</form>

<!--
'------------------------------------------------
' COMMITTEE ACCOUNTS SUMMARY
'------------------------------------------------
-->
<%
	If 	intCommMemCount <> -1 Then
		Response.Write _
			"<span class=hdg24>Committee Memebership Summary</span>" & _
			"<div class=box20 style='padding:10'>" & _
			"<table border=0 cellpadding=0 cellspacing=0 width=700 class=det00 align=center>" & _
			"<col width=300><col width=100><col width=300>"
		Response.Write _
			"<tr class=shd24>" & _
			"<td><u>Member Name</u></td><td></td>" & _
			"<td><u>Member Name</u></td></tr>" & _
			"<tr valign=top><td>"
		b=0
		e=Int(intCommMemCount/2)
		for i = 0 to 1
			For j = b to e
				Response.Write _
					"<div onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
					" style='cursor:hand' onclick='memSelect(" & aCommMems(0,j) & ")'>" & _
					aCommMems(2,j)
				If Len(aCommMems(1,j)) <> 0 Then
					Response.Write ", " & aCommMems(1,j)
				End If
				Response.Write "</div>"
			Next 'j
			b=e+1
			e=intCommMemCount
			Response.Write "</td><td></td><td>"
		Next 'i
		Response.Write "</td></tr></table></div>"
	End If
%>

<!--
'------------------------------------------------
' MEMBER DETAILS
'------------------------------------------------
-->
<form id=MemDetailForm action="maint-comm.asp" method=post>
<input type=hidden name=CommID value=<%=CommID%>>
<input type=hidden name=UpdateMem value=True>
<%
	If CommID <> 0 and CommID <> 9999999 Then
		Response.Write _
			"<span class=hdg24>Member Details:</span> " & _
			"<select name=MemID style='width:250'>" & _
			"<option value=9999999>-Add New-"

		If 	intCommMemCount <> -1 Then
			For i = 0 to intCommMemCount
				Response.Write "<option value=" & aCommMems(0,i)
				If MemID = aCommMems(0,i) Then Response.Write " selected"
				Response.Write ">" & aCommMems(2,i)
			Next 'i
		End If

		Response.Write _
			"</select>" & _
			"&nbsp; <input type=button value='Select' onclick='memSelect(-1)'>"
	End If

	If MemID <> 0 Then
		If MemID <> 9999999 Then
			Response.Write _
				"<div class=box20 style='padding:7'>" & _
				"<table border=0 cellpadding=0 cellspacing=0 width=700 class=shd24 align=center>"
			Response.Write _
				"<tr><td width=57>Position:</td>" & _
				"<td width=643><input name=Pos type=text style='width:150'" & _
				" value='" & aCommMems(1,CommMemIndex) & "'></td></tr>"

			Response.Write _
				"<tr style='height:25'><td colspan=2 align=center valign=bottom>" & _
				"<input type=submit value=Submit><span style='width:200'></span>" & _
				"<input type=button onclick='memCancel()' value=Cancel><span style='width:200'></span>" & _
				"<input type=button onclick='memDelete()' value=Delete>" & _
				"</td></tr>"
			Response.Write _
				"</table></div>"
		Else
			If aComm(3,0)="Senate" Then
				SQLWhere = "AND L.Seat=0 "
			ElseIf aComm(3,0)="House" Then
				SQLWhere = "AND L.Seat<>0 "
			Else
				SQLWhere = ""
			End If

			strSQL = _
				"(SELECT * FROM [Committee Membership] " & _
				"WHERE [Year]=DATEPART(yyyy, GETDATE()) AND CommitteeID=" & CommID & ") M"
			SQLJoin = "([Legislator Details] L INNER JOIN Politicians P ON L.PoliticianID = P.PoliticianID)"
			SQLJoin = "(" & SQLJoin & " LEFT JOIN " & strSQL & " ON L.LegislatorID = M.LegislatorID) "
			strSQL = "SELECT" & _
				" L.LegislatorID," & _
				" Chamber = CASE L.Seat WHEN 0 THEN 'Sen. ' ELSE 'Rep. ' END," & _
				" P.LastName + ', ' + P.FirstName + ' (' + L.Party + '-' + CAST(L.DistrictID AS varchar(2)) + ')' AS Name," & _
				" ISNULL(M.CommitteeID,0) " & _
				"FROM " & SQLJoin & _
				"WHERE (L.EndDate = CONVERT(DATETIME, '2299-12-31 00:00:00', 102)) " & SQLWhere & _
				"ORDER BY P.LastName, P.FirstName"
			Set rsLegs=Server.CreateObject("ADOR.Recordset")
			rsLegs.Open strSQL, strConnection
			aLegs = rsLegs.GetRows()
			rsLegs.Close
			Set rsLegs = Nothing
			intLegCount = UBound(aLegs,2)
			third = Int((intLegCount+1)/3)

			Response.Write _
				"<br><br><table width=750 cellpadding=0 cellspacing=0 border=0 class=det00 align=center>" & _
				"<colgroup span=3 width=250><tr valign=top><td>" & _
				"<input type=hidden name=LegCount value=" & intLegCount & ">"
			b = 0
			e = third
			For j = 1 to 3
				For i = b to e
					Response.Write "<input type=checkbox"
					If aLegs(3,i) = CommID Then Response.Write " checked disabled"
					Response.Write " name=Leg" & i & " value=" & aLegs(0,i) & "> &nbsp;"
					Response.Write aLegs(1,i) & aLegs(2,i) & "<br>"
				Next ' i
				Response.Write "</td><td>"
				b = j*third+j
				If j = 2 Then
					e = intLegCount
				Else
					e = (j+1)*third+1
				End If
			Next ' j

			Response.Write _
				"<tr style='height:40'><td colspan=3 align=center valign=bottom>" & _
				"<input type=submit value=Submit><span style='width:200'></span>" & _
				"<input type=button onclick='memCancel()' value=Cancel>" & _
				"</td></tr>"
			Response.Write _
				"</td></tr></table>"
		End If
	End If

	cxnSQL.Close
	Set cxnSQL = Nothing

%>
</form>


</body>
</html>
