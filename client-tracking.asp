<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<!--#include virtual="includes/client-tracking-post.asp"-->
<%
' CHECK TO MAKE SURE WE HAVE AN ACTIVE CLIENT WITH WHICH TO WORK
	If ClientID = 0 Then
		Response.Write "Need tracking list ID. Panic stop. Please report this error."
		Response.End
	End If

' LOAD LOCATION LEVEL LIST
	strSQL = "SELECT LocLevelID, Level FROM [Location Levels] ORDER BY LocLevelID"
	Set rsResults = Server.CreateObject("ADOR.Recordset")
	rsResults.Open strSQL, strConnReadOnly
	aLocs = rsResults.GetRows()
	rsResults.Close
	Set rsResults = Nothing
	intNumLocs = UBound(aLocs,2)+1

' LOAD FILTER PREFERENCES
	Dim PriSel(5), PosSel(6), LocSel(20)
	intFilterPri=CInt("0" & Request.Cookies("LegiTrak")("FilterPri"))
	PriSel(intFilterPri)=" selected"
	intFilterPos=CInt("0" & Request.Cookies("LegiTrak")("FilterPos"))
	PosSel(intFilterPos)=" selected"
	intFilterLoc=CInt("0" & Request.Cookies("LegiTrak")("FilterLoc"))
	LocSel(intFilterLoc)=" selected"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>u{color:blue;cursor:pointer}</style>
<script src="js/bts.js"></script>
<script src="js/client-tracking.js"></script>
</head>
<body onload='init()' onclick='hideDetail(arguments[0])' onscroll='scr()' onmousewheel='scr()' onresize='scr()' class=bkg03>
<iframe name=post style='display:none'></iframe>
<form id=BillSummaryForm method=post action='client-tracking.asp' style='margin:0'>
<table id=Bills width=100% cellspacing=4 cellpadding=0 class=det00 style='cursor:default'>
<col span=4 align=center><col align=left><col width=90%><col align=center>
<tr class=hdg29>
<td id=Sel style='cursor:pointer' onclick='selectBills(this)' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)' title="Select All">Sel</td>
<td style='cursor:pointer' onclick='sortBy("Bill")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Bill</td>
<td style='cursor:pointer' onclick='sortBy("Pri")' onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'>Priority</td>
<td style='cursor:pointer' onclick='sortBy("Pos")' onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'>Position</td>
<td style='cursor:pointer' onclick='sortBy("Location")' onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Location</td>
<td>Comments</td><td>Links</td>
</tr>
<tr class=bkg09>
<td id=checkmark class=lnk70 align=center valign=center onclick='selectMult()'><font face=Wingdings>&#252;</font></td>
<td class=lnk70 valign=center onclick='addDetail()'>New</td>
<td><select name=FilterPri onchange='updateFilter()'>
  <option value=0<%=PriSel(0)%>>All
  <option value=1<%=PriSel(1)%>>High
  <option value=2<%=PriSel(2)%>>Medium
  <option value=3<%=PriSel(3)%>>Low
  <option value=4<%=PriSel(4)%>>TBD
  <option value=5<%=PriSel(5)%>>Dead
</select></td>
<td><select name=FilterPos onchange='updateFilter()'>
  <option value=0<%=PosSel(0)%>>All
  <option value=1<%=PosSel(1)%>>Support
  <option value=2<%=PosSel(2)%>>Oppose
  <option value=3<%=PosSel(3)%>>Concerns
  <option value=4<%=PosSel(4)%>>Neutral
  <option value=5<%=PosSel(5)%>>Monitor
  <option value=6<%=PosSel(6)%>>-Blank-
</select></td>
<td><select name=FilterLoc onchange='updateFilter()' id="Select1">
<option value=0<%=LocSel(0)%>>All
<%
	For i = 1 to intNumLocs
		Response.Write _
			"<option value=" & i & LocSel(i) & ">" & _
			aLocs(1,i-1)
	Next ' i
%>
</select></td>
<td class=shd29>Change the selections to filter the list of bills.</td>
<td></td>
</tr>
<%
' BILL TRACKING
	OrderField = Request.Cookies("LegiTrak")("OrderField")
	If Len(OrderField) = 0 Then
		OrderField = "C.[Dead],C.[PriorityNum],C.[Bill Number]"
	End If

	strSQLWhere = ""
	intFilterPri=CInt("0" & Request.Cookies("LegiTrak")("FilterPri"))
	If intFilterPri = 5 Then
		strSQLWhere = " AND Dead=1"
	Else
		If intFilterPri <> 0 Then strSQLWhere = " AND PriorityNum=" & intFilterPri
	End If

	intFilterPos=CInt("0" & Request.Cookies("LegiTrak")("FilterPos"))
	If intFilterPos <> 0 Then strSQLWhere = strSQLWhere & " AND PositionNum=" & intFilterPos
	If intFilterLoc <> 0 Then
		strSQLWhere = strSQLWhere & " AND LocLevelID=" & aLocs(0,intFilterLoc-1)
		strSQLJoin = "[Client Specific Bill Info] C"
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Daily Status] D ON C.[Bill Number]=D.[Bill Number]) "
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Committees] COM ON D.CommitteeID=COM.CommitteeID) "
		strSQL = _
			"SELECT C.*, D.Title, D.Status," & _
			" CASE WHEN LTRIM(RTRIM(D.House))='' THEN D.Location ELSE D.House+', '+D.Location END AS HouseLoc " & _
			"FROM " & strSQLJoin & _
			"WHERE C.ClientID=" & ClientID & strSQLWhere & " ORDER BY " & OrderField
	Else
		strSQLJoin = "[Client Specific Bill Info] C"
		strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Daily Status] D ON C.[Bill Number]=D.[Bill Number]) "
		strSQL = _
			"SELECT C.*, ISNULL(D.Title,'(Title Unavailable)') AS Title, ISNULL(D.Status,'') AS Status," & _
			" CASE WHEN LTRIM(RTRIM(ISNULL(D.House,'')))='' THEN ISNULL(D.Location,'(Unavail.)') ELSE D.House+', '+ISNULL(D.Location,'(Unavail.)') END AS HouseLoc " & _
			"FROM " & strSQLJoin & _
			"WHERE C.ClientID=" & ClientID & strSQLWhere & " ORDER BY " & OrderField
	End If
	Set rsBillInfo=Server.CreateObject("ADOR.Recordset")
	rsBillInfo.Open strSQL, strConnReadOnly
	i=0
	Do Until rsBillInfo.EOF
		Select Case rsBillInfo("PriorityNum")
			Case 1: strPriority="High"
			Case 2: strPriority="Medium"
			Case 3: strPriority="Low"
			Case 4: strPriority="TBD"
			Case 5: strPriority="Dead"
       End Select
       Select Case rsBillInfo("PositionNum")
			Case 1: strPosition="Support"
          Case 2: strPosition="Oppose"
          Case 3: strPosition="Concerns"
          Case 4: strPosition="Neutral"
          Case 5: strPosition="Monitor"
          Case 6: strPosition="&nbsp;"
       End Select
		If rsBillInfo("Dead") Then
			strDead=" class=bkg08 style='padding:0 5'"
		Else
			strDead=""
		End If

		If Trim(rsBillInfo("Title")) <> "" Then
			strTitle=rsBillInfo("Title")
		Else
			strTitle="(No Title Available)"
		End If
		If Trim(rsBillInfo("Notes")) <> "" Then
			strNote= "(" & MakeHTML(rsBillInfo("Notes")) & ")"
		Else
			strNote=""
		End If
       
		Response.Write _
			"<tr valign=top class=bkg04>" & _
			"<td><input type=checkbox name=ckbx value=" & rsBillInfo("Bill Number") & "></td>" & _
			"<td id=bill" & i & _
			" class=lnk70 onclick='selectDetail(" & i & ")'>" & _
			rsBillInfo("Bill Number") & "</td>"
		Response.Write _
			"<td><span id=pri" & i & strDead & ">" & strPriority & "</span>" & _
			"<span id=dead" & i & " style='display:none'>" & rsBillInfo("Dead") & "</span></td>" & _
			"<td id=pos" & i & ">" & strPosition & "</td>" & _
			"<td>" & rsBillInfo("HouseLoc") & "</td>"
		Response.Write "<td>" & _
			"<span id=title" & i & " style='font-weight:bold'>" & strTitle & "</span> " & _
			"<span id=notes" & i & ">" & strNote & "</span><br>" & _
			"<span id=com" & i & ">" & MakeHTML(rsBillInfo("Comments")) & "</span></td>"
		Response.Write _
			"<td class=lnk40 onclick='lnk(arguments[0],""" & rsBillInfo("Status") & """)'>" & _
			"<u>D</u>_<u>F</u>_<u>A</u></td></tr>"
		rsBillInfo.MoveNext
		i=i+1
	Loop
	rsBillInfo.Close
	Set rsBillInfo=Nothing
%>
</table>

<!-- MULTIPLE BILL UPDATE BOX -->
<input type=hidden name=BillCount value=<%=i%>>
<div id=MultipleDetails class=div1A style="z-index:2;display:none;position:absolute;left:15;padding:5 0;height:140;width:95%;overflow:hidden">
<table border=0 cellspacing=0 cellpadding=0 class=hdg10 style='padding-left:10'>
<col width=90 align=right>
<tr><td>Priority:</td><td>
<select name=Pri style='width:95'>
  <option value=0>No Change
  <option value=1>High
  <option value=2>Medium
  <option value=3>Low
  <option value=4>TBD
</select></td></tr>
<tr><td>Dead:</td><td>
<select name=Dead style='width:95'>
  <option value=-1>No Change
  <option value=1>Yes
  <option value=0>No
</select></td></tr>
<tr><td>Position:</td><td>
<select name=Pos style='width:95'>
  <option value=0>No Change
  <option value=1>Support
  <option value=2>Oppose
  <option value=3>Concerns
  <option value=4>Neutral
  <option value=5>Monitor
  <option value=6>-Blank-
</select></td></tr>
<tr><td>Short Note:</td><td>
<input name=Notes maxlength=20><span style='position:relative;top:-3'> (Leave blank for no change)</span>
</td></tr>
<tr style='height:50'><td></td><td>
<input type=button class=btn61 value=Submit onclick='submitMult()'>
<input type=checkbox name=Delete value=Delete>Delete bill tracking information &nbsp;
<input type=button class=btn61 value=Cancel onclick='hideDetail(arguments[0],1)'>
</td></tr>
</table>
</div>
</form>
<div class=bkg04 style='position:relative;height:100%;margin:0 4'></div>

<!-- TRACKING DETAIL BOX -->
<div id=TrackingDetails class=div1A style="z-index:2;display:none;position:absolute;left:15;padding:5 0;height:200;width:95%;overflow:hidden">

<form id=BillDetailForm action="client-tracking.asp" method=post onsubmit='submitDetail()'>
<input type=hidden name=UpdateBillTracking value=True>
<input type=hidden name=Index value=-1>
<table width=100% border=0 cellspacing=0 cellpadding=0 class=hdg10 style='padding-left:10'>
<col width=90 align=right>
<tr><td>Bill Number:</td><td>
<input name=Bill size=4 maxlength=4 style="font-weight:bold" onchange='return isBill(this)' tabindex=1>
<input name=Title readonly class=hdg1A style='border:0px;margin:0 0 2 5;width:250'>
</td></tr>
<tr><td>Priority:</td><td>
<select name=Pri style='width:83' tabindex=2>
  <option value=High>High
  <option value=Medium>Medium
  <option value=Low>Low
  <option value=TBD>TBD
</select>
&nbsp; <input type=checkbox name=Dead value=True tabindex=8>&nbsp;Dead</td></tr>
<tr><td>Position:</td><td>
<select name=Pos style='width:83' tabindex=3>
  <option value=Support>Support
  <option value=Oppose>Oppose
  <option value=Concerns>Concerns
  <option value=Neutral>Neutral
  <option value=Monitor>Monitor
  <option value="">-Blank-
</select></td></tr>
<tr><td>Short Note:</td><td>
<input name=Notes maxlength=20 tabindex=4><span style='position:relative;top:-3'> (Maximum 20 characters)</span>
</td></tr>
<tr valign=top><td>Comments:
<iframe id=DigestFrame style='display:none'></iframe>
<br><input name=digCom type=button onclick='dig2Com()' style='font-size:10px;width:40px;height:35px;margin-right:13px'</td>
<td><textarea name=Comments rows=3 style='width:98%' tabindex=5></textarea>
</td></tr>
<tr style='height:50'><td></td><td>
<input type=submit class=btn61 value=Submit tabindex=6>
<input type=checkbox name=Delete value=Delete tabindex=9>Delete bill tracking information &nbsp;
<input type=button class=btn61 value=Cancel onclick='hideDetail(arguments[0],1)' tabindex=7>
</td></tr>
</table>
</form>
</div>

</body>
</html>
