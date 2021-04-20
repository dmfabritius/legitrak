<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script src="js/reports-customize.js"></script>
</head>
<body onload='init()' class=shd24 style='margin:0'>
<div style='display:none'>
<%
	Set rsResults=Server.CreateObject("ADOR.Recordset")

' LOAD CUSTOMER REPORT PREFERENCES AND CUSTOM SECTIONS
	strSQL = _
		"SELECT C.[Report Style], C.[Report Format], C.[Automatic Daily], C.[Active Report], ISNULL(C.ContribMatrix,'R') CM, R.* " & _
		"FROM [Customer List] C LEFT JOIN [Customer Custom Reports] R ON C.CustomerID=R.CustomerID " & _
		"WHERE C.CustomerID=" & CustomerID & " ORDER BY R.CustomID"
	rsResults.Open strSQL, strConnReadOnly
	Response.Write _
		"<div id=R0>" & _
		rsResults("Report Style") & "," & _
		rsResults("Report Format") & "," & _
		rsResults("Active Report") & "," & _
		rsResults("Automatic Daily") & ";"
	strCustCM = rsResults("CM")
	c=1
	Do Until rsResults.EOF
		If rsResults("CustomID") <> c Then
			Response.Write ";"
			c = rsResults("CustomID")
		End If
		Response.Write "," & rsResults("SubReportID")
		rsResults.MoveNext
	Loop
	Response.Write ";;</div>"
	rsResults.Close


' LOAD CUSTOMER'S CLIENT REPORT PREFERENCES AND CUSTOM SECTIONS
	intNumClts = 0
	strSQL = _
		"SELECT" & _
		" CC.ClientID AS CltID, R.*," & _
		" CL.[Client Company Name], CL.[Short Company Name], CL.[Report Comments Header], CL.[Report Comments]," & _
		" CL.[Report Style], CL.[Report Format], CL.[Active Report], CL.[Automatic Weekly]," & _
		" ISNULL(CL.[Automatic Register],0) AutoReg, ISNULL(CL.ContribMatrix,0) CM, ISNULL(CL.[Report Priority],0) RptPri " & _
		"FROM" & _
		" [Client List] CL" & _
		" INNER JOIN [Customer Clients] CC ON CL.ClientID=CC.ClientID" & _
		" LEFT JOIN [Client Custom Reports] R ON CC.ClientID = R.ClientID " & _
		"WHERE CC.CustomerID=" & CustomerID & " ORDER BY CL.[Short Company Name], R.CustomID"
	rsResults.Open strSQL, strConnReadOnly

	' Create array of client information
	Dim aClts(4,100)
	clt=0
	Do Until rsResults.EOF
		If rsResults("CltID") <> clt Then
			If clt <> 0 Then Response.Write ";;</div>"
			intNumClts = intNumClts+1
			aClts(0,intNumClts) = Encrypt(rsResults("CltID"))
			aClts(1,intNumClts) = rsResults("Client Company Name")
			aClts(2,intNumClts) = rsResults("Short Company Name")
			aClts(3,intNumClts) = rsResults("AutoReg")
			aClts(4,intNumClts) = rsResults("CM")
			Response.Write _
				"<div id=H" & aClts(0,intNumClts) & ">" & rsResults("Report Comments Header") & "</div>" & _
				"<div id=C" & aClts(0,intNumClts) & ">" & rsResults("Report Comments") & "</div>" & _
				"<div id=R" & aClts(0,intNumClts) & ">" & _
				rsResults("Report Style") & "," & _
				rsResults("Report Format") & "," & _
				rsResults("Active Report") & "," & _
				rsResults("Automatic Weekly") & "," & _
				rsResults("RptPri") & ";"
			clt = rsResults("CltID")
			c=1
		End If
		If rsResults("CustomID") <> c Then
			Response.Write ";"
			c = rsResults("CustomID")
		End If
		Response.Write "," & rsResults("SubReportID")
		rsResults.MoveNext
	Loop
	If clt <> 0 Then Response.Write ";;</div>"
	rsResults.Close
%>
</div>
<iframe name=post style='display:none'></iframe>
<form id=RptForm method=post target=post action='reports-customize-post.asp' onsubmit='updateCache()'>
<%
	strClts = ""
	For i = 1 to intNumClts
		strClts = strClts & "," & aClts(0,i)
	Next 'i
	Response.Write "<input type=hidden name=Clients value='" & strClts & "'>"
%>
<table border=0 cellpadding=0 cellspacing=0 class=hdg24 style='margin:10'>
<col width=200><col width=300>
<tr><td>Report to customize:</td><td><select name=RptCustom style='width:300' onchange='updateActive(null,this.value)'>
<option value=0>Daily Report
<option value=29>Contribution Analysis
<option value=14>Register Rated by Keywords
<%
	For i = 1 to intNumClts
		Response.Write "<option value='" & aClts(0,i) & "'>Weekly Report - " & aClts(2,i)
	Next 'i
%>
</select></td></tr>
<tr><td>Active report:</td><td id=ActiveGrp class=shd24>
<div id=ActiveRpt name=ActiveRpt style='display:none'><!-- Weekly HTML -->
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=2>Default
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=19 style='margin-left:15'>Custom #1
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=20 style='margin-left:15'>Custom #2
</div>
<div id=ActiveRpt name=ActiveRpt style='display:none'><!-- Weekly Excel (also used for Contrib Matrix, Auto Reg) -->
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=24>Default
</div>
<div id=ActiveRpt name=ActiveRpt style='display:none'><!-- Daily HTML -->
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=1>Default
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=17 style='margin-left:15'>Custom #1
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=18 style='margin-left:15'>Custom #2
</div>
<div id=ActiveRpt name=ActiveRpt style='display:none'><!-- Daily Excel -->
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=25>Default
<input type=radio onmouseup='updateActive(this.value)' name=Rpt value=27 style='margin-left:15'>Columnar
</div>
</td></tr>
<tr><td>Generate automatically:</td><td id=AutoGrp class=shd24>
<input type=radio name=Auto value=1 onclick='mark(autoGrp)' checked>Yes
<input type=radio name=Auto value=0 onclick='mark(autoGrp)' style='margin-left:15'>No
</td></tr><tr height=8><td></td></tr>
<tr><td>Report format:</td><td><select name=RptFormat style='width:200' onchange='updateFormat(this)'>
<option value='.doc '>Word
<option value='.html'>HTML message
<option value='.htm '>HTML attachment
<option value='.xls '>Excel
</select></td></tr>
<tr><td>Report Style:</td><td><select name=RptStyle style='width:200' onchange='updateStyle(this)'>
<option value=1>Book Antiqua - Blue
<option value=2>Book Antiqua - Green
<option value=3>Book Antiqua - Black
<option value=4>Times Roman/Arial
</select></td></tr>
</table>

<table class=shd24 style='margin:0;width:100%;border:0 solid white;border-width:3px 0'><tr><td align=center>
<input type=submit value=Submit>
<span id=ApplyToAll style='display:none'><input type=checkbox name=ApplyToAll value=True>Apply to all tracking lists</span>
<input type=button value=Cancel style='margin:0 50' onclick='updateActive()'>
</td></tr></table>

<table width=97% border=0 cellpadding=0 cellspacing=0 class=shd24 style='margin:10'>
<col><col width=100%>
<tr valign=top>
<td><div style='background-color:gray;margin:10;width:260;height:340'>
<div style='background-color:white;position:relative;top:-5;left:-5;width:100%;height:100%;border:1px solid black'>
<iframe name=SR frameborder=0 width=100% height=100% scrolling=no src='report-preview.htm'></iframe></div>
</div></td>
<td>
<%
' WEEKLY REPORT CHOICES
'----------------------
	Response.Write "<div id=Choices name=Choices style='display:none'>"
	strSQL = _
		"SELECT ISNULL([Report Comments Header],'') Header, ISNULL([Report Comments],'') Comm " & _
		"FROM [Customer List] WHERE CustomerID=" & CustomerID
	rsResults.Open strSQL, strConnReadOnly
'	- GENERAL COMMENTS
	Response.Write _
		"<br>General Comments:" & _
		"<br><input name=GH style='width:100%' onchange='mark(this)' onkeyup='updateComm()' value='" & TweakQuote(rsResults("Header")) & "'>" & _
		"<textarea name=GC rows=4 style='width:100%' onchange='mark(this)' onkeyup='updateComm()'>" & rsResults("Comm") & "</textarea><br>"
	rsResults.Close
	Set rsResults = Nothing

'	- SPECIFIC COMMENTS
	Response.Write _
		"Tracking List Specific Comments:" & _
		"<br><input name=SH style='width:100%' onchange='mark(this)' onkeyup='updateComm()'>" & _
		"<textarea name=SC rows=4 style='width:100%' onchange='mark(this)' onkeyup='updateComm()'></textarea><br>"

'	- REPORT SECTIONS
	For i = 1 to 3
		Response.Write _
			"<br><span style='width:110'>Detail section #" & i & ":</span><select name=RptSec onchange='updateSec(-1,this," & 3+i & ")'>" & _
			"<option value=0>-NONE-" & _
			"<option value=8>Bill Tracking Summary" & _
			"<option value=11>Bill Tracking Summary w/ Comments" & _
			"<option value=12>Bill Tracking Summary w/ Notes" & _
			"<option value=9>Calendar References" & _
			"<option value=10>Bill Details" & _
			"</select>"
	Next 'i

	Response.Write _
	"<br><br><span style='width:110'>Bill priorities:</span><select name=RptPri onchange='updatePri(this)'>" & _
	"<option value=0>All" & _
	"<option value=1>High" & _
	"<option value=2>Medium" & _
	"<option value=3>Low" & _
	"<option value=4>TBD</select>"

	Response.Write "</div><div id=Choices name=Choices style='display:none'>"
	
' DAILY REPORT CHOICES
'----------------------
	For i = 1 to 4
		Response.Write _
			"<br>Detail section #" & i & ":&nbsp;<select name=RptSec onchange='updateSec(5,this," & 3+i & ")'>" & _
			"<option value=0>-NONE-" & _
			"<option value=1>Calendar References" & _
			"<option value=2>Bills with Activity" & _
			"<option value=3>Bills Rated by Keywords" & _
			"<option value=4>Bill Details: Keywords" & _
			"<option value=5>Bill Details: Calendar" & _
			"</select>"
	Next 'i

	Response.Write "</div><div id=Choices name=Choices class=det00 style='display:none'>"
	
' REGISTER RATINGS REPORT CHOICES
'--------------------------------
	Response.Write "<div class=shd24>" & _
		"<br>The Washington State Register is imported on the 1st and 3rd Wednesday of every month. " & _
		"Select the Tracking Lists for which you would like to automatically receive a report rating the article texts by the lists' keywords:</div>"
	For i = 1 to intNumClts
		Response.Write "<br><input type=checkbox name=CltAutoReg value=" & aClts(0,i)
		If aClts(3,i) = 1 Then Response.Write " checked"
		Response.Write "> " & aClts(1,i)
	Next 'i

	Response.Write "</div><div id=Choices name=Choices class=det00 style='display:none'>"
	
' CONTRIBUTION MATRIX REPORT CHOICES
'-----------------------------------
	If strCustCM = "R" Then
		strRec = " checked"
	ElseIf strCustCM = "A" Then
		strAct = " checked"
	Else
		strBoth = " checked"
	End If
	Response.Write _
		"<br><div class=shd24>Contribution Type:</div>" & _
		"<input type=radio name=ContribType value=R" & strRec & ">Recommendations" & _
		"<input type=radio name=ContribType value=A style='margin-left:15'" & strAct & ">Actual" & _
		"<input type=radio name=ContribType value=B style='margin-left:15'" & strBoth & ">Both" & _
		"<br><br>" & _
		"<div class=shd24>Tracking Lists to include on the report:</div>"
	For i = 1 to intNumClts
		Response.Write "<br><input type=checkbox name=CltCM value=" & aClts(0,i)
		If aClts(4,i) = 1 Then Response.Write " checked"
		Response.Write "> " & aClts(1,i)
	Next 'i

	Response.Write "</div>"
%>
</td>
</tr></table>

</form>
</body>
</html>
