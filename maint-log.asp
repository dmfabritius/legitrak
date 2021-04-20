<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body onload='selectTab(5)' class=bkg03>
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"

	strSQL = "SELECT * FROM [System Log] WHERE Severity IN ('Critical','Error') AND [TimeStamp] > GETDATE()-30 ORDER BY EntryID DESC"
	Set rsLog=Server.CreateObject("ADOR.Recordset")
	rsLog.Open strSQL, strConnection
	If Not rsLog.EOF Then
		Response.Write _
			"<table width=100% cellspacing=4 cellpadding=0 class=det00 style='cursor:default'>" & _
			"<col width=150><col width=50>"
		Do Until rsLog.EOF
			Response.Write _
			"<tr valign=top class=bkg04>" & _
			"<td>" & rsLog("TimeStamp") & "</td>" & _
			"<td>" & rsLog("Severity") & "</td>" & _
			"<td>" & rsLog("Activity") & "</td>" & _
			"</tr>"
			rsLog.MoveNext
		Loop
		Response.Write "</table>"
	Else
		Response.Write "No errors"
	End If
	rsLog.Close
	Set rsLog = Nothing
%>
</body>
</html>
