<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set rsQueue=Server.CreateObject("ADOR.Recordset")

' REMOVE REQUESTED ENTRIES FROM THE QUEUE
    QID = Request.Form("QID")
	If Len(QID) <> 0 Then
'response.Write qid
        QIDs = Split(QID,",")
        QIDCnt = UBound(QIDs)
        For i = 0 to QIDCnt
            If Not IsNumeric(QIDs(i)) Then Response.End
        Next 'i

		strSQL = "DELETE FROM [Report Queue] WHERE ReportStatusID IN (" & QID & ")"
		Set cxnSQL = CreateObject("ADODB.Connection")
		cxnSQL.Open strConnection
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		cxnSQL.Close
		Set cxnSQL = Nothing
	End If

' DETERMINE NEXT QUEUE PROCESSING TIME (10 minute cycles)
	intHour = Hour(Now)
	intMin = Int((Minute(Now)+10)/10)*10
	If intMin = 60 Then
		intHour = intHour + 1
		strMin = "00"
	Else
		strMin = CStr(intMin)
	End If
	If intHour < 12 Then
		strAMPM = "AM"
	Else
		strAMPM = "PM"
	End If
	intHour = intHour mod 12
	If intHour = 0 Then intHour = 12
	strTime = intHour & ":" & strMin & " " & strAMPM

' BUILD COMMON SQL STATEMENT
	strSQLJoin = "[Report Queue] Q"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN Reports R ON Q.ReportID=R.ReportID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Client List] C ON Q.ClientID=C.ClientID) "
	strSQL = "SELECT" & _
		" Q.ReportStatusID, Q.[Effective Date], Q.[Report Created Date]," & _
		" R.[Report Display Name]," & _
		" ISNULL(C.[Short Company Name],'All Lists') Client " & _
		"FROM " & strSQLJoin & _
		"WHERE" & _
		 " Q.CustomerID=" & CustomerID & " AND"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg03 onload='selectTab(1)'>
<form method=post action='reports-requested.asp' style='margin:0'>
<table width=100% border=0 cellpadding=0 cellspacing=4 class=det00>
<col align=center><tr class=hdg29><td colspan=4>Next report processing cycle starts at <%=strTime%> (Current time is <%=Time%>)</td></tr>
<tr class=hdg29><td colspan=4>To delete selected report requests, click <input type=submit value=Delete></td></tr>
<tr class=hdg29><td>Sel</td><td>Report Name</td><td>Tracking List</td><td>Requested</td></tr>
<%
' QUEUED REPORTS
	rsQueue.Open strSQL & " Q.[Report Status]='REQUESTED' ORDER BY Q.[Effective Date]", strConnReadOnly
	If rsQueue.EOF Then
		Response.Write "<tr class=hdg14><td colspan=4 style='padding:40'>You have no reports queued for processing.</td></tr>"
	Else
		n = 0
		Do Until rsQueue.EOF
			n = n + 1
			Response.Write _
				"<tr class=bkg04><td><input type=checkbox name=QID value=" & rsQueue("ReportStatusID") & "></td>" & _
				"<td>" & rsQueue("Report Display Name") & "</td>" & _
				"<td>" & rsQueue("Client") & "</td>" & _
				"<td>" & rsQueue("Effective Date") & "</td></tr>"
			rsQueue.MoveNext
		Loop
	End If
	rsQueue.Close

	rsQueue.Open strSQL & " Q.[Report Status]='PROCESSING' ORDER BY Q.[Effective Date]", strConnReadOnly
	If Not rsQueue.EOF Then
		Response.Write _
			"<tr class=hdg24><td colspan=4>&nbsp;</td></tr>" & _
			"<tr class=hdg29><td colspan=3>Reports being processed</td><td>Requested</td></tr>"
		Do Until rsQueue.EOF
			n = n + 1
			Response.Write _
				"<tr class=bkg04><td></td>" & _
				"<td>" & rsQueue("Report Display Name") & "</td>" & _
				"<td>" & rsQueue("Client") & "</td>" & _
				"<td>" & rsQueue("Effective Date") & "</td></tr>"
			rsQueue.MoveNext
		Loop
	End If
	rsQueue.Close

	strSQL = Replace(strSQL, "SELECT", "SELECT TOP 5")
	rsQueue.Open strSQL & " Q.[Report Status]='CREATED' ORDER BY Q.[Report Created Date] DESC", strConnReadOnly
	If Not rsQueue.EOF Then
		Response.Write _
			"<tr class=hdg24><td colspan=4>&nbsp;</td></tr>" & _
			"<tr class=hdg29><td colspan=3>Recently completed reports</td><td>Completed</td></tr>"
		Do Until rsQueue.EOF
			n = n + 1
			Response.Write _
				"<tr class=bkg04><td></td>" & _
				"<td>" & rsQueue("Report Display Name") & "</td>" & _
				"<td>" & rsQueue("Client") & "</td>" & _
				"<td>" & rsQueue("Report Created Date") & "</td></tr>"
			rsQueue.MoveNext
		Loop
	End If
	rsQueue.Close

	Set rsQueue = Nothing
%>
</table>
</form>
<div class=bkg04 style='position:relative;height:100%;margin:0 4'></div>
</body>
</html>
