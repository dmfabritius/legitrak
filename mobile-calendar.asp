<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<meta name=HandheldFriendly content=true>
<meta name=PalmComputingPlatform content=true>
<link rel=stylesheet href="mobile-styles.css" type="text/css">
</head>
<body>
<b>LegiTrak</b> <i>Mobile!</i><br>
<b><%=CustomerName%></b><br><br>
<table width=153 cellspacing=0 cellpadding=0 border=0>
<tr><td><a href='mobile-customer.asp'>Tracking List</a></td>
<td><i>Calendar</i></td></tr>
<tr><td><a href='mobile-quickadd.asp'>Quick Add</a></td>
<td><a href='mobile-votecards.asp'>Vote Cards</a></td></tr>
<tr><td colspan=2><a href='mobile.asp'>Logout</a></td></tr>
</table><br>
<%
' CALENDAR ITEMS - LEGISLATIVE SESSION
	strSQLJoin = "[Client Specific Bill Info] CS"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Client List] CL ON CS.ClientID=CL.ClientID)"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Customer Clients] CC ON CL.ClientID=CC.ClientID)"
	strSQLJoin = "(" & strSQLJoin & "  LEFT JOIN [Daily Status] DS ON CS.[Bill Number]=DS.[Bill Number])"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Calendar Bills] CB ON CS.[Bill Number]=CB.[Bill Number])"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Calendar] CA ON CB.CalendarID=CA.CalendarID)"
	strSQL = _
		"(SELECT DISTINCT" & _
		"   CS.[Bill Number], DS.Status, DS.Title," & _
		"   MAX(CB.CalendarID) AS CalID," & _
		"   CA.Date, CA.Time," & _
		"   CA.Location1" & _
		" FROM " & strSQLJoin & _
		" GROUP BY" & _
		"   CS.[Bill Number], DS.Status, DS.Title," & _
 		"   CA.Date, CA.Time, CA.Location1, CC.CustomerID" & _
		" HAVING CC.CustomerID=" & CustomerID & _
		") AS CI"

	strSQLJoin = "(" & strSQL & " INNER JOIN [Calendar] CA ON CI.[CalID]=CA.[CalendarID])"
	strSQL = "SELECT" & _
		" CI.*, CA.Agenda, CA.TVW" & _
		" FROM " & strSQLJoin & _
		" ORDER BY CI.[Date], CI.[Time], CI.[Bill Number]"
	Set rsCalendar=Server.CreateObject("ADOR.Recordset")
	rsCalendar.Open strSQL, strConnReadOnly

	Response.Write "<table width=153 cellspacing=0 cellpadding=0 border=0>"
	i = 0
	prevDate = ""
	Do Until rsCalendar.EOF
        If Trim(rsCalendar("Title")) <> "" Then
            strTitle = rsCalendar("Title")
        Else
            strTitle = "(No Title Available)"
        End If
		If rsCalendar("Date") <> prevDate Then
			Response.Write "<tr><td colspan=3>"
			If prevDate <> "" Then Response.Write "<br>"
			Response.Write "<b>" & rsCalendar("Date") & "</b></td></tr>"
			prevDate = rsCalendar("Date")
		End If
		Response.Write "<tr><td><a href='mobile-cal-item.asp?id=" & rsCalendar("CalID") & "'>"
		Response.Write rsCalendar("Bill Number") & "</a> &nbsp; </td>"
		Response.Write "<td colspan=2>"
		Response.Write strTitle & "</td></tr>"
		calTime = rsCalendar("Time")
		calTime = mid(calTime,1,len(calTime)-6) & lcase(right(calTime,3))
		Response.Write "<tr><td></td>"
		Response.Write "<td>" & calTime & "</td>"
		Response.Write "<td>" & rsCalendar("Location1") & "</td>"
		Response.Write "</tr>"
		rsCalendar.MoveNext
		i = i + 1
	Loop
	If i = 0 Then Response.Write _
		"<tr><td align=center><br>" & _
		"There are no current Legislative session calendar items that correspond" & _
		" to your clients' tracking lists." & _
		"<br></td></tr>"
	Response.Write "</table>"

	rsCalendar.Close
	Set rsCalendar = Nothing
%>    
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>