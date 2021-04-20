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
<td><a href='mobile-calendar.asp'>Calendar</a></td></tr>
<tr><td><a href='mobile-quickadd.asp'>Quick Add</a></td>
<td><a href='mobile-votecards.asp'>Vote Cards</a></td></tr>
<tr><td colspan=2><a href='mobile.asp'>Logout</a></td></tr>
</table><br>
<%
' CALENDAR ITEMS
	strSQLJoin = "[Calendar Bills] CB"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Daily Status] DS ON CB.[Bill Number]=DS.[Bill Number])"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Calendar] CA ON CB.CalendarID=CA.CalendarID)"
	strSQL = _
		"SELECT" & _
		" CB.[Bill Number]," & _
		" DS.Status, DS.Title," & _
		" CA.* " & _
		"FROM " & strSQLJoin & _
		"WHERE CA.CalendarID=" & Request.QueryString("id")
	Set rsCalendar=Server.CreateObject("ADOR.Recordset")
	rsCalendar.Open strSQL, strConnReadOnly

    If rsCalendar("TVW") Then
        strTVW = "TVW"
    Else
        strTVW = ""
    End If
        
    If Trim(rsCalendar("Title")) <> "" Then
        strTitle = rsCalendar("Title")
    Else
        strTitle = "(No Title Available)"
    End If

	strTime = FormatDateTime(rsCalendar("Time"), 3)
	strTime = Mid(strTime,1,Len(strTime)-6) & LCase(Right(strTime,3))
		
	strAgenda = Replace(rsCalendar("Agenda"),vbCrLf,"<br>")
	strAgenda = Replace(strAgenda,vbTab,"&nbsp; &nbsp; &nbsp; ")
	strAgenda = Replace(strAgenda,"  ","&nbsp; ")
		
'<!-- Client calendar item subsection -->
	Response.Write _
		"<table width=153 cellspacing=0 cellpadding=0 border=0>" & _
		"<col width=90><col width=138>" & _
		"<tr><td colspan=2><b>" & _
		rsCalendar("Status") & rsCalendar("Bill Number") & "</b></td></tr>" & _
		"<tr><td colspan=2><b>" & _
		strTitle & "</b></td></tr>"
		
	Response.Write _
		"<tr><td><br>" & _
		rsCalendar("Date") & "</td><td><br>" & _
		strTime & "</td></tr>" & _
		"<tr><td>" & _
		rsCalendar("Location1") & "</td><td>" & _
		strTVW & "</td></tr>"

	Response.Write _
		"<tr><td colspan=2><br>" & _
		rsCalendar("Committee") & "</td></tr>" & _
		"<tr><td colspan=2><br>" & _
		strAgenda & "</td></tr>" & _
		"</table>"
    
	rsCalendar.Close
	Set rsCalendar = Nothing
%>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>