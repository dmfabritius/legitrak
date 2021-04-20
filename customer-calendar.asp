<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg04>
<%
	Set rsCalendar=Server.CreateObject("ADOR.Recordset")

	bolGetCal = (Request.Cookies("LegiTrak")("SessionStatus") <> 3)

' ONLY ATTEMPT TO GET LEGISLATIVE CALENDAR ITEMS DURING THE REGULAR SESSION
	If bolGetCal Then
	
' CALENDAR ITEMS - LEGISLATIVE SESSION
        strSQLJoin = "[Client Specific Bill Info] CS"
        strSQLJoin = strSQLJoin & " INNER JOIN [Client List] CL ON CS.[ClientID]=CL.[ClientID]"
        strSQLJoin = strSQLJoin & " INNER JOIN [Customer Clients] CC ON CL.[ClientID]=CC.[ClientID]"
        strSQLJoin = strSQLJoin & " INNER JOIN [Calendar Bills] CB ON CS.[Bill Number]=CB.[Bill Number]"
        strSQLJoin = strSQLJoin & " INNER JOIN [Calendar] CA ON CB.[CalendarID]=CA.[CalendarID]"
        strSQLJoin = strSQLJoin & "  LEFT JOIN [Daily Status] DS ON CS.[Bill Number]=DS.[Bill Number]"
        strSQL = _
			"(SELECT" & _
            "  CS.[Bill Number]," & _
            "  MAX(CASE WHEN CS.PriorityNum = 4 THEN 0 ELSE CS.PriorityNum END) AS MaxPri," & _
            "  MIN(CS.PositionNum) AS PosNum," & _
			"  DS.Status, DS.Title," & _
			"  MAX(CB.CalendarID) AS CalID," & _
			"  CA.Date, CA.Time, CA.Location1" & _
            " FROM " & strSQLJoin & _
            " GROUP BY" & _
			"  CS.[Bill Number], DS.Status, DS.Title," & _
			"  CA.Date, CA.Time, CA.Location1," & _
			"  CC.CustomerID" & _
            " HAVING CC.CustomerID=" & CustomerID & _
            ") AS CI"
        strSQLJoin = strSQL & " INNER JOIN [Calendar] CA ON CI.CalID=CA.CalendarID "
        strSQL = _
			"SELECT CI.*, CA.TVW, CA.Agenda " & _
            "FROM " & strSQLJoin & _
            "ORDER BY CI.[Date], CI.[Time], CI.[Bill Number]"
		rsCalendar.Open strSQL, strConnReadOnly

		Response.Write _
			"<br><table width=350 border=0 cellspacing=0 cellpadding=0 class=det00>" & _
			"<col width=35><col width=8><col width=190><col width=67 align=right><col width=50>"
		i = 0
		prevDate = ""
		Do Until rsCalendar.EOF
			If Trim(rsCalendar("Title")) <> "" Then
				strTitle = rsCalendar("Title")
			Else
				strTitle = "(No Title Available)"
			End If
			Select Case rsCalendar("MaxPri")
				Case 1: strPri = "<td align=center style='color:red;font:bold 10pt Georgia'>!</td>"
				Case 3: strPri = "<td align=center><font face='Wingdings 3' color=blue>&#148;</font></td>"
				Case Else: strPri = "<td></td>"
			End Select

			If rsCalendar("Date") <> prevDate Then
				Response.Write "<tr><td colspan=4 class=hdg24>"
				If prevDate <> "" Then Response.Write "<br>"
				Response.Write _
					WeekDayName(WeekDay(rsCalendar("Date"))) & ", " & _
					MonthName(Month(rsCalendar("Date")),True) & " " & _
					Day(rsCalendar("Date")) & "</td></tr>"
				prevDate = rsCalendar("Date")
			End If

			strStyle = " style='font-weight:bold;color:"
			Select Case rsCalendar("PosNum")
				Case 1: strStyle = strStyle & "#009000'"
				Case 2: strStyle = strStyle & "red'"
				Case 3: strStyle = strStyle & "orange'"
				Case Else
					strStyle = ""
			End Select

			Response.Write "<tr " & strStyle & ">"
			Response.Write "<td>" & rsCalendar("Bill Number") & "</td>"
			Response.Write strPri
			Response.Write "<td nowrap>" & strTitle & "</td>"
			calTime = rsCalendar("Time")
			calTime = Mid(calTime,1,Len(calTime)-6) & LCase(Right(calTime,3))
			If Left(calTime,1) = "0" Then calTime = Mid(calTime,2)
			Response.Write "<td>" & calTime & "</td>"
			Response.Write "<td style='padding-left:10'>" & rsCalendar("Location1") & "</td>"
			Response.Write "</tr>"
			rsCalendar.MoveNext
			i = i + 1
		Loop
		Response.Write "</table><br>"

		If i = 0 Then Response.Write _
			"<div class=hdg24 style='margin:25'>" & _
			"There are no current Legislative session calendar items<br>" & _
			"that correspond to your clients' tracking lists.</div><br>"

		rsCalendar.Close
	End If

	If Not bolGetCal Then
%>
<iframe src="http://www.google.com/calendar/embed?src=MzA0ZTVucGhkdGMwOWFoaHM5ZmQ2aDcxb2tAZ3JvdXAuY2FsZW5kYXIuZ29vZ2xlLmNvbQ" style=" border-width:0 " width="1024" height="768" frameborder="0" scrolling="no"></iframe>
<%

	End If

%>
</body>
</html>
