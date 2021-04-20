<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg04>
<%
' CALENDAR ITEMS
	strSQLJoin = "[Client Specific Bill Info] CS"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Client List] CL ON CS.ClientID=CL.ClientID)"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Customer Clients] CC ON CL.ClientID=CC.ClientID)"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Calendar Bills] CB ON CS.[Bill Number]=CB.[Bill Number])"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Calendar] CAL ON CB.CalendarID=CAL.CalendarID)"
	strSQLJoin = "(" & strSQLJoin & "  LEFT JOIN [Daily Status] D ON CS.[Bill Number]=D.[Bill Number]) "
	strSQL = _
		"(SELECT DISTINCT" & _
		" CS.[Bill Number], D.Status, D.Title," & _
		" Max(CB.CalendarID) AS CalID," & _
		" CAL.Date, CAL.Time, CAL.Location1," & _
		" Min(CAL.Committee) AS Comm1, Max(CAL.Committee) AS Comm2 " & _
		"FROM " & strSQLJoin & _
		"GROUP BY" & _
		" CS.[Bill Number], D.Status, D.Title," & _
 		" CAL.Date, CAL.Time, CAL.Location1," & _
		" CC.CustomerID " & _
		"HAVING CC.CustomerID=" & CustomerID & _
		") AS CI"
	strSQLJoin = strSQL & " INNER JOIN [Calendar] ON CI.CalID=[Calendar].CalendarID "
	strSQL = _
		"SELECT CI.*, [Calendar].TVW, [Calendar].Agenda " & _
		"FROM " & strSQLJoin & _
		"ORDER BY [Calendar].[Date], [Calendar].[Time], CI.[Bill Number]"
	Set rsCalendar=Server.CreateObject("ADOR.Recordset")
	rsCalendar.Open strSQL, strConnReadOnly

	dtPrevDate = Now()

	Response.Write _
		"<br><table width=98% border=0 cellspacing=0 cellpadding=0 class=det00>" & _
		"<col span=3><col width=35><col width=65><col width=40>"

	i = 0
    Do Until rsCalendar.EOF
		i = i + 1
'<!-- Date header section -->
        If rsCalendar("Date") <> dtPrevDate Then
            dtPrevDate = rsCalendar("Date")
            Response.Write _
                "<tr><td colspan=6 class=hdg24>" & _
				WeekDayName(WeekDay(rsCalendar("Date"))) & ", " & _
				MonthName(Month(rsCalendar("Date")),True) & " " & _
				Day(rsCalendar("Date")) & "</td></tr>"
        End If
        
        If rsCalendar("TVW") Then
            strTVW = "TVW"
        Else
            strTVW = ""
        End If
        
        If Trim(rsCalendar("Comm2")) <> "" And rsCalendar("Comm2") <> rsCalendar("Comm1") Then
            strComm2 = "<tr><td colspan=2></td><td colspan=4>" & _
                       "Jt. w/" & rsCalendar("Comm2") & "</td></tr>"
        Else
            strComm2 = ""
        End If

        If Trim(rsCalendar("Title")) <> "" Then
            strTitle = rsCalendar("Title")
        Else
            strTitle = "(No Title Available)"
        End If

	strTime = FormatDateTime(rsCalendar("Time"), 3)
	strTime = Mid(strTime,1,Len(strTime)-6) & LCase(Right(strTime,3))
	If Left(strTime,1) = "0" Then strTime = Mid(strTime,2)
		
	strAgenda = MakeHTML(rsCalendar("Agenda"))
	strAgenda = Replace(strAgenda,vbTab,"&nbsp; &nbsp; &nbsp; ")
	strAgenda = Replace(strAgenda,"  ","&nbsp; ")
		
'<!-- Client calendar item subsection -->
	Response.Write _
            "<tr class=shd24 valign=top><td>" & rsCalendar("Status") & rsCalendar("Bill Number") & " &nbsp;</td>" & _
            "<td>" & strTitle & "</td>" & _
            "<td>" & rsCalendar("Comm1") & "</td>" & _
            "<td>" & strTVW & "</td>" & _
            "<td nowrap>" & strTime & "</td>" & _
            "<td align=right>" & rsCalendar("Location1") & "</td></tr>" & _
            strComm2 & _
            "<tr><td></td><td colspan=5>" & strAgenda & "<br><br></td></tr>"
        rsCalendar.MoveNext
    Loop
	Response.Write "</table><br>"

	If i = 0 Then Response.Write _
		"<div class=hdg24 style='margin:25'>" & _
		"There are no current Legislative session calendar items<br>" & _
		"that correspond to your clients' tracking lists.</div><br>"
    
	rsCalendar.Close
	Set rsCalendar = Nothing
%>    
</body>
</html>