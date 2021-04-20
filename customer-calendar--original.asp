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

' CALENDAR ITEMS - CAMPAIGN EVENTS
	strSQLJoin = "[Campaign Events]"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Campaign Event Candidates] ON [Campaign Events].EventID = [Campaign Event Candidates].EventID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Candidates] ON [Campaign Event Candidates].CandidateID = Candidates.CandidateID)"

	strSQL = "SELECT DISTINCT" & _
		" [Campaign Events].[Date], [Campaign Events].City," & _
		" Candidates.[LastName], Candidates.Party, Candidates.DistrictID" & _
		" FROM " & strSQLJoin & _
		" WHERE [Campaign Events].[Date] >= CONVERT(varchar(10),GETDATE(),120)" & _
		" ORDER BY [Campaign Events].[Date], Candidates.[LastName]"
	rsCalendar.Open strSQL, strConnReadOnly

	If Not rsCalendar.EOF And bolGetCal Then
		Response.Write "<br><span class=hdg24>Campaign Events</span><br><br>"
	End If

	Response.Write _
		"<br><table width=350 border=0 cellspacing=0 cellpadding=0 class=det00>" & _
		"<col width=15><col width=200><col width=135>"
	i = 0
	prevDate = ""
	Do Until rsCalendar.EOF
		If rsCalendar("Date") <> prevDate Then
			Response.Write "<tr><td colspan=3 class=hdg24>"
			If prevDate <> "" Then Response.Write "<br>"
			Response.Write _
				WeekDayName(WeekDay(rsCalendar("Date"))) & ", " & _
				MonthName(Month(rsCalendar("Date")),True) & " " & _
				Day(rsCalendar("Date")) & "</td></tr>"
			prevDate = rsCalendar("Date")
		End If

       If Trim(rsCalendar("City")) <> "" Then
           strCity = rsCalendar("City")
       Else
           strCity = "(TBA)"
       End If
		Response.Write "<tr><td></td>"
		Response.Write "<td>" & rsCalendar("LastName") & " ("
		If rsCalendar("Party") <> "" Then
			Response.Write rsCalendar("Party")
		Else
			Response.Write "-"
		End If
		If rsCalendar("DistrictID") <> 0 Then
			Response.Write "-" & rsCalendar("DistrictID")
		End If
		Response.Write ")</td>"
		Response.Write "<td>" & strCity & "</td>"
		Response.Write "</tr>"
		rsCalendar.MoveNext
		i = i + 1
	Loop
	Response.Write "</table>"
	If Not bolGetCal And i = 0 Then Response.Write "<div id=NoEvt class=hdg24 style='margin:25'>There are no current campaign events.</div>"

	rsCalendar.Close
	Set rsCalendar = Nothing
%>
</body>
</html>
