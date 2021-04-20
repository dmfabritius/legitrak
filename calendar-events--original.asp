<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
</head>
<body class=bkg04 style='margin:0 3'>
<%
' CAMPAIGN EVENTS
	strSQLJoin = "[Campaign Events]"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Campaign Event Candidates] ON [Campaign Events].EventID = [Campaign Event Candidates].EventID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Candidates] ON [Campaign Event Candidates].CandidateID = Candidates.CandidateID)"

	strSQL = "SELECT" & _
		" Candidates.[FirstName], Candidates.[LastName], Candidates.Party, Candidates.DistrictID," & _
		" [Campaign Events].EventID, [Campaign Events].Title," & _
		" [Campaign Events].[Date], [Campaign Events].[Time], [Campaign Events].[Length]," & _
		" [Campaign Events].[Address Line 1], [Campaign Events].[Address Line 2]," & _
		" [Campaign Events].City, [Campaign Events].State, [Campaign Events].Zip," & _
		" [Campaign Events].Comments" & _
		" FROM " & strSQLJoin & _
		" WHERE [Campaign Events].[Date] >= CONVERT(varchar(10),GETDATE(),120)" & _
		" ORDER BY " & _
		" [Campaign Events].[Date], [Campaign Events].[Time], [Campaign Events].[EventID], Candidates.[LastName]"
	Set rsCalendar=Server.CreateObject("ADOR.Recordset")
	rsCalendar.Open strSQL, strConnReadOnly

	dtPrevDate = Now()
	intPrevEventID = 0

	Response.Write "<br><table border=0 cellspacing=0 cellpadding=0 class=det00>"

    Do Until rsCalendar.EOF
        
'<!-- Date header section -->
        If rsCalendar("Date") <> dtPrevDate Then
            dtPrevDate = rsCalendar("Date")
            Response.Write "<tr><td colspan=3 class=hdg24>" & rsCalendar("Date") & "</td></tr>"
        End If

'<!-- Format the event details -->
        strLocation = ""
        If Len(rsCalendar("Address Line 1")) <> 0 Then
			strLocation = strLocation & rsCalendar("Address Line 1")
		End If
        If Len(rsCalendar("Address Line 2")) <> 0 Then
			strLocation = strLocation & "<br>" & rsCalendar("Address Line 2")
		End If
		If Len(rsCalendar("City")) <> 0 Then
			strLocation = strLocation & "<br>" & rsCalendar("City")
		End If
		If Len(rsCalendar("State")) <> 0 Then
			strLocation = strLocation & ", " & rsCalendar("State")
		End If
		If Len(rsCalendar("Zip")) <> 0 Then
			strLocation = strLocation & "&nbsp; " & rsCalendar("Zip")
		End If

		If Len(rsCalendar("Time")) <> 0 Then
			strTime = FormatDateTime(rsCalendar("Time"), 3)
			strTime = Mid(strTime,1,Len(strTime)-6) & LCase(Right(strTime,3))
			If CInt(rsCalendar("Length")) <> 0 Then
				strTime = strTime & "<br>" & rsCalendar("Length") & " hour"
				If rsCalendar("Length") <> 1 Then strTime = strTime & "s"
			End If
		Else
			strTime = ""
		End If
			
		If Len(rsCalendar("Comments")) <> 0 Then
			strComments = Replace(rsCalendar("Comments"),vbCrLf,"<br>")
			strComments = Replace(strComments,vbCr,"&nbsp; ")
			strComments = Replace(strComments,vbLf,"&nbsp; ")
			strComments = Replace(strComments,vbTab,"&nbsp; &nbsp; &nbsp; ")
			strComments = Replace(strComments,"  ","&nbsp; ")
		Else
			strComments = ""
		End If
		
'<!-- Display an event -->
			
'			 style='cursor:hand' onclick='gotoEvent(" & rsCalendar("EventID") & ")'" & _
'			" onMouseOver='highLight(1,0,1)' onMouseOut='highLight(0,0,1)'>"

			
'			 & _
'			"<col width=330 class=bt1><col width=290 class=bt1><col width=100 class=bt1>"
			
		Response.Write "<tr valign=top><td><b>" & rsCalendar("Title") & "</b><br>"
		' Display all of the candidates for this event in the first table cell
		' separated by line breaks
		intPrevEventID = rsCalendar("EventID")
		Do
			Response.Write rsCalendar("LastName")
			If Len(rsCalendar("FirstName")) <> 0 Then
				Response.Write  ", " & rsCalendar("FirstName")
			End If
			If rsCalendar("Party") <> "" Then
				Response.Write " (" & rsCalendar("Party")
			Else
				Response.Write " (" & "-"
			End If

			If rsCalendar("DistrictID") <> 0 Then
				Response.Write "-" & rsCalendar("DistrictID")
			End If
 			Response.Write ")<br>"
	       rsCalendar.MoveNext
	       If rsCalendar.EOF Then Exit Do
		Loop Until rsCalendar("EventID") <> intPrevEventID
		Response.Write "</td>"

		' Display the rest of the event details
		Response.Write "<td>" & strLocation & "</td>"
		Response.Write "<td>" & strTime & "</td></tr>"

		Response.Write "<tr><td colspan=3><div style='margin-left:40;margin-right:40'>"
		Response.Write strComments & "</div><br></td><tr>"

    Loop
    
	rsCalendar.Close
	set rsCalendar = Nothing
%>    
</body>
</html>
