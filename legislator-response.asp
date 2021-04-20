<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
' COLLECT URL PARAMETERS

	On Error Resume Next
	VoteRequestID = CInt("0" & Request.QueryString("vid"))
	If VoteRequestID = 0 Then Response.Redirect "/"

	LegislatorID = CInt("0" & Request.QueryString("lid"))
	If LegislatorID = 0 Then Response.Redirect "/"

	SecurityKey = Request.QueryString("key")
	If Request.QueryString("response") = "yes" Then
		intVote = 1
	Else
		intVote = 5
	End If
	
' PROCESS REPONSE
		Set cmdSQL = CreateObject("ADODB.Connection")
		With cmdSQL
			.CommandTimeout = 5 ' seconds
			.Open strConnection

		' Look up security key
			strCommand = _
				"SELECT * " & _
				"FROM [Vote Request Queue] V INNER JOIN [Vote Request Details] D " & _
				"ON V.VoteRequestID = D.VoteRequestID " & _
				"WHERE D.VoteRequestID=" & VoteRequestID & " AND D.LegislatorID=" & LegislatorID
			Set rsResult = .Execute(strCommand)
			If Not rsResult.EOF Then bolAuthenticated = (rsResult("SecurityKey") = SecurityKey)

		' Record response
			If bolAuthenticated Then
			' Delete any previous entry
				VotecardID = rsResult("VotecardID")
				strCommand = _
					"DELETE FROM [Votecard Details] " & _
					"WHERE VotecardID=" & VotecardID & " AND LegislatorID=" & LegislatorID
				.Execute strCommand, , adExecuteNoRecords
			' Add new vote response
				strCommand = "INSERT INTO [Votecard Details] VALUES (" & _
					VotecardID & "," & _
					LegislatorID & "," & _
					intVote & "," & _
					"'" & Now & "'," & _
					"1," & _
					"NULL)"
				.Execute strCommand, , adExecuteNoRecords
			' Reset security key
			' ** The queue entries are deleted after 15 days by the nightly batch processing
'				strCommand = _
'					"UPDATE [Vote Request Details] " & _
'					"SET SecurityKey='" & CStr(Now) & "' " & _
'					" WHERE" & _
'						" VoteRequestID=" & VoteRequestID & " AND" & _
'						" LegislatorID=" & LegislatorID
'				.Execute strCommand, , adExecuteNoRecords
			End If

			.Close
		End With
		Set cmdSQL = Nothing
%>
<html>

<head>
<title>Welcome to LegiTrak</title>
<script src="js/bts.js"></script>
</head>
<body style="font:12pt Tahoma, Arial">
<%
	If bolAuthenticated Then
		Response.Write "<h2>Thank you for your response!</h2>"
	Else
		Response.Write "<h2>Sorry! No matching vote request was found.</h2>"
	End If
%>
To visit our home page, click here: <a href="/">LegiTrak</a>

</body>
</html>