<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "mobile-customer.asp"
	End If

' UPDATE CLIENT SPECIFIC BILL INFO
	If Request.Form("UpdateEntry") = "Submit" Then
		On Error Resume Next
		BillNum = CInt(Request.Form("BillNum"))
		If BillNum > 1000 And BillNum < 9000 Then

			Response.Cookies("LegiTrak")("ClientID") = Request.Form("ClientID")
			ClientID = Decrypt(Request.Form("ClientID"))
			If ClientID = 0 Then Response.End

			Set cxnSQL = CreateObject("ADODB.Connection")
			With cxnSQL
			    .Open strConnection
			    strSQL = _
			    	"SELECT * FROM [Client Specific Bill Info] " & _
					"WHERE ClientID=" & ClientID & _
					" AND [Bill Number]=" & BillNum
				Set rsResult = .Execute(strSQL)

		' If no record was returned, then go ahead and add the new record.
				If rsResult.EOF Then
					strSQL = _
						"INSERT [Client Specific Bill Info]" & _
						" (ClientID, [Bill Number], PriorityNum, PositionNum, Comments) " & _
						"VALUES (" & _
						ClientID & "," & _
						BillNum & "," & _
						CInt(Request.Form("Pri")) & "," & _
						CInt(Request.Form("Pos")) & ",'" & _
						TweakQuote(Request.Form("Com")) & "')"
					.Execute strSQL, , adExecuteNoRecords
				Else
					strInvalid = "<b><i>Duplicate bill number.&nbsp; Please re-enter.</i></b>"
				End If
				.Close
			End With
			Set rsResult = Nothing
			Set cxnSQL = Nothing
			If Len(strInvalid) = 0 Then Response.Redirect "mobile-customer.asp"
		Else
			strInvalid = "<b><i>Invalid bill number.&nbsp; Please re-enter.</i></b>"
		End If
	End If
%>
<html>
<head>
<meta name=HandheldFriendly content=true>
<meta name=PalmComputingPlatform content=true>
<link rel=stylesheet href="mobile-styles.css" type="text/css">
</head>
<body>
<b>LegiTrak</b> <i>Mobile!</i><br>
<b><%=CustomerName%></b><br><br>
<%=strInvalid%>
<form action="mobile-quickadd.asp" method=post>
<input type=submit name=UpdateEntry value=Submit>
<input type=submit name=UpdateEntry value=Cancel>
<br><br>
<table width=153 cellspacing=0 cellpadding=0 border=0>
<col width=50><col width=100>
<tr><td>Client:</td>
<td><select name=ClientID style='width:100'>
<%
' CLIENT LIST
	strSQL = _
		"SELECT [Client List].*" & _
		" FROM [Customer Clients] INNER JOIN [Client List]" & _
		" ON [Customer Clients].ClientID = [Client List].ClientID" & _
		" WHERE [Customer Clients].CustomerID=" & CustomerID & _
		" ORDER BY [Short Company Name]"
	set rsClients=Server.CreateObject("ADOR.Recordset")
	rsClients.Open strSQL, strConnReadOnly
	Do Until rsClients.EOF
		Response.Write "<option value=" & Encrypt(rsClients("ClientID")) & ">"
		Response.Write rsClients("Short Company Name")
		rsClients.MoveNext
	Loop
	rsClients.Close
	set rsClients = Nothing
%>
</select></td></tr>
<tr><td>Bill:</td>
<td><input type=text name=BillNum style='width:100'></td></tr>
<tr><td>Priority:</td>
<td><select name=Pri style='width:100'>
<option value=1>High
<option value=2>Medium
<option value=3>Low
<option value=4>TBD
</select></td></tr>
<tr><td>Position:</td>
<td><select name=Pos style='width:100'>
<option value=1>Support
<option value=2>Oppose
<option value=3>Concerns
<option value=4>Neutral
<option value=5>Monitor
<option value=6>-Blank-
</select></td></tr>
<tr><td colspan=2>Comments:<br>
<textarea name=Com cols=32 rows=10></textarea>
</td></tr>
</table>
</form>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>