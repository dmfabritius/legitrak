<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Form("UpdateEntry") = "Cancel" Then
		Response.Redirect "mobile-customer.asp"
	End If

    BillNum = CInt("0" & Request.Form("BillNum"))

' UPDATE CLIENT SPECIFIC BILL INFO
	If Request.Form("UpdateEntry") = "Submit" And ClientID <> 0 And _
	    BillNum >= 1000 And BillNum <= 9999 Then

		strDead = "0"
		If Request.Form("Dead") = "True" Then strDead = "1"
		strSQL = _
			"UPDATE [Client Specific Bill Info] SET " & _
			" PriorityNum=" & CInt(Request.Form("Pri")) & "," & _
			" PositionNum=" & CInt(Request.Form("Pos")) & "," & _
			" Dead=" & strDead & "," & _
			" Comments='" & TweakQuote(Request.Form("Com")) & "' " & _
			"WHERE" & _
			" ClientID=" & ClientID & " AND" & _
			" [Bill Number]=" & BillNum
		Set cxnSQL = CreateObject("ADODB.Connection")
		With cxnSQL
		    .Open strConnection
		    .Execute strSQL, , adExecuteNoRecords
		    .Close
		End With
		Set cmdSQL = Nothing
		Response.Redirect "mobile-customer.asp"
	End If

' LOAD CLIENT SPECIFIC BILL INFO FOR EDITING
	BillNum = CInt(Request.QueryString("bill"))
	strSQL = _
		"SELECT * FROM [Client Specific Bill Info]" & _
		" LEFT JOIN [Client List] ON [Client Specific Bill Info].ClientID=[Client List].[ClientID]" & _
		" WHERE [Client Specific Bill Info].ClientID=" & ClientID & " AND [Bill Number]=" & BillNum
	set rsBillInfo=Server.CreateObject("ador.Recordset")
	rsBillInfo.Open strSQL, strConnection
    If Not rsBillInfo.EOF Then
	    ClientName = rsBillInfo("Short Company Name")
	    Dim strPri(4)
	    strPri(rsBillInfo("PriorityNum")) = " selected"
	    Dim strPos(6)
	    strPos(rsBillInfo("PositionNum")) = " selected"
	    If rsBillInfo("Dead") = "True" Then strDead = " checked"
	    strCom = rsBillInfo("Comments")
	    If IsNull(strCom) or Trim(strCom) = "" Then
		    strCom = ""
	    Else
		    strCom = Replace(strCom,"<br>",vbCrLf)
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
<form action="mobile-edit.asp" method=post>
<input type=hidden name=BillNum value=<%=BillNum%>>
<input type=submit name=UpdateEntry value=Submit>
<input type=submit name=UpdateEntry value=Cancel>
<br><br>
<table width=153 cellspacing=0 cellpadding=0 border=0>
<col width=50><col width=100>
<tr><td><b>Client:</b></td><td><%=ClientName%></td></tr>
<tr><td><b>Bill:</b><br><br></td><td><%=BillNum%><br><br></td></tr>
<tr><td>Priority:</td>
<td><select name=Pri style='width:100'>
<option value=1<%=strPri(1)%>>High
<option value=2<%=strPri(2)%>>Medium
<option value=3<%=strPri(3)%>>Low
<option value=4<%=strPri(4)%>>TBD
</select></td></tr>
<tr><td>Position:</td>
<td><select name=Pos style='width:100'>
<option value=1<%=strPos(1)%>>Support
<option value=2<%=strPos(2)%>>Oppose
<option value=3<%=strPos(3)%>>Concerns
<option value=4<%=strPos(4)%>>Neutral
<option value=5<%=strPos(5)%>>Monitor
<option value=6<%=strPos(6)%>>-Blank-
</select></td></tr>
<tr><td>Dead:</td>
<td><input name=Dead type=checkbox value=True <%=strDead%>></td></tr>
<tr><td colspan=2>Comments:<br>
<textarea name=Com cols=32 rows=10><%=strCom%></textarea>
</td></tr>
</table>
</form>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>