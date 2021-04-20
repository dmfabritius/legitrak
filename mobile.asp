<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	' If we're logged in, then we want to log out
	If CustomerID <> 0 Then
		Response.Cookies("LegiTrak")("CustomerID") = ""
		Response.Cookies("LegiTrak")("ClientID") = ""
		strSQL = "SELECT * FROM [Customer List] WHERE [CustomerID] =" & CustomerID
		set rsCustomer=Server.CreateObject("ador.Recordset")
		rsCustomer.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		rsCustomer("In Use") = 0
		rsCustomer.Update
		rsCustomer.Close
		Set rsCustomer = Nothing
	End If
%>
<html>
<head>
<meta name=HandheldFriendly content=true>
<meta name=PalmComputingPlatform content=true>
<link rel=stylesheet href="mobile-styles.css" type="text/css">
</head>
<body onload='document.all.item("SignOn").CustomerUsername.focus()'>
<b>Welcome to LegiTrak</b><br>
<i>Mobile Edition!</i><br>
<form id=SignOn method=post action="mobile-customer.asp">
<table width=153 cellspacing=0 cellpadding=0 border=0>
<tr><td colspan=2>Please enter your account information:<br><br></td></tr>
<tr><td align=right>Sign-On:&nbsp; </td>
<td><input type=text name=CustomerUsername size=20></td></tr>
<tr><td align=right>Password:&nbsp; </td>
<td><input type=password name=CustomerPassword size=20></td></tr>
</table>
<br>
<input type=submit value="Submit">
</form>
<br>
<!--#include virtual="includes/copyright.asp"-->
</body>
</html>