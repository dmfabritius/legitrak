<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	strSQL = "SELECT [Body] FROM [Supplements with Unique Bill Numbers] WHERE [Bill Number]=" & CInt(Request.QueryString("bill"))
	Set rsDig=Server.CreateObject("ADOR.Recordset")
	With rsDig
		.Open strSQL, strConnReadOnly
		If Not .EOF Then
			Response.Cookies("LegiTrak")("digest") = rsDig("Body")
		Else
			Response.Cookies("LegiTrak")("digest") = "(Digest not yet available.)"
		End If
		.Close
	End With
	Set rsDig = Nothing
%>
<html>
</html>
