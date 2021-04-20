<%
	' Check to see if we're logged in and the Customer ID was decrypted successfully
	If CustomerID = 0 Then
		Response.Redirect "errors/403-17.htm"
	End If
%>
