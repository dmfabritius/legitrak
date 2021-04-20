<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' LOAD CLIENT LIST
	If Request.Cookies("LegiTrak")("SessionStatus") = 3 And Request.Cookies("LegiTrak")("SessionOnly") = "True" Then
		strSQL = " AND CL.[Year-Round]=1"
	Else
		strSQL = ""
	End If
	strSQL = _
		"SELECT CL.*" & _
		" FROM [Customer Clients] CC INNER JOIN [Client List] CL ON CC.ClientID=CL.ClientID" & _
		" WHERE CC.CustomerID=" & CustomerID & strSQL & _
		" ORDER BY CL.[Short Company Name]"

	Set rsNavClts=Server.CreateObject("ADOR.Recordset")
	rsNavClts.Open strSQL, strConnReadOnly
	If rsNavClts.EOF Then strRedirect = "selectMenu('mnu5','customer-info.htm')"

	If Not rsNavClts.EOF Then
		If ClientID = 0 Then
			ClientID = rsNavClts("ClientID")
			Response.Cookies("LegiTrak")("ClientID") = Encrypt(rsNavClts("ClientID"))
		End If
	End If

%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function init(){<%=strRedirect%>
	if ((m=getCookie("menuItem"))=="") return
	if (/clt/.test(m)){
		o=document.getElementById(m)
		o.className="mnu61"
		o.style.color=myStyles[".mnu61"].color
		document.getElementById(m+"b").src="img/m1.gif"
	}
}
</script>
</head>
<body onload='init()' class=grd5b>
<table id=ClientMenu border=0 cellspacing=0 cellpadding=0>
<%
	i = 0
	Do Until rsNavClts.EOF
		navCltID = EnCrypt(rsNavClts("ClientID"))
		Response.Write _
			"<tr><td id=clt" & i & " class=mnu29 onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'" & _
			" onclick='menuSelect(this)'>"
		Response.Write _
			"<div style='width:95;height:17;overflow:hidden'>" & _
			"<span id=" & navCltID & " style='width:1000' title='" & Trim(rsNavClts("Client Company Name")) & "'>" & _
			rsNavClts("Short Company Name") & _
			"</span></div>"
		Response.Write _
			"</td><td><img id=clt" & i & "b src='img/m0.gif' /></td></tr>" & _
			"<tr style='height:4'><td></td></tr>"

		rsNavClts.MoveNext
		i = i + 1
	Loop
	Response.Write "<tr style='display:none'><td id=CltCount>" & i & "</td></tr>"

	rsNavClts.Close
	set rsNavClts = Nothing
%>
</table>
</body>
</html>