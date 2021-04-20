<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	bolSystem = ((CustomerID = 1) Or (CustomerID = 267))
'	bolSystem = (CustomerID = 1)
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function init(){
	if ((m=getCookie("menuItem"))=="") setCookie("menuItem",(m="mnu1"))
	if (/mnu/.test(m)){
		o=document.getElementById(m)
		o.className="mnu61"
		o.style.color=myStyles[".mnu61"].color
		document.getElementById(m+"b").src="img/m1.gif"
	}
	if ("<%=bolSystem%>"=="True") top.document.getElementById("nav").rows="220,*"
}
function logout(){
	setCookie("Logout",1)
	top.menu.location.href="menu-logged-out.htm"
	top.contents.location.href="blank.htm"
	top.details.location.href="signon.asp"
	top.document.getElementById("nav").rows="*,0"
}
function help(h){
	if (!(o=top.details.data)) o=top.details
	u=o.location.href
	u=u.substring(u.lastIndexOf("/")+1,u.lastIndexOf("."))
	
	helpTopics=
		"customer"+
		"client-tracking"+
		"allclts-browse"
//	if (helpTopics.indexOf(u)!=-1) h="help/"+u+"-help.htm"
	
	if (!h)
		alert("This page does not currently have an associated help topic.\n\nPlease use the 'Contact Us' option to request assistance.")
	else
		window.open(h,"help")
}
</script>
</head>
<body onload='init()' class=grd5a style='margin-top:8'>
<table id=MainMenu width=111 border=0 cellspacing=0 cellpadding=0>
<col width=105><col width=6>

<tr><td id=mnu1 class=mnu29 onclick='menuSelect(this,"customer.htm")' onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'>
Home</td><td><img id=mnu1b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu2 class=mnu29 onclick='menuSelect(this,"reports.htm")' onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'>
Reports</td><td><img id=mnu2b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu3 class=mnu29 onclick='menuSelect(this,"allclts.htm")' onmouseover='menuHover(this,1)' onmouseout='menuHover(this,0)'>
All Lists</td><td><img id=mnu3b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu4 class=mnu29 onclick='menuSelect(this,"votecard-summary.asp")' onmouseover='menuHover(this,1)' onmouseout='menuHover(this,0)'>
Vote Cards</td><td><img id=mnu4b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu5 class=mnu29 onclick='menuSelect(this,"customer-info.htm")' onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'>
Account Info</td><td><img id=mnu5b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu6 class=mnu29 onclick='help()' onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'>
Help</td><td><img id=mnu6b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu7 class=mnu29 onclick='window.location.href="mailto:Brad Tower <bhtower@towerltd.org>?Subject=LegiTrak Web Site"'
 onMouseOver='menuHover(this,1)' onMouseOut='menuHover(this,0)'>
Contact Us</td><td><img id=mnu7b src='img/m0.gif' /></td></tr>
<tr style='height:4'><td></td></tr>

<tr><td id=mnu8 class=mnu29 onclick='logout()' onmouseover='menuHover(this,1)' onmouseout='menuHover(this,0)'>
Logout</td><td><img id=mnu8b src='img/m0.gif' /></td></tr>
<%
	If bolSystem Then
%>
<tr style='height:10'><td></td></tr>
<tr><td id=mnu9 class=mnu29 onclick='menuSelect(this,"maint.htm")' onmouseover='menuHover(this,1)' onmouseout='menuHover(this,0)'>
Sys Maint</td><td><img id=mnu9b src='img/m0.gif' /></td></tr>
<%
	End If
%>
</table>
</body>
</html>
