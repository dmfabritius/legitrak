<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	If Request.Cookies("LegiTrak")("SessionStatus") = 3 And Request.Cookies("LegiTrak")("SessionOnly") = "True" Then
		strSQL = " AND C.[Year-Round]=1"
	Else
		strSQL = ""
	End If

	If Request.Cookies("LegiTrak")("FilterPri") <> "" Then
		intPri = Request.Cookies("LegiTrak")("FilterPri")
		If intPri <> 0 Then strSQL = strSQL & " AND CS.PriorityNum=" & intPri
		strSQLJoin = "[Client Specific Bill Info] CS "
		strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Client List] C ON CS.[ClientID]=C.[ClientID])"
		strSQLJoin = "(" & strSQLJoin & "INNER JOIN [Customer Clients] CC ON C.[ClientID]=CC.[ClientID])"
		strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Daily Status] D ON CS.[Bill Number]=D.[Bill Number]) "
		strSQL = _
			"SELECT" & _
			" C.ClientID, C.[Client Company Name]," & _
			" CS.[Bill Number], CS.[PositionNum]," & _
			" D.Title, D.House, D.Location, D.Action " & _
			"FROM " & strSQLJoin & _
			"WHERE CC.CustomerID=" & CustomerID & strSQL & _
			" ORDER BY C.[Client Company Name], CS.[Bill Number]"
		Set rsBillInfo=Server.CreateObject("ADOR.Recordset")
		rsBillInfo.Open strSQL, strConnReadOnly
		If rsBillInfo.EOF Then
			If intPri <> 0 Then
				Response.Cookies("LegiTrak")("ClearHomePageFilter") = "True"
				strRedirect = "onload='parent.trackingheader.location.href=""customer-header.asp""'"
			Else
				Response.Cookies("LegiTrak")("ClientID") = ""
				strRedirect = "onload='init()'"
			End If
		End If
	Else
		strRedirect = "onload='parent.trackingheader.location.href=""customer-header.asp""'"
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>table,tr,td{cursor:pointer}</style>
<script src="js/bts.js"></script>
<script>
function init(){<%=strReload%>
	if (top.contents.document.getElementById("CltCount").innerHTML==0)
		selectMenu("mnu5","customer-info.htm")
	else
		menuSelect(null,null,0)
}
function billHover(e,h){
	e = (!e) ? event.srcElement : e.target
	while (e.tagName!="TR") e=e.parentNode
	if (h) {
		e.style.saveBkg=e.style.backgroundColor
		e.style.backgroundColor=myStyles[".bkg0B"].backgroundColor
	} else
		e.style.backgroundColor=e.style.saveBkg
}
function goToClient(e){
	e = (!e) ? event.srcElement : e.target
	while (e.tagName!="TR") e=e.parentNode
	if (!/x/.test(e.id)){
		if (getCookie("SessionStatus")!=3) setCookie("ClientBill",e.cells[0].innerHTML)
		bt=document.getElementById("Bills")
		for (i=e.id;true;i--) if (/x/.test(bt.rows[i].id)) break
		e=bt.rows[i]
	}
	name=e.cells[0].innerHTML.fromHTML().trim()
	c=top.contents.document.getElementById("ClientMenu")
	for (var i=0;i<c.rows.length;i+=2) if (c.rows[i].cells[0].childNodes[0].childNodes[0].title==name) break
	menuSelect(null,null,i)
}
</script>
</head>
<body <%=strRedirect%> class=bkg04 style='margin:0 0 0 3;text-align:center'>
<%
	If strRedirect <> "" Then Response.End

	Response.Write _
		"<table id=Bills width=350 border=0 cellspacing=0 cellpadding=0 class=det00" & _
		" onclick='goToClient(arguments[0])' onmouseover='billHover(arguments[0],1)' onmouseout='billHover(arguments[0],0)'>" & _
		"<col width=35><col width=195><col width=120>"

	prevClt = 0
	r = 0
	Do Until rsBillInfo.EOF
		If Trim(rsBillInfo("Title")) <> "" Then
			strTitle = rsBillInfo("Title")
		Else
			strTitle = "(No Title Available)"
		End If
		If Trim(rsBillInfo("House")) <> "" Then
			strHouseLoc = rsBillInfo("House") & ", " & rsBillInfo("Location")
		ElseIf Trim(rsBillInfo("Location")) <> "" Then
			strHouseLoc = rsBillInfo("Location")
		Else
			strHouseLoc = ""
		End If
		If Trim(rsBillInfo("Action")) <> "" Then
			strBackground = " bgcolor=#E7E7E7" ' yellow=FFFFA0
		Else
			strBackground = ""
		End If
		strStyle = " style='font-weight:bold;color:"
		Select Case rsBillInfo("PositionNum")
			Case 1: strStyle = strStyle & "#009000'"
			Case 2: strStyle = strStyle & "red'"
			Case 3: strStyle = strStyle & "orange'"
			Case Else
				strStyle = ""
		End Select
		cltID = rsBillInfo("ClientID")
		If cltID <> prevClt Then
			Response.Write _
				"<tr id=x" & Encrypt(cltID) & "><td colspan=3 class=hdg24><br>" & rsBillInfo("Client Company Name") & "</td></tr>"
			prevClt = cltID
			r=r+1
		End If
		Response.Write _
			"<tr" & strBackground & strStyle & " id=" & r & "><td>" & _
			rsBillInfo("Bill Number") & "</td><td>" & _
			strTitle & "</td><td>" & _
			strHouseLoc & "</td></tr>" & vbCrLf
		r=r+1
		rsBillInfo.MoveNext
	Loop
	Response.Write "</table>"
	rsBillInfo.Close
	Set rsBillInfo = Nothing
%>    
</body>
</html>
