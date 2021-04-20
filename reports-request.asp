<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set rsQueue=Server.CreateObject("ADOR.Recordset")

' DETERMINE WHICH REPORT SELECTION SECTION TO DISPLAY
	If Request.Cookies("LegiTrak")("RptSection") = "Clt" Then
		strCltSec = ""
		strCustSec = "none"
		intSection = 1
	Else
		strCltSec = "none"
		strCustSec = ""
		intSection = 0
	End If

' LOAD CLIENT LIST
	If Request.Cookies("LegiTrak")("SessionStatus") = 3 And Request.Cookies("LegiTrak")("SessionOnly") = "True" Then
		strSQL = " AND C.[Year-Round]=1"
	Else
		strSQL = ""
	End If
	strSQL = _
		"SELECT C.ClientID, C.[Short Company Name], 'x' AS EncryptedID, C.[Active Report] " & _
		"FROM [Customer Clients] CC INNER JOIN [Client List] C ON CC.ClientID = C.ClientID " & _
		"WHERE CC.CustomerID=" & CustomerID & strSQL & " ORDER BY C.[Client Company Name]"
	rsQueue.Open strSQL, strConnReadOnly
	aClients = rsQueue.GetRows()
	intClients = UBound(aClients,2)
	rsQueue.Close

' QUEUE UP REQUESTED REPORTS
	If Len(Request.Form("Rpts")) <> 0 Then
	
		' Load Daily ReportID preference
		strSQL = "SELECT [Active Report] FROM [Customer List] WHERE CustomerID=" & CustomerID
		rsQueue.Open strSQL, strConnReadOnly
		DailyID = rsQueue("Active Report")
		rsQueue.Close
	
		strSQL = "SELECT * FROM [Report Queue]"
		rsQueue.Open strSQL, strConnection, adOpenStatic, adLockPessimistic
		aRpts = Split(Request.Form("Rpts"),",")
		For i = 0 to UBound(aRpts)
			params = Split(aRpts(i),"_")
			If CStr(params(1)) = "0" Then
				ClientID = 0
			Else
				ClientID = Decrypt(params(1))
			End If
			If params(0) = 1 Then
				ReportID = DailyID
			ElseIf params(0) = 2 Then
				For j = 0 to intClients
					If aClients(0,j) = ClientID Then ReportID = CInt(aClients(3,j))
				Next 'j
			Else
				ReportID = CInt(params(0))
			End If

			rsQueue.AddNew
			rsQueue("CustomerID") = CustomerID
			rsQueue("ClientID") = ClientID
			rsQueue("ReportID") = ReportID
			rsQueue("Report Status") = "REQUESTED"
			rsQueue("Effective Date") = Now
			rsQueue.Update
		Next 'i
		rsQueue.Close

	End If

' LOAD COUNT OF QUEUED REPORTS
	strSQL = _
		"SELECT COUNT(ReportStatusID) Queued FROM [Report Queue] WHERE" & _
		 " CustomerID=" & CustomerID & " AND" & _
		 " [Report Status]='REQUESTED'"
	rsQueue.Open strSQL, strConnReadOnly
	If rsQueue.EOF Then
		intQueued = 0
	Else
		intQueued = rsQueue("Queued")
	End If
	rsQueue.Close

' SHOW HIDDEN REPORTS FOR ADMIN ACCOUNTS
	If CustomerID = 1 Or CustomerID = 41 Then
		strWhere = ""
	Else
		strWhere = "AND [Hidden]=0 "
	End If

' LOAD CUSTOMER REPORT LIST
	strSQL = _
		"SELECT ReportID, [Report Display Name], Hidden FROM [Reports] " & _
		"WHERE [Report Type]=1 " & strWhere & _
		"ORDER BY [Display Order]"
	rsQueue.Open strSQL, strConnReadOnly
	aCustRpts = rsQueue.GetRows()
	intCustRpts = UBound(aCustRpts,2)
	intCustWidth = 470 + 20*(intCustRpts+1)
	rsQueue.Close

' LOAD CLIENT REPORT LIST
	strSQL = _
		"SELECT ReportID, [Report Display Name], Hidden FROM [Reports] " & _
		"WHERE [Report Type]=2 " & strWhere & _
		"ORDER BY [Display Order]"
	rsQueue.Open strSQL, strConnReadOnly
	aReports = rsQueue.GetRows()
	intReports = UBound(aReports,2)
	intWidth = 370 + 20*(intReports+1)
	rsQueue.Close

' VERIFY E-MAIL ADDRESS EXISTS	
	strSQL = "SELECT Email FROM [Customer List] WHERE CustomerID=" & CustomerID
	rsQueue.Open strSQL, strConnReadOnly
	bolEmail = InStr(rsQueue("Email"),"@") <> 0
	rsQueue.Close

	Set rsQueue = Nothing
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var DC,CR0,CR1,CR2
function init(){
	Cust=document.getElementById("CustSection")
	Clt=document.getElementById("CltSection")
	DC=document.getElementById("divClts")
	CR0=document.getElementById("CustRpts")
	CR1=document.getElementById("CltRpts1")
	CR2=document.getElementById("CltRpts2")

	selectTab(0)
	
	if (!MSIE) colorCols(<%=intSection%>)
}
function colorCols(c){
	s=myStyles[".bkg09"].backgroundColor
	if (c){
		scrollClients()
		m=CR2.rows[0].cells.length-2
	} else {
		l=CR0.rows.length-1
		m=CR0.rows[l].cells.length-2
	}
	for (r=1;r<=m;r+=2){
		if (c) {
			CR1.rows[r-1].cells[1].style.backgroundColor=s
			if (x=CR1.rows[r]) x.cells[0].style.backgroundColor=s
			CR2.rows[0].cells[r].style.backgroundColor=s
		} else {
			CR0.rows[r-1].cells[1].style.backgroundColor=s
			if (r!=l) CR0.rows[r].cells[0].style.backgroundColor=s
			CR0.rows[l].cells[r].style.backgroundColor=s
		}
	}
}
function verifyEmail(){
<%
	If Not bolEmail Then
		Response.Write "alert('" & _
			"Warning!\n\nYou must specify a valid e-mail address on the Account Information page\n" & _
			"before the requested report(s) will be processed.')"
	End If
%>
	ReportRequests.submit()
}
function rHover(e,h,r,c){
	e = (!e) ? event.srcElement : e.target
	if (e.tagName.toLowerCase()=="input") document.getElementById(e.value.split("_")[1]).style.fontWeight=(h)? "bold" : ""
	while(e.tagName!="TD"&&e.parentNode!=null) e=e.parentNode
	if (!e.parentNode) return
	s=(h)? myStyles[".bkg0A"].backgroundColor : ((r%2) ? myStyles[".bkg09"].backgroundColor : myStyles[".bkg04"].backgroundColor)
	if (c) {
		CR0.rows[r-1].cells[1].style.backgroundColor=s
		if (r!=(l=CR0.rows.length-1)) CR0.rows[r].cells[0].style.backgroundColor=s
		CR0.rows[l].cells[r].style.backgroundColor=s
	} else {
		CR1.rows[r-1].cells[1].style.backgroundColor=s
		if (c=CR1.rows[r]) c.cells[0].style.backgroundColor=s
		CR2.rows[0].cells[r].style.backgroundColor=s
	}
}
function scrollClients(){
	DC.style.width=parseInt(CR2.width)+17
	DC.style.height=document.body.clientHeight-200
	if (!MSIE){
		c=CR2.rows[0].cells[0]
		for(i=0;i<c.childNodes.length;i++) c.childNodes[i].style.height="10px"
	}
}
function showRpt(r){
	if (r) {
		Cust.style.display="none"
		Clt.style.display=""
		setCookie("RptSection","Clt")
	} else {
		Cust.style.display=""
		Clt.style.display="none"
		setCookie("RptSection","Cust")
	}
	colorCols(r)
}
</script>
</head>
<body class=bkg04 style='margin:20 0 0 20' onload='init()' onresize='scrollClients()'>
<form method=post action="reports-request.asp">
<%
'
'
' LIST OF CUSTOMER REPORTS
' ------------------------
'
	Response.Write _
		"<div id=CustSection style='display:" & strCustSec & "'>" & _
		"<table id=CustRpts width=" & intCustWidth & " border=0 cellspacing=0 cellpadding=0 class=det00>"

	Response.Write "<col width=170>"
	strBGColor = " class=bkg09"
	For i = 0 to intCustRpts
		Response.Write "<col width=20" & strBGColor & ">"
		If strBGColor = "" Then
			strBGColor = " class=bkg09"
		Else
			strBGColor = ""
		End If
	Next
	Response.Write "<col width=300>"

	For i = 0 to intCustRpts
		If aCustRpts(2,i) Then
			str1 = "<span style='color:purple'><b>"
			str2 = "</b></span>"
		Else
			str1 = ""
			str2 = ""
		End If
		
		j = intCustRpts-i+1
		Response.Write "<tr><td rowspan=" & j
		If i = 0 Then
			Response.Write _
				" valign=top>" & _
 				"<div class=lnk70 onclick='window.location.href=""reports-requested.asp""'><b>Reports Requested: " & intQueued & "</b></div><br>" & _
				"<div class=hdg24>Summary Reports</div>" & _
				"<div class=lnk70 onclick='showRpt(1)'>Detail Reports</div><br>" & _
				"<input type=submit value=Submit> &nbsp; <input type=reset value=Cancel>"
			strRptName = "<div class=lnk70 onclick='window.location.href=""reports-customize.asp?0""'>" & aCustRpts(1,i) & "</div>"
		ElseIf aCustRpts(0,i) = 29 Then
			Response.Write ">&nbsp;"
			strRptName = "<div class=lnk70 onclick='window.location.href=""reports-customize.asp?29""'>" & aCustRpts(1,i) & "</div>"
		Else
			Response.Write ">&nbsp;"
			strRptName = aCustRpts(1,i)
		End If
		Response.Write _
			"</td><td colspan=" & j+1 & " width=" & 20*j+300 & ">" & _
			str1 & strRptName & str2 & _
			"</td></tr>"
	Next 'i

	Response.Write "<tr valign=bottom height=40><td id=0>All Tracking Lists</td>"
	For j = 0 to intCustRpts
		Response.Write _
			"<td onmouseover='rHover(arguments[0],1," & j+1 & ",1)' onmouseout='rHover(arguments[0],0," & j+1 & ",1)'>" & _
			"<input type=checkbox" & " name=Rpts value='"  & aCustRpts(0,j) & "_0'>"
		Response.Write "</td>"
	Next 'j
	Response.Write "<td>&nbsp;</td></tr>"

	Response.Write "</table></div>"
'
'
' LIST OF CLIENT REPORTS
' ----------------------
'
	Response.Write _
		"<div id=CltSection style='display:" & strCltSec & "'>" & _
		"<table id=CltRpts1 width=" & intWidth & " border=0 cellspacing=0 cellpadding=0 class=det00>"

	Response.Write "<col width=170>"
	strBGColor = " class=bkg09"
	For i = 0 to intReports
		Response.Write "<col width=20" & strBGColor & ">"
		If strBGColor = "" Then
			strBGColor = " class=bkg09"
		Else
			strBGColor = ""
		End If
	Next
	Response.Write "<col width=200>"

	For i = 0 to intReports
		If aReports(2,i) Then
			str1 = "<span style='color:purple'><b>"
			str2 = "</b></span>"
		Else
			str1 = ""
			str2 = ""
		End If
		j = intReports-i+1
		Response.Write "<tr><td rowspan=" & j
		If i = 0 Then
			Response.Write _
				" valign=top>" & _
 				"<div class=lnk70 onclick='window.location.href=""reports-requested.asp""'><b>Reports Requested: " & intQueued & "</b></div><br>" & _
				"<div class=lnk70 onclick='showRpt(0)'>Summary Reports</div>" & _
				"<div class=hdg24>Detail Reports</div><br>" & _
				"<input type=submit value=Submit> &nbsp; <input type=reset value=Cancel>"
			strRptName = "<div class=lnk70 onclick='window.location.href=""reports-customize.asp?" & Encrypt(aClients(0,i)) & """'>" & aReports(1,i) & "</div>"
		ElseIf aReports(0,i) = 14 Then
			Response.Write ">&nbsp;"
			strRptName = "<div class=lnk70 onclick='window.location.href=""reports-customize.asp?14""'>" & aReports(1,i) & "</div>"
		Else
			Response.Write ">&nbsp;"
			strRptName = aReports(1,i)
		End If
		Response.Write _
			"</td><td colspan=" & j & " width=" & 20*j+200 & ">" & _
			str1 & strRptName & str2 & _
			"</td></tr>"
	Next 'i
	Response.Write "</table>"
'
'
' LIST OF CLIENTS AND REPORT SELECTION CHECKBOXES
' -----------------------------------------------
'
	Response.Write _
		"<div id=divClts style='overflow:auto;height:1000;width:587'>"

	Response.Write _
		"<table id=CltRpts2 width=" & intWidth & " border=0 cellspacing=0 cellpadding=0" & _
		" class=det00>" ' style='padding:0 2 0 2' 

	Response.Write "<col width=170>"
	strBGColor = " class=bkg09"
	For i = 0 to intReports
		Response.Write "<col width=10" & strBGColor & ">"
		If strBGColor = "" Then
			strBGColor = " class=bkg09"
		Else
			strBGColor = ""
		End If
	Next
	Response.Write "<col width=200>"

	Response.Write "<tr><td>"
	For i = 0 to intClients
		aClients(2,i) = Encrypt(aClients(0,i))
		Response.Write _
			"<div id=" & aClients(2,i) & " class=div04 style='width:150px'>" & aClients(1,i) & "</div>"
	Next 'i
	Response.Write "</td>"
	
	For j = 0 to intReports
		Response.Write "<td onmouseover='rHover(arguments[0],1," & j+1 & ")' onmouseout='rHover(arguments[0],0," & j+1 & ")'>"
		For i = 0 to intClients
			Response.Write _
				"<input type=checkbox" & _
				" name=Rpts value='"  & aReports(0,j) & _
				"_" & aClients(2,i) & "'><br>"
		Next 'i
		Response.Write "</td>"
	Next 'j
	Response.Write "<td>&nbsp;</td></tr>"

	Response.Write "</table></div>"
%>
</div>
</form>
</body>
</html>