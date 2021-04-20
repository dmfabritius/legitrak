<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' RESET CUSTOMER'S HOME PAGE PRIORITY FILTER PREFERENCE
	If Request.Cookies("LegiTrak")("ClearHomePageFilter") = "True" Then
		Set cmdSQL = CreateObject("ADODB.Connection")
		cmdSQL.Open strConnection
		strCommand = _
			"UPDATE [Customer List] " & _
			"SET HomePagePriority=0 " & _
			"WHERE CustomerID=" & CustomerID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		cmdSQL.Close
		Set cmdSQL = Nothing
		Response.Cookies("LegiTrak")("ClearHomePageFilter") = ""
	End If

' SAVE CUSTOMER'S HOME PAGE PRIORITY FILTER PREFERENCE
	If Request.Form("UpdatePriority") = "True" Then
		Set cmdSQL = CreateObject("ADODB.Connection")
		cmdSQL.Open strConnection
		strCommand = _
			"UPDATE [Customer List] SET " & _
			" HomePagePriority=" & CInt(Request.Form("Priority")) & " " & _
			"WHERE CustomerID=" & CustomerID
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		cmdSQL.Close
		Set cmdSQL = Nothing
	End If

' LOAD CUSTOMER'S HOME PAGE PRIORITY FILTER PREFERENCE
	Dim strPri(4)
	Set cmdSQL = CreateObject("ADODB.Connection")
	cmdSQL.Open strConnReadOnly
	strCommand = _
		"SELECT HomePagePriority " & _
		"FROM [Customer List] " & _
		"WHERE CustomerID=" & CustomerID
	Set rsResult = cmdSQL.Execute(strCommand)
	intPri = rsResult("HomePagePriority")
	Response.Cookies("LegiTrak")("FilterPri") = intPri
	strPri(intPri) = " checked"

	rsResult.Close

	Set rsResult = Nothing
	cmdSQL.Close
	Set cmdSQL = Nothing
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>var F</script>
</head>
<body
 onload='parent.tracking.location.href="customer.asp";F=document.getElementById("F1")'
 class=bkg02 style='text-align:center;margin:2 0 0 0'>
<div class=hdg30>Tracking Lists</div>
<span class=det30 style='font:8pt Arial;cursor:default;overflow:hidden;height:25'>
<form id=F1 action="customer-header.asp" method=post>
<input name=UpdatePriority type=hidden value=True>
<input name=Priority type=radio onclick='F.submit()' value=0<%=strPri(0)%>>All Priorities&nbsp;
<input name=Priority type=radio onclick='F.submit()' value=1<%=strPri(1)%>>High&nbsp;
<input name=Priority type=radio onclick='F.submit()' value=2<%=strPri(2)%>>Medium&nbsp;
<input name=Priority type=radio onclick='F.submit()' value=3<%=strPri(3)%>>Low&nbsp;
<input name=Priority type=radio onclick='F.submit()' value=4<%=strPri(4)%>>TBD
</form></span>
</body>
</html>
