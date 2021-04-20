<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection
	strRedirect = ""

' UPDATE CLIENT ACCOUNT INFORMATION
	If Request.Form("UpdateAcctInfo") = "True" Then

	' Check to see if the client already exists
        strClientName = TweakQuote(Request.Form("ClientName"))
		strCommand = _
			"SELECT ClientID FROM [Client List] " & _
			"WHERE" & _
				" [Client Company Name] = '" & strClientName & "' AND " & _
				" [CustomerID]=" & CustomerID
		Set rsResult = cxnSQL.Execute(strCommand)
		If rsResult.EOF Then
			strCommand = _
				"INSERT INTO [Client List] (" & _
				" [CustomerID], [Client Company Name], [Short Company Name]," & _
				" [Contact First Name], [Contact Last Name], [Address Line 1]," & _
				" [City], [State], [Zip]," & _
				" [Email], [Modified Date], [Active Report]) " & _
				"VALUES (" & CustomerID & "," & _
				"'" & strClientName & "'," & _
				"'" & TweakQuote(Trim(Request.Form("ClientShort"))) & "'," & _
				"'" & TweakQuote(Request.Form("FirstName")) & "'," & _
				"'" & TweakQuote(Request.Form("LastName")) & "'," & _
				"'" & TweakQuote(Request.Form("Addr1")) & "'," & _
				"'" & TweakQuote(Request.Form("City")) & "'," & _
				"'" & Replace(Request.Form("State"),"'","") & "'," & _
				"'" & Replace(Request.Form("Zip"),"'","") & "'," & _
				"'" & TweakQuote(Request.Form("Email")) & "'," & _
				"'" & Now & "'," & _
                        "2)"
			cxnSQL.Execute strCommand, , adExecuteNoRecords

		' Look up newly created ClientID
			strCommand = "SELECT MAX(ClientID) AS ClientID FROM [Client List]"
			rsResult = cxnSQL.Execute(strCommand)
			ClientID = rsResult("ClientID")

		' Add client ownership record
			strCommand = "INSERT INTO [Customer Clients] VALUES (" & CustomerID & "," & ClientID & ")"
			cxnSQL.Execute strCommand, , adExecuteNoRecords

		' Add client lobbyist sharing records
			If Request.Form("CustCount") <> 0 Then
				For i = 1 to Request.Form("CustCount")
				    ID=Decrypt(Request.Form("c" & i))
					If ID <> 0 Then
						strCommand = "INSERT INTO [Customer Clients] VALUES (" & ID & "," & ClientID & ")"
						cxnSQL.Execute strCommand, , adExecuteNoRecords
					End If
				Next ' i
			End If
			strRedirect = _
				"top.contents.location.href='contents.asp';" & _
				"window.location.href='customer-info.asp'"
			
	' Client already exists
		Else
			Response.Cookies("LegiTrak")("ClientID") = Encrypt(rsResult("ClientID"))
			strRedirect = "top.details.location.href='client.htm'"
		End If
	End If

' DETERMINE AUTHORIZATION TO ADD A CLIENT
	strCommand = _
		"SELECT O.[Org Type], O.[Billing Clients], COUNT(DISTINCT CC.ClientID) AS [Actual Clients], C.CreateLists " & _
		"FROM [Organization List] O INNER JOIN" & _
		" [Customer List] C ON O.OrganizationID = C.OrganizationID INNER JOIN" & _
		" [Customer List] OC ON C.OrganizationID = OC.OrganizationID INNER JOIN" & _
		" [Customer Clients] CC ON OC.CustomerID = CC.CustomerID " & _
		"GROUP BY O.[Org Type], O.[Billing Clients], C.CustomerID, C.CreateLists " & _
		"HAVING C.CustomerID=" & CustomerID
	Set rsResult = cxnSQL.Execute(strCommand)
	If Not rsResult.EOF Then
		intOrgType = rsResult("Org Type")
		intNumClients = rsResult("Actual Clients")
		If Not rsResult("CreateLists") Then
			intMaxClients = 0
			intNumClients = 1
		ElseIf intOrgType = 0 Then
			If rsResult("Billing Clients") = 1 Then
				intMaxClients = 2
			Else
				intMaxClients = 50 ' unlimited
			End If
		Else
			If rsResult("Billing Clients") < 4 Then
				intMaxClients = rsResult("Billing Clients")
			Else
				intMaxClients = 50 ' unlimited
			End If
		End If
	Else
		intOrgType = 0
		intMaxClients = 1
		intNumClients = 0
	End If
	bolCanAddClient = intNumClients < intMaxClients

	cxnSQL.Close
	Set cxnSQL = Nothing
%>
<html>
<head>
<link rel="stylesheet" href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var info
function init(){
	info=document.getElementById("Info")
<%
	If strRedirect <> "" Then
		Response.Write strRedirect
	Else
		Response.Write "selectTab(1);"
		If bolCanAddClient Then Response.Write "info.ClientName.focus()"
	End If
%>}
function validateForm() {
	c=info.ClientName
	if (c.value.length==0) {
		alert("Please enter a Client Company Name.")
		c.style.backgroundColor = "#FFFFFF"
		c.focus()
		return false
	}
	return true
}
function mark2(){
	info.ClientName.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	info.ClientShort.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	info.ClientShort.value=info.ClientName.value
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:10 20'>
<%
	If strRedirect <> "" Then Response.End
	If bolCanAddClient Then
%>
<form id=Info action="customer-addclient.asp" method=post onsubmit='return validateForm()'>
<input type=hidden name=UpdateAcctInfo value=True>
<table border=0 cellpadding=0 cellspacing=1 class=det00 style='padding-left:10'>
<col width=150 align=right class=hdg24>
<tr><td align=right>List Name:</td><td><input type=text style='width:325' onchange='mark2()' name=ClientName></td></tr>
<tr><td align=right>Short Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=ClientShort></td></tr>
<tr><td align=right>Contact First Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=FirstName></td></tr>
<tr><td align=right>Contact Last Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=LastName></td></tr>
<tr><td align=right>Email:</td><td><input type=text style='width:325' onchange='return isEmail(this)' name=Email></td></tr>
<tr><td align=right>Address:</td><td><input type=text style='width:325' onchange='mark(this)' name=Addr1></td></tr>
<tr><td align=right></td><td>
<input type=text style='width:200' onchange='mark(this)' name=City>
<input type=text style='width:30' onchange='mark(this)' name=State maxlength=2>
<input type=text style='width:87' onchange='mark(this)' name=Zip maxlength=10></td></tr>
<!-- align=center-->
<tr><td></td><td valign=bottom style='height:30'><input type=submit value=Submit></td></tr>
<%
' CUSTOMER LIST
	strSQL = "SELECT OrganizationID FROM [Customer List] WHERE CustomerID=" & CustomerID
	SQLJoin = "([Customer List] C INNER JOIN (" & strSQL & ") O ON C.OrganizationID=O.OrganizationID) "
	strSQL = "SELECT" & _
		" C.CustomerID, C.[Contact First Name], C.[Contact Last Name] " & _
		"FROM " & SQLJoin & _
		"WHERE CustomerID <> " & CustomerID & " ORDER BY C.[Contact First Name]"

	Set rsCustomers=Server.CreateObject("ADOR.Recordset")
	rsCustomers.Open strSQL, strConnReadOnly
	If Not rsCustomers.EOF Then
		aCustomers = rsCustomers.GetRows()
		intCustCount = UBound(aCustomers,2)
		Response.Write _
			"<tr valign=top><td><input type=hidden name=CustCount value=" & intCustCount & ">" & _
			"<br>Share list with:</td><td><br>"
		For i = 0 to intCustCount
			Response.Write _
				"<input type=checkbox name=c" & i & _
				" value=" & Encrypt(aCustomers(0,i)) & "> &nbsp;" & _
				aCustomers(1,i) & " " & aCustomers(2,i) & "<br>"
		Next 'i
		Response.Write "</td></tr>"
	Else
		Response.Write _
			"<tr><td><input type=hidden name=CustCount value=-1>"
	End If
	rsCustomers.Close
	Set rsCustomers = Nothing
%>
</table>
</form>
<%
' NEW CLIENT CREATION AUTHORIZATION DENIED
	Else
		Response.Write _
			"<div class=shd24><br><br>" & _
			"You are not currently authorized to create additional tracking lists.</div>"
	End If
%>
</body>
</html>
