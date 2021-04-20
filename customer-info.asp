<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection
	strRedirect = ""

' DELETE ALL CLIENT ACCOUNTS
	If Request.Cookies("LegiTrak")("DeleteClients") = "True" Then
		strCommand = "DELETE FROM [Client List] WHERE CustomerID=" & CustomerID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteClients") = ""
		Response.Cookies("LegiTrak")("ClientID") = ""
		strRedirect = "deletedClients()"

' DELETE CLIENTS' TRACKING LIST ENTRIES
	ElseIf Request.Cookies("LegiTrak")("DeleteEntries") = "True" Then
		strCommand = _
			"DELETE CS " & _
			"FROM [Client Specific Bill Info] CS INNER JOIN [Client List] CL ON CS.ClientID=CL.ClientID " & _
			"WHERE CL.CustomerID=" & CustomerID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteEntries") = ""
'		strRedirect = _
'			"top.contents.location.href='contents.asp';" & _
'			"window.location.href=window.location.href"

' DELETE CLIENTS' CANDIDATE RECOMMENDATIONS
	ElseIf Request.Cookies("LegiTrak")("DeleteRecs") = "True" Then
		strCommand = _
			"UPDATE [Client Politician Comments] SET " & _
			" [Group]=NULL," & _
			" [Primary Rec]=NULL," & _
			" [General Rec]=NULL " & _
			"WHERE ClientID IN (SELECT ClientID FROM [Client List] WHERE CustomerID=" & CustomerID & ")"
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteRecs") = ""
	End If
	cxnSQL.Close
	Set cxnSQL = Nothing

' LOAD CUSTOMER ACCOUNT INFORMATION
	strCommand = "SELECT * FROM [Customer List] WHERE CustomerID=" & CustomerID
	Set rsCustomer=Server.CreateObject("ADOR.Recordset")
	rsCustomer.Open strCommand, strConnection, adOpenDynamic, adLockPessimistic

	' If we just updated the basic customer info, then write it to the database
	If Request.Form("UpdateAcctInfo") = "True" Then
		rsCustomer("Contact First Name") = TweakQuote(Request.Form("FirstName"))
		rsCustomer("Contact Last Name") = TweakQuote(Request.Form("LastName"))
		rsCustomer("Address Line 1") = TweakQuote(Request.Form("Addr1"))
		rsCustomer("Address Line 2") = TweakQuote(Request.Form("Addr2"))
		rsCustomer("City") = TweakQuote(Request.Form("City"))
		rsCustomer("State") = Replace(Request.Form("State"),"'","")
		rsCustomer("Zip") = Replace(Request.Form("Zip"),"'","")
		rsCustomer("Email") = Replace(TweakQuote(Request.Form("Email")),":",";")
		rsCustomer("TextMsg Address") = TweakQuote(Request.Form("TxtMsg"))
		If Request.Form("Notify") = "True" Then
		    strNotify = "1"
		Else
		    strNotify = NULL
		End If
		rsCustomer("TextMsg Notify") = strNotify
		strUser = TweakQuote(Trim(Request.Form("Username")))
		If strUser <> "" Then rsCustomer("Username") = strUser
		strPass = TweakQuote(Trim(Request.Form("Password")))
		If strPass <> "" Then rsCustomer("Password") = strPass
		rsCustomer("DefPriority") = CInt(Request.Form("Priority"))
		rsCustomer("DefPosition") = CInt(Request.Form("Position"))
		rsCustomer("Modified Date") = Now
		rsCustomer.Update
		rsCustomer.MoveFirst

		Response.Cookies("LegiTrak")("DefPriority") = rsCustomer("DefPriority")
		Response.Cookies("LegiTrak")("DefPosition") = rsCustomer("DefPosition")

	End If

	Dim Pri(4), Pos(6)
	Pri(rsCustomer("DefPriority")) = " selected"
	Pos(rsCustomer("DefPosition")) = " selected"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var info
function init(){
	info=document.getElementById("Info")
<%
	If strRedirect <> "" Then
		Response.Write strRedirect
	Else
		Response.Write "selectTab(0);info.FirstName.focus()"
	End If
%>}
function deletedClients(){
	top.contents.location.href="contents.asp"
	sh=top.subheading.document
	sh.getElementById('cltName').innerHTML=""
	sh.getElementById('sep').innerHTML=""
	selectMenu("mnu5","customer-info.htm")
}
function validateForm() {
    if (info.Username.value.trim()==""){
		alert("Username cannot be blank.")
		info.Username.value=""
        return false
    }
    if (info.Password.value.trim()==""){
		alert("Password cannot be blank.")
		info.Password.value=""
		info.Password2.value=""
        return false
    }
	if (info.Password.value!=info.Password2.value) {
		alert("The passwords you entered do not match.  Please try again.")
		info.Password.value=""
		info.Password2.value=""
		return false
	}
	return true
}
function deleteEntries(){
	if (confirm("Do you want to delete the entries from all of your tracking lists?"))
		if (confirm(
				"Deleting the tracking list entries will permanently erase all of the bills in all of your lists.\n\n"+
				"Are you absolutely sure you want to do this?")) {
			setCookie("DeleteEntries","True")
			window.location.href=window.location.href
		}
}
function deleteClients(){
	if (confirm("Do you want to delete all of your tracking lists?"))
		if (confirm(
				"Deleting your tracking lists will permanently erase all associated information, including tracking list entries,\n"+
				"report comments, keywords, and candidate contributions.\n\n"+
				"Are you absolutely sure you want to delete all of your tracking lists?")) {
			setCookie("DeleteClients","True")
			window.location.href=window.location.href
		}
}
function deleteRecs(){
	if (confirm("Do you want to delete the candidate contribution recommendations for all of your tracking lists?"))
		if (confirm(
				"This will permanently erase all of the candidate contribution recommendations for all of your lists.\n\n"+
				"Are you absolutely sure you want to do this?")) {
			setCookie("DeleteRecs","True")
			window.location.href=window.location.href
		}
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:10 20 0 20'>
<%
	If strRedirect <> "" Then Response.End
%>

<form id=Info method=post action="customer-info.asp" onsubmit='return validateForm()'>
<input type=hidden name=UpdateAcctInfo value=True>
<table border=0 cellpadding=0 cellspacing=1 class=det00 style='padding-left:10'>
<col width=150 align=right class=hdg24>
<tr><td>Contact First Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=FirstName value='<%=TweakQuote(rsCustomer("Contact First Name"))%>'></td></tr>
<tr><td>Contact Last Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=LastName value='<%=TweakQuote(rsCustomer("Contact Last Name"))%>'></td></tr>
<tr><td>Email:</td><td><input type=text style='width:325' onchange='return isEmail(this,1)' name=Email value='<%=TweakQuote(rsCustomer("Email"))%>'></td></tr>
<tr><td><a class=hdg24 style="cursor:help" target=_blank href="txtmsg.htm">Text Msg Address:</a></td><td><input type=text style='width:325' onchange='return isEmail(this,1)' name=TxtMsg value='<%=TweakQuote(rsCustomer("TextMsg Address"))%>'></td></tr>
<tr><td></td><td class=shd24><input type=checkbox name=Notify value='True' <% If rsCustomer("TextMsg Notify")=1 Then Response.Write "checked"%>>Send meeting agenda change notices</td></tr>
<tr><td>Address:</td><td><input type=text style='width:325' onchange='mark(this)' name=Addr1 value='<%=TweakQuote(rsCustomer("Address Line 1"))%>'></td></tr>
<tr><td></td><td>
<input type=text style='width:200' onchange='mark(this)' name=City value='<%=TweakQuote(rsCustomer("City"))%>'>
<input type=text style='width:30' onchange='mark(this)' name=State value='<%=Replace(rsCustomer("State"),"'","")%>' maxlength=2>
<input type=text style='width:87' onchange='mark(this)' name=Zip value='<%=Replace(rsCustomer("Zip"),"'","")%>' maxlength=10></td></tr>
<tr><td colspan=2>&nbsp;</td></tr>
<tr><td>Sign-On Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=Username value='<%=TweakQuote(rsCustomer("Username"))%>'></td></tr>
<tr><td>Password:</td><td><input type=password style='width:325' onchange='mark(this)' name=Password value='<%=TweakQuote(rsCustomer("Password"))%>'></td></tr>
<tr><td>Confirm Password:</td><td><input type=password style='width:325' onchange='mark(this)' name=Password2 value='<%=TweakQuote(rsCustomer("Password"))%>'></td></tr>
<tr><td colspan=2>&nbsp;</td></tr>
<tr><td>Default Priority:</td><td>
<select name=Priority style='width:95' onchange='mark(this)'>
  <option value=1<%=Pri(1)%>>High
  <option value=2<%=Pri(2)%>>Medium
  <option value=3<%=Pri(3)%>>Low
  <option value=4<%=Pri(4)%>>TBD
</select></td></tr>
<tr><td>Default Position:</td><td>
<select name=Position style='width:95' onchange='mark(this)'>
  <option value=1<%=Pos(1)%>>Support
  <option value=2<%=Pos(2)%>>Oppose
  <option value=3<%=Pos(3)%>>Concerns
  <option value=4<%=Pos(4)%>>Neutral
  <option value=5<%=Pos(5)%>>Monitor
  <option value=6<%=Pos(6)%>>-Blank-
</select></td></tr>

<%
	rsCustomer.Close
	Set rsCustomer = Nothing
%>
<tr class=shd24><td></td><td><br><input type=submit value=Submit>
<span style='margin-left:40'><input type=button value=Delete onclick=deleteEntries()></span>
<span style='position:relative;top:-3'>All Tracking List Entries</span><br>
<span style='margin-left:102'><input type=button value=Delete onclick=deleteClients()></span>
<span style='position:relative;top:-3'>All Tracking Lists</span><br><br>
<span style='margin-left:102'><input type=button value=Delete onclick=deleteRecs()></span>
<span style='position:relative;top:-3'>All Tracking List Contribution Recommendations</span>
</td></tr>
</table>
</form>
</body>
</html>
