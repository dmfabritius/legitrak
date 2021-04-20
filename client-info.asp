<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection
	strRedirect = ""
	bolDeleted = False

' DELETE TRACKING LIST
	If Request.Cookies("LegiTrak")("DeleteClient") = "True" Then
		strCommand = "DELETE FROM [Client List] WHERE ClientID=" & ClientID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteClient") = ""
		Response.Cookies("LegiTrak")("ClientID") = ""
		Response.Cookies("LegiTrak")("ClientName") = ""
		strRedirect = "deletedClient()"
    	bolDeleted = True

' DELETE TRACKING LIST ENTRIES
	ElseIf Request.Cookies("LegiTrak")("DeleteEntries") = "True" Then
		strCommand = "DELETE FROM [Client Specific Bill Info] WHERE ClientID=" & ClientID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteEntries") = ""
		strRedirect = "window.location.href='client-tracking.asp'"

' DELETE TRACKING LIST CANDIDATE RECOMMENDATIONS
	ElseIf Request.Cookies("LegiTrak")("DeleteRecs") = "True" Then
		strCommand = _
			"UPDATE [Client Politician Comments] SET " & _
			" [Group]=NULL," & _
			" [Primary Rec]=NULL," & _
			" [General Rec]=NULL " & _
			"WHERE ClientID=" & ClientID
		cxnSQL.Execute strCommand, , adExecuteNoRecords
		Response.Cookies("LegiTrak")("DeleteRecs") = ""

' UPDATE CLIENT ACCOUNT INFORMATION
	ElseIf Request.Form("UpdateAcctInfo") = "True" Then
		ClientOwner = TweakQuote(Request.Form("ClientOwner"))
		ClientName = TweakQuote(Trim(Request.Form("ClientName")))
		ClientShort = TweakQuote(Trim(Request.Form("ClientShort")))
		FirstName = TweakQuote(Request.Form("FirstName"))
		LastName = TweakQuote(Request.Form("LastName"))
		Addr1 = TweakQuote(Request.Form("Addr1"))
		City = TweakQuote(Request.Form("City"))
		State = Replace(Request.Form("State"),"'","")
		Zip = Replace(Request.Form("Zip"),"'","")
		Email = TweakQuote(Request.Form("Email"))
		strCommand = _
			"UPDATE [Client List] SET" & _
			" [Client Company Name]='" & ClientName & "'," & _
			" [Short Company Name]='" & ClientShort & "'," & _
			" [Contact First Name]='" & FirstName & "'," & _
			" [Contact Last Name]='" & LastName & "'," & _
			" [Address Line 1]='" & Addr1 & "'," & _
			" [City]='" & City & "'," & _
			" [State]='" & State & "'," & _
			" [Zip]='" & Zip & "'," & _
			" [Email]='" & Email & "'," & _
			" [Modified Date]=' " & Now & "' " & _
			"WHERE ClientID=" & ClientID
		cxnSQL.Execute strCommand, , adExecuteNoRecords

		' Add client lobbyist sharing records
		If Request.Form("CustCount") <> "-1" Then
			strCommand = "SELECT CustomerID FROM [Client List] WHERE ClientID=" & ClientID
			Set rsResult = cxnSQL.Execute(strCommand)
			OwnerID = rsResult("CustomerID")
			rsResult.Close
			Set rsResult = Nothing

			strCommand = "DELETE FROM [Customer Clients] WHERE ClientID=" & ClientID & " AND CustomerID <> " & OwnerID
			cxnSQL.Execute strCommand, , adExecuteNoRecords
			For i = 0 to Request.Form("CustCount")
				If Len(Request.Form("c" & i)) <> 0 Then
				    cltID = Decrypt(Request.Form("c" & i))
				    If cltID <> 0 Then
					    strCommand = "INSERT INTO [Customer Clients] VALUES (" & _
						    cltID & "," & ClientID & ")"
					    cxnSQL.Execute strCommand, , adExecuteNoRecords
					End If
				End If
			Next ' i
		End If
		strRedirect = _
			"top.contents.location.href='contents.asp';" & _
			"top.subheading.document.getElementById('cltName').innerHTML='" & ClientName & "';"
    End If
	
' LOAD CLIENT ACCOUNT INFORMATION
    If Request.Form("UpdateAcctInfo") <> "True" And Not bolDeleted Then
		strCommand = _
			"SELECT [Customer List].* " & _
			"FROM [Customer List] INNER JOIN [Client List]" & _
			" ON [Customer List].CustomerID = [Client List].CustomerID " & _
			"WHERE [Client List].ClientID=" & ClientID
		Set rsCust = cxnSQL.Execute(strCommand)
		OwnerID = rsCust("CustomerID")
		ClientOwner = rsCust("Contact First Name") & " " & rsCust("Contact Last Name")
		rsCust.Close
		Set rsCust = Nothing

		strCommand = "SELECT * FROM [Client List] WHERE ClientID=" & ClientID
		Set rsClient = cxnSQL.Execute(strCommand)
		ClientName = rsClient("Client Company Name")
		ClientShort = rsClient("Short Company Name")
		FirstName = rsClient("Contact First Name")
		LastName = rsClient("Contact Last Name")
		Addr1 = rsClient("Address Line 1")
		City = rsClient("City")
		State = rsClient("State")
		Zip = rsClient("Zip")
		Email = rsClient("Email")
		rsClient.Close
		Set rsClient = Nothing
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function init(){<%=strRedirect%>
	selectTab(2)
	document.getElementById("Info").ClientName.focus()
}
function deleteEntries(){
	if (confirm("Do you want to delete this tracking list's entries?"))
		if (confirm(
				"Deleting this tracking list's entries will permanently erase all of the bills in the list.\n\n"+
				"Are you absolutely sure you want to do this?")) {
			setCookie("DeleteEntries","True")
			window.location.href=window.location.href
		}
}
function deleteClient(){
	if (confirm("Do you want to delete this tracking list?"))
		if (confirm(
				"Deleting this tracking list will permanently erase all associated information, including list entries, report comments, "+
				"keywords, and candidate contributions.\n\nAre you absolutely sure you want to delete this tracking list?")) {
			setCookie("DeleteClient","True")
			window.location.href=window.location.href
		}
}
function deleteRecs(){
	if (confirm("Do you want to delete this tracking list's candidate contribution recommendations?"))
		if (confirm(
				"This will permanently erase all of the candidate contribution recommendations for this list.\n\n"+
				"Are you absolutely sure you want to do this?")) {
			setCookie("DeleteRecs","True")
			window.location.href=window.location.href
		}
}
function deletedClient(){
	setCookie("menuItem","mnu1")
	top.contents.location.href="contents.asp"
	sh=top.subheading.document
	sh.getElementById('cltName').innerHTML=""
	sh.getElementById('sep').innerHTML=""
	selectMenu("mnu1","customer.htm")
}
function switchInfo() {
	switch (document.getElementByName("InfoItem")[0].value){
		case 'CltInfo' : window.location.href='client-info.asp'; break
		case 'CltMem'  : window.location.href='client-members.asp'
	}
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:10 20'>
<%
' DON'T DISPLAY THE LINK TO THE ASSOCIATION MEMBERSHIP PAGE EXCEPT FOR BRAD
'	If CustomerID=1 Then
'		Response.Write _
'			"<form><select name=InfoItem class=hdg10 onchange='switchInfo()'>" & _
'			"<option value='CltInfo' selected>General Account Info" & _
'			"<option value='CltMem'>Client Members" & _
'			"</select></form>"
'	End If
%>
<form id=Info method=post action="client-info.asp">
<input type=hidden name=UpdateAcctInfo value=True>
<table border=0 cellpadding=0 cellspacing=1 class=det00 style='padding-left:10'>
<col width=150 align=right class=hdg24>
<tr class=hdg24><td>Owner:</td><td><input type=text readonly class=bkg04 style='border-width:0' name=ClientOwner value='<%=TweakQuote(ClientOwner)%>'></td></tr>
<tr><td>List Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=ClientName value='<%=ClientName%>'></td></tr>
<tr><td>Short Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=ClientShort value='<%=ClientShort%>'></td></tr>
<tr><td>Contact First Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=FirstName value='<%=FirstName%>'></td></tr>
<tr><td>Contact Last Name:</td><td><input type=text style='width:325' onchange='mark(this)' name=LastName value='<%=LastName%>'></td></tr>
<tr><td>Email:</td><td><input type=text style='width:325' onchange='return isEmail(this)' name=Email value='<%=Email%>'></td></tr>
<tr><td>Address:</td><td><input type=text style='width:325' onchange='mark(this)' name=Addr1 value='<%=Addr1%>'></td></tr>
<tr><td></td><td>
<input type=text style='width:200' onchange='mark(this)' name=City value='<%=City%>'>
<input type=text style='width:30' onchange='mark(this)' name=State value='<%=State%>' maxlength=2>
<input type=text style='width:87' onchange='mark(this)' name=Zip value='<%=Zip%>' maxlength=10></td></tr>

<tr class=shd24><td></td><td><br><input type=submit value=Submit>
<span style='margin-left:60'><input type=button value=Delete onclick=deleteEntries()></span>
<span style='position:relative;top:-3'>All Tracking List Entries</span><br>
<span style='margin-left:122'><input type=button value=Delete onclick=deleteClient()></span>
<span style='position:relative;top:-3'>Entire Tracking List</span><br><br>
<span style='margin-left:122'><input type=button value=Delete onclick=deleteRecs()></span>
<span style='position:relative;top:-3'>All Contribution Recommendations</span>
</td></tr>
<%
' ALLOW MODIFICATION OF CLIENT SHARING ONLY IF
' THE SIGNED ON CUSTOMER IS THE CLIENT'S OWNER
	If CustomerID = OwnerID Then
	
	' Include only those customers that are in this customer's organization
		strSQL = "SELECT OrganizationID FROM [Customer List] WHERE CustomerID=" & CustomerID
		SQLJoin = "([Customer List] CL INNER JOIN (" & strSQL & ") O ON CL.OrganizationID = O.OrganizationID)"
	' All customers, whether or not they're linked to this client
		strSQL = "SELECT * FROM [Customer Clients] WHERE ClientID=" & ClientID
		SQLJoin = "(" & SQLJoin & " LEFT JOIN (" & strSQL & ") CC ON CL.CustomerID = CC.CustomerID) "
		strSQL = _
			"SELECT" & _
			" CL.CustomerID," & _
			" CL.[Contact First Name], CL.[Contact Last Name]," & _
			" CC.ClientID " & _
			"FROM " & SQLJoin & _
			"WHERE CL.CustomerID <> " & OwnerID & " ORDER BY CL.[Contact First Name]"
		Set rsCustomers = cxnSQL.Execute(strSQL)
		If Not rsCustomers.EOF Then
			aCustomers = rsCustomers.GetRows()
			intCustCount = UBound(aCustomers,2)
			Response.Write _
				"<tr valign=top><td><input type=hidden name=CustCount value=" & intCustCount & ">" & _
				"<br>Share list with:</td><td><br>"
			For i = 0 to intCustCount
				Response.Write "<input type=checkbox"
				If Not IsNull(aCustomers(3,i)) Then
					Response.Write " checked"
				End If
				Response.Write _
					" name=c" & i & " value=" & Encrypt(aCustomers(0,i)) & "> &nbsp;" & _
					aCustomers(1,i) & " " & aCustomers(2,i) & "<br>"
			Next 'i
			Response.Write "</td></tr>"
		Else
			Response.Write "<tr><td><input type=hidden name=CustCount value=-1></td></tr>"
		End If
		rsCustomers.Close
		Set rsCustomers = Nothing
	Else
		Response.Write "<tr><td><input type=hidden name=CustCount value=-1></td></tr>"
	End If

	cxnSQL.Close
	Set cxnSQL = Nothing
%>
</table>
</form>
</body>
</html>
