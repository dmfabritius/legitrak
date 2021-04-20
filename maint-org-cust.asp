<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection

	OrgID = CLng("0" & Request.Form("OrgID"))
	OrgCustID = CLng("0" & Request.Form("OrgCustID"))

' DELETE ORGANIZATION
	If Request.Form("UpdateOrg") = "Delete" Then
		strSQL = "DELETE FROM [Organization List] WHERE OrganizationID=" & OrgID
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		OrgID = 0
		OrgCustID = 0
	End If

' ADD/UPDATE ORGANIZATION
	If Request.Form("UpdateOrg") = "True" Then
		If OrgID <> 9999999 Then
			If Request.Form("PriCust") <> "" Then
				intPriCust = Request.Form("PriCust")
			Else
				intPriCust = 0
			End If
			strSQL = _
				"UPDATE [Organization List] SET" & _
				" Organization='" & TweakQuote(Request.Form("Org")) & "'," & _
				" [Primary CustomerID]=" & intPriCust & "," & _
				" [Billing Type]=" & Request.Form("BType") & "," & _
				" [Org Type]=" & Request.Form("OType") & "," & _
				" [Billing Amount]=" & CInt("0" & Request.Form("BAmt")) & "," & _
				" [Billing Clients]=" & Request.Form("BClts") & "," & _
				" [Billing Notes]='" & TweakQuote(Request.Form("Notes")) & "' " & _
				"WHERE OrganizationID=" & OrgID
			cxnSQL.Execute strSQL, , adExecuteNoRecords
		Else
			strSQL = _
				"INSERT INTO [Organization List] " & _
				"(Organization, [Billing Type], [Org Type]," & _
				" [Billing Amount], [Billing Clients], [Billing Notes]) " & _
				"VALUES (" & _
				"'" & TweakQuote(Request.Form("Org")) & "'," & _
				Request.Form("BType") & "," & _
				Request.Form("OType") & "," & _
				CInt("0" & Request.Form("BAmt")) & "," & _
				Request.Form("BClts") & "," & _
				"'" & TweakQuote(Request.Form("Notes")) & "')"
			cxnSQL.Execute strSQL, , adExecuteNoRecords
			strSQL = "SELECT MAX(OrganizationID) AS MaxID FROM [Organization List]"
			Set rsResult = cxnSQL.Execute(strSQL)
			OrgID = rsResult("MaxID")
			OrgCustID = 9999999
			Set rsResult = Nothing
		End If
	End If

' DELETE ORGANIZATION CUSTOMER ACCOUNT
	If Request.Form("UpdateCust") = "Delete" Then
		strSQL = "DELETE FROM [Customer List] WHERE CustomerID=" & OrgCustID
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		OrgCustID = 0
	End If

' ADD/UPDATE ORGANIZATION CUSTOMER ACCOUNT
	If Request.Form("UpdateCust") = "True" Then
		intVcards = 0
		If Request.Form("Vcards") = "True" Then intVcards = 1
		intLists = 0
		If Request.Form("CreateLists") = "True" Then intLists = 1
		If OrgCustID <> 9999999 Then
			strSQL = _
				"UPDATE [Customer List] SET" & _
				" [Contact First Name]='" & TweakQuote(Request.Form("First")) & "'," & _
				" [Contact Last Name]='" & TweakQuote(Request.Form("Last")) & "'," & _
				" Email='" & TweakQuote(Request.Form("Email")) & "'," & _
				" Username='" & TweakQuote(Request.Form("User")) & "'," & _
				" Password='" & TweakQuote(Request.Form("Pass")) & "'," & _
				" CreateVotecards=" & intVcards & "," & _
				" CreateLists=" & intLists & "," & _
				" [Customer Company Name]='" & TweakQuote(Request.Form("Company")) & "'," & _
				" [Address Line 1]='" & TweakQuote(Request.Form("Addr1")) & "'," & _
				" [Address Line 2]='" & TweakQuote(Request.Form("Addr2")) & "'," & _
				" City='" & TweakQuote(Request.Form("City")) & "'," & _
				" State='" & TweakQuote(Request.Form("State")) & "'," & _
				" Zip='" & TweakQuote(Request.Form("Zip")) & "' " & _
				"WHERE CustomerID=" & OrgCustID
			cxnSQL.Execute strSQL, , adExecuteNoRecords
		Else
			strSQL = _
				"INSERT INTO [Customer List] " & _
				"([Active Report], [Contact First Name], [Contact Last Name], Email," & _
				" Username, Password, CreateVotecards, CreateLists, [Customer Company Name]," & _
				" [Address Line 1], [Address Line 2], City, State, Zip," & _
				" OrganizationID, [Created Date]) " & _
				"VALUES (1," & _
				"'" & TweakQuote(Request.Form("First")) & "'," & _
				"'" & TweakQuote(Request.Form("Last")) & "'," & _
				"'" & TweakQuote(Request.Form("Email")) & "'," & _
				"'" & TweakQuote(Request.Form("User")) & "'," & _
				"'" & TweakQuote(Request.Form("Pass")) & "'," & _
				intVcards & "," & _
				intLists & "," & _
				"'" & TweakQuote(Request.Form("Company")) & "'," & _
				"'" & TweakQuote(Request.Form("Addr1")) & "'," & _
				"'" & TweakQuote(Request.Form("Addr2")) & "'," & _
				"'" & TweakQuote(Request.Form("City")) & "'," & _
				"'" & TweakQuote(Request.Form("State")) & "'," & _
				"'" & TweakQuote(Request.Form("Zip")) & "'," & _
				OrgID & "," & _
				"'" & Date & "')"
			cxnSQL.Execute strSQL, , adExecuteNoRecords
			strSQL = "SELECT [Primary CustomerID] FROM [Organization List] WHERE OrganizationID=" & OrgID
			Set rsResult = cxnSQL.Execute(strSQL)
			If IsNull(rsResult("Primary CustomerID")) Then
				strSQL = "SELECT MAX(CustomerID) AS MaxID FROM [Customer List]"
				Set rsResult = cxnSQL.Execute(strSQL)
				strSQL = _
					"UPDATE [Organization List]" & _
					" SET [Primary CustomerID]=" & rsResult("MaxID") & _
					" WHERE OrganizationID=" & OrgID
				cxnSQL.Execute strSQL, , adExecuteNoRecords
				Set rsResult = Nothing
			End If
			
		End If

' SET YEAR-ROUND STATUS FOR CUSTOMER'S CLIENTS
		strSQL = _
			"UPDATE [Client List] SET [Year-Round]=NULL WHERE ClientID IN " & _
			"(SELECT ClientID FROM [Customer Clients] WHERE CustomerID=" & OrgCustID & ")"
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		If Len(Request.Form("YR")) <> 0 Then
			aClts = Split(Request.Form("YR"),",")
			For i = 0 to UBound(aClts)
				strSQL = "UPDATE [Client List] SET [Year-Round]=1 WHERE ClientID=" & aClts(i)
				cxnSQL.Execute strSQL, , adExecuteNoRecords
			Next 'i
		End If

		OrgCustID = 0
	End If

	cxnSQL.Close
	Set cxnSQL = Nothing

' LOAD ORGANIZATION CUSTOMER INFORMATION
	intOrgCustCount = -1
	If OrgID <> 0 And OrgID <> 9999999 Then
		strSQL = 	"(" & _
			"SELECT" & _
			" C.CustomerID, COUNT(CC.ClientID) AS Clients " & _
			"FROM [Customer List] C INNER JOIN [Customer Clients] CC" & _
			" ON C.CustomerID = CC.CustomerID " & _
			"GROUP BY C.CustomerID" & _
			") TL"
		strSQL = 	_
			"SELECT" & _
			" CL.CustomerID, CL.[Contact First Name], CL.[Contact Last Name]," & _
			" CL.Email, CL.Username, CL.Password," & _
			" CL.[Session Started], CL.CreateVotecards," & _
			" CL.[Customer Company Name]," & _
			" CL.[Address Line 1], CL.[Address Line 2]," & _
			" CL.City, CL.State, CL.Zip," & _
			" ISNULL(TL.Clients,0), CL.CreateLists " & _
			"FROM [Customer List] CL LEFT JOIN " & strSQL & _
			" ON CL.CustomerID=TL.CustomerID " & _
			"WHERE CL.OrganizationID=" & OrgID & _
			" ORDER BY CL.[Contact First Name]+CL.[Contact Last Name]"
		Set rsOrgCusts=Server.CreateObject("ADOR.Recordset")
		rsOrgCusts.Open strSQL, strConnection
		If Not rsOrgCusts.EOF Then
			aOrgCusts = rsOrgCusts.GetRows()
			intOrgCustCount = UBound(aOrgCusts,2)
			For i = 0 to intOrgCustCount
				If aOrgCusts(0,i) = OrgCustID Then OrgCustIndex = i
			Next 'i
		End If
		rsOrgCusts.Close
		Set rsOrgCusts = Nothing
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var odf,cdf

function init(){
	odf=document.getElementById("OrgDetailForm")
	cdf=document.getElementById("CustDetailForm")

	selectTab(0)
	if (<%=OrgCustID%>!=0) cdf.First.focus()
	if (<%=OrgID%>!=0) odf.Org.focus()
}
function orgSelect(){
	odf.UpdateOrg.value="False"
	odf.submit()
}
function orgCancel(){
	odf.UpdateOrg.value="False"
	odf.OrgID.disabled=true
	odf.submit()
}
function orgDelete(){
	if (confirm("Click OK to confirm delete for this entire ORGANIZATION!")){
		odf.UpdateOrg.value="Delete"
		odf.submit()
	}
}
function custSelect(c){
	if (c!=-1) cdf.OrgCustID.value=c
	cdf.UpdateCust.value="False"
	cdf.submit()
}
function custCancel(){
	cdf.UpdateCust.value="False"
	cdf.OrgCustID.disabled=true
	cdf.submit()
}
function custDelete(){
	if (confirm("Click OK to confirm delete for this customer.")){
		cdf.UpdateCust.value="Delete"
		cdf.submit()
	}
}
function recalcBilling(){
	t=odf.BType.selectedIndex
	o=odf.OType.selectedIndex
	a=odf.AClts.value
	c=odf.BClts.selectedIndex

	base=750
	if (o==1)	// Contract Lobbyists
		camt=base*c
	else		// Organizations
		camt=base*((c==2)? 8/3 : 1)*((a>1)? 1.5 : 1)
	camt*=(t==0)? 1.5 : 1

	odf.BAmt.value=camt
}
function orgChanged(){
	if (odf.odfSubmit) odf.odfSubmit.disabled=true
	if (odf.odfDelete) odf.odfDelete.disabled=true
}
</script>
</head>
<body onload='init()' class=bkg04 style='margin:10'>

<!-- ORGANIZATION LIST -->
<form id=OrgDetailForm action="maint-org-cust.asp" method=post>
<input type=hidden name=UpdateOrg value=True>
<span class=hdg24>Organization: </span><select name=OrgID style='width:250' onchange='orgChanged()'>
<option value=9999999>-Add New-
<%
	strSQL = 	"SELECT OrganizationID,Organization FROM [Organization List] ORDER BY Organization"
	Set rsOrgs=Server.CreateObject("ADOR.Recordset")
	rsOrgs.Open strSQL, strConnection
	Do Until rsOrgs.EOF
		Response.Write "<option value=" & rsOrgs("OrganizationID")
		If OrgID=rsOrgs("OrganizationID") Then Response.Write " selected"
		Response.Write ">" & rsOrgs("Organization")
		rsOrgs.MoveNext
	Loop
	rsOrgs.Close
	Set rsOrgs = Nothing
%>
</select>&nbsp; <input type=button value=Select onclick='orgSelect()'>

<!-- ORGANIZATION DETAILS -->
<%
	If OrgID <> 0 Then
		Dim SelBType(3),SelOType(1),SelBClts(5)
		intBAmt = 0
		If OrgID <> 9999999 Then
			strSQL = 	_
				"SELECT" & _
				" O.OrganizationID, O.Organization, O.[Primary CustomerID]," & _
				" O.[Billing Type], O.[Org Type], O.[Billing Amount]," & _
				" O.[Billing Clients], O.[Billing Notes], COUNT(C.CustomerID) AS Accounts " & _
				"FROM [Organization List] O LEFT JOIN [Customer List] C ON O.OrganizationID = C.OrganizationID " & _
				"GROUP BY" & _
				" O.OrganizationID, O.Organization, O.[Primary CustomerID]," & _
				" O.[Billing Type], O.[Org Type], O.[Billing Amount]," & _
				" O.[Billing Clients], O.[Billing Notes] " & _
				"HAVING O.OrganizationID=" & OrgID
			Set rsOrg=Server.CreateObject("ADOR.Recordset")
			rsOrg.Open strSQL, strConnection

			strOrg=rsOrg("Organization")
			intPriCust=rsOrg("Primary CustomerID")
			SelBType(rsOrg("Billing Type"))=" selected"
			SelOType(rsOrg("Org Type"))=" selected"
			intBAmt=rsOrg("Billing Amount")
			SelBClts(rsOrg("Billing Clients"))=" selected"
			intAClts=rsOrg("Accounts")
			strNotes=rsOrg("Billing Notes")

			rsOrg.Close
			Set rsOrg = Nothing
		Else
			selBType(2)=" selected"
		End If
%>
<div class=box20 style='padding:7'>
<table border=0 cellpadding=0 cellspacing=0 width=800 class=shd24 style='padding-left:3'>
<col width=150 align=right><col width=300>
<col width=125 align=right><col width=225>
<tr><td>Organization:</td><td><input name=Org type=text style='width:280' value="<%=strOrg%>"></td>
<td>Contact:</td><td class=det00>
<%
		If 	intOrgCustCount > 0 Then
			Response.Write "<select name=PriCust style='width:160'>"
			For i = 0 to intOrgCustCount
				Response.Write "<option value=" & aOrgCusts(0,i)
				If intPriCust = aOrgCusts(0,i) Then Response.Write " selected"
				Response.Write ">" & aOrgCusts(1,i) & " " & aOrgCusts(2,i)
			Next 'i
			Response.Write "</select>"
		ElseIf intOrgCustCount = 0 Then
			Response.Write "<input type=hidden name=PriCust value=" & aOrgCusts(0,0) & ">"
			Response.Write aOrgCusts(1,0) & " " & aOrgCusts(2,0)
		Else
			Response.Write "<span style='color:gray'>Not available</span>"
		End If
%>
</td></tr>

<tr><td>Organization Type:</td><td>
<select name=OType onchange='recalcBilling()' style='width:160'>
<option value=0<%=SelOType(0)%>>Association
<option value=1<%=SelOType(1)%>>Contract Lobbyist
</select></td>
<td>Billing Type:</td><td>
<select name=BType onchange='recalcBilling()' style='width:160'>
<option value=0<%=SelBType(0)%>>Year-Round (Annual)
<option value=1<%=SelBType(1)%>>Session-Only
<option value=2<%=SelBType(2)%>>Prospect
<option value=3<%=SelBType(3)%>>Other
</select></td></tr>
<tr><td>Billing Clients/Lists:</td><td>
<input name=AClts type=hidden value=<%=intAClts%>>
<select name=BClts onchange='recalcBilling()' style='width:160'>
<option value=0<%=SelBClts(0)%>>Disabled
<option value=1<%=SelBClts(1)%>>1 (Up to 2 for Assoc)
<option value=2<%=SelBClts(2)%>>2 (Unlimited for Assoc)
<option value=3<%=SelBClts(3)%>>3
<option value=4<%=SelBClts(4)%>>Unlimited
</select></td>
<td>Billing Amount:</td><td><input name=BAmt type=text style='width:50' value=<%=intBAmt%>></td></tr>
<tr><td valign=top>Notes:</td><td colspan=3>
<!--
<input name=Notes type=text style='width:280' value="<%=strNotes%>">
-->
<textarea name=Notes cols=115 rows=4><%=strNotes%></textarea>

</td></tr>
</table>

<center>
<span style='height:25'></span>
<input id=odfSubmit type=submit value=Submit>
<span style='width:200'></span>
<input type=button onclick='orgCancel()' value=Cancel><span style='width:200'></span>
<input id=odfDelete type=button onclick='orgDelete()' value=Delete>
</center>

</div>
<%
	End If
%>
</form>

<!-- ORGANIZATION ACCOUNTS SUMMARY -->
<%
	If 	intOrgCustCount <> -1 Then
		Response.Write _
			"<span class=hdg24>Organization Accounts Summary</span>" & _
			"<div class=box20 style='padding:10'>" & _
			"<table border=0 cellpadding=0 cellspacing=0 width=800 class=det00>" & _
			"<col width=150><col width=200><col width=100 align=center>" & _
			"<col width=100><col width=100><col width=150>"
		Response.Write _
			"<tr class=hdg24 style='text-decoration:underline'>" & _
			"<td>Customer Name</td>" & _
			"<td>E-mail Address</td>" & _
			"<td>Tracking Lists</td>" & _
			"<td>Username</td>" & _
			"<td>Password</td>" & _
			"<td>Last Log-on</td></tr>"

		For i = 0 to intOrgCustCount
			Response.Write _
				"<tr valign=top><td>" & _
				"<div onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
				" style='cursor:pointer' onclick='custSelect(" & aOrgCusts(0,i) & ")'>" & _
				aOrgCusts(1,i) & " " & aOrgCusts(2,i) & "</div></td><td>" & _
				aOrgCusts(3,i) & "</td><td>" & _
				aOrgCusts(14,i) & "</td><td>" & _
				aOrgCusts(4,i) & "</td><td>"
			If aOrgCusts(5,i)="changeme" Then
				Response.Write "<span style='color:red'>changeme</span>"
			Else
				Response.Write "******"
			End If
			Response.Write _
				"</td><td>" & _
				aOrgCusts(6,i) & "</td></tr>"
		Next 'i
		Response.Write "</table></div>"
	End If
%>

<!-- CUSTOMER DETAILS -->
<form id=CustDetailForm action="maint-org-cust.asp" method=post>
<input type=hidden name=OrgID value=<%=OrgID%>>
<input type=hidden name=UpdateCust value=True>
<%
	If OrgID <> 0 and OrgID <> 9999999 Then
		Response.Write _
			"<span class=hdg24>Customer Details: </span>" & _
			"<select name=OrgCustID style='width:250'>" & _
			"<option value=9999999>-Add New-"

		If 	intOrgCustCount <> -1 Then
			For i = 0 to intOrgCustCount
				Response.Write "<option value=" & aOrgCusts(0,i)
				If OrgCustID = aOrgCusts(0,i) Then Response.Write " selected"
				Response.Write ">" & aOrgCusts(1,i) & " " & aOrgCusts(2,i)
			Next 'i
		End If

		Response.Write _
			"</select>" & _
			"&nbsp; <input type=button value='Select' onclick='custSelect(-1)'>"
	End If

	strCompany = strOrg
	If OrgCustID <> 0 Then
		strPass = "changeme"
		If OrgCustID <> 9999999 Then
			strFirst = aOrgCusts(1,OrgCustIndex)
			strLast = aOrgCusts(2,OrgCustIndex)
			strEmail = aOrgCusts(3,OrgCustIndex)
			strUser = aOrgCusts(4,OrgCustIndex)
			strPass = aOrgCusts(5,OrgCustIndex)
			If aOrgCusts(7,OrgCustIndex) Then strVC = " checked"
			strCompany = aOrgCusts(8,OrgCustIndex)
			strAddr1 = aOrgCusts(9,OrgCustIndex)
			strAddr2 = aOrgCusts(10,OrgCustIndex)
			strCity = aOrgCusts(11,OrgCustIndex)
			strState = aOrgCusts(12,OrgCustIndex)
			strZip = aOrgCusts(13,OrgCustIndex)
			If aOrgCusts(15,OrgCustIndex) Then strCL = " checked"
		Else
			strCL = " checked"
		End If
%>
<div class=box20 style='padding:7'>
<table border=0 cellpadding=0 cellspacing=0 width=800 class=shd24 style='padding-left:3'><col align=right>
<tr><td>Name:</td><td>
<input name=First type=text style='width:150' value="<%=strFirst%>">
<input name=Last type=text style='width:150' value="<%=strLast%>"></td></tr>
<tr><td>Sign-On:</td><td>
<input name=User type=text style='width:150' value="<%=strUser%>">
<input name=Pass type=text style='width:150' value="<%=strPass%>"></td></tr>
<tr><td>E-Mail:</td><td><input name=Email type=text style='width:304' value="<%=strEmail%>" onchange='return isEmail(this)'></td></tr>
<tr><td>Company Name:</td><td><input name=Company type=text style='width:304' value="<%=strCompany%>"></td></tr>
<tr><td>Address:</td><td><input name=Addr1 type=text style='width:304' value="<%=strAddr1%>"></td></tr>
<tr><td></td><td>
<input name=City type=text style='width:180' value="<%=strCity%>">
<input name=State type=text style='width:30' value="<%=strState%>">
<input name=Zip type=text style='width:86' value="<%=strZip%>"></td></tr>
<tr><td>Votecards:</td><td><input name=Vcards type=checkbox value=True <%=strVC%>></td></tr>
<tr><td>Create Lists:</td><td><input name=CreateLists type=checkbox value=True <%=strCL%>></td></tr>
<tr style='height:25'><td colspan=2 align=center valign=bottom>
<input type=submit value=Submit><span style='width:200'></span>
<input type=button onclick='custCancel()' value=Cancel><span style='width:200'></span>
<input type=button onclick='custDelete()' value=Delete>
</td></tr>
</table>

<!-- CUSTOMER TRACKING LIST SUMMARY -->
<%
		Response.Write _
			"<br><table border=0 cellpadding=0 cellspacing=0 width=700 class=det00>" & _
			"<col width=274><col width=150><col width=150><col span=2 width=100 align=center>" & _
			"<tr class=hdg24 style='text-decoration:underline'>" & _
			"<td>Client Name</td>" & _
			"<td>Owner</td>" & _
			"<td>Contact</td>" & _
			"<td>Bills Tracked</td>" & _
			"<td>Year-Round</td></tr>"

		strSQL = 	_
			"SELECT" & _
			" CL.ClientID, CL.[Client Company Name], ISNULL(CL.[Year-Round],0) YR," & _
			" CL.[Contact First Name], CL.[Contact Last Name]," & _
			" CU.[Contact First Name] AS First, CU.[Contact Last Name] AS Last," & _
			" COUNT(CS.[Bill Number]) AS Bills " & _
			"FROM [Customer Clients] CC" & _
			" LEFT JOIN [Client List] CL ON CC.ClientID = CL.ClientID" & _
			" LEFT JOIN [Customer List] CU ON CL.CustomerID = CU.CustomerID" & _
			" LEFT JOIN [Client Specific Bill Info] CS ON CL.ClientID = CS.ClientID " & _
			"GROUP BY" & _
			" CL.[ClientID], CL.[Client Company Name], CL.[Year-Round]," & _
			" CL.[Contact First Name], CL.[Contact Last Name]," & _
			" CU.[Contact First Name], CU.[Contact Last Name]," & _
			" CC.CustomerID " & _
			"HAVING CC.CustomerID=" & OrgCustID & _
			" ORDER BY CL.[Client Company Name]"
		Set rsClts=Server.CreateObject("ADOR.Recordset")
		rsClts.Open strSQL, strConnection

		Do Until rsClts.EOF
			If rsClts("YR") = 1 Then
				strChecked = " checked"
			Else
				strChecked = ""
			End If
			Response.Write _
				"<tr><td>" & _
				rsClts("Client Company Name") & _
				"</td><td>" & _
				rsClts("First") & " " & _
				rsClts("Last") & _
				"</td><td>" & _
				rsClts("Contact First Name") & " " & _
				rsClts("Contact Last Name") & _
				"</td><td>" & _
				rsClts("Bills") & _
				"</td><td>" & _
				"<input type=checkbox name=YR value=" & rsClts("ClientID") & strChecked & ">" & _
				"</td></tr>"
			rsClts.MoveNext
		Loop
		rsClts.Close
		Set rsClts = Nothing

		Response.Write "</table>"
	End If
%>
</div>
</form>
</body>
</html>