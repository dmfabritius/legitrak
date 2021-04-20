<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"

	If Request.Form("UpdateBilling")="True" Then
	' UPDATE ORGANIZATION BILLING
		Set cmdSQL = CreateObject("ADODB.Connection")
		cmdSQL.Open strConnection
		If Request.Form("OType") = 1 Then
			intClients = Request.Form("BClts1")
		Else
			intClients = Request.Form("BClts2")
		End If
		strCommand = _
			"UPDATE [Organization List] SET" & _
			" Organization='" & TweakQuote(Request.Form("Org")) & "'," & _
			" [Billing Type]=" & Request.Form("BType") & "," & _
			" [Org Type]=" & Request.Form("OType") & "," & _
			" [Billing Amount]=" & CInt("0" & Request.Form("BAmt")) & "," & _
			" [Billing Clients]=" & intClients & "," & _
			" [Billing Notes]='" & TweakQuote(Request.Form("Notes")) & "' " & _
			"WHERE OrganizationID=" & Request.Form("OrgID")
		cmdSQL.Execute strCommand, , adExecuteNoRecords
		cmdSQL.Close
		Set cmdSQL = Nothing
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var detailActive=0
var bdf

function init() {
	bdf=document.getElementById("BillingDetailForm")
	selectTab(3)
	if ((scrollTop=getCookie("scrollTop"))!=null) document.body.scrollTop=scrollTop
}
function sortBy(field) {
	setCookie("BillingOrderField",field)
	window.location.href="maint-billing.asp"
}
function hideDetail(e,override) {
	e = (!e) ? event.srcElement : e.target
	while (e.parentNode!=null && e.id!="BillingDetailForm") e=e.parentNode
	if (e.id=="BillingDetailForm" && !override) return
	if (detailActive==-1)
		detailActive=1;
	else if (detailActive) {
		document.getElementById("BillingDetails").style.display="none"
		setCookie("scrollTop",0)
	}
}
function selectDetail(i) {
	detailActive=-1
	e=document.getElementById("BillingDetails").style
	e.top=document.body.clientHeight+document.body.scrollTop-265

	var camt=document.getElementById("camt"+i).innerHTML
	var bamt=document.getElementById("bamt"+i).innerHTML
	
	bdf.OrgID.value=document.getElementById("orgid"+i).innerHTML
	bdf.Org.value=document.getElementById("org"+i).innerHTML.fromHTML()
	document.getElementById("Con").innerHTML=document.getElementById("con"+i).innerHTML
	document.getElementById("Accts").innerHTML=document.getElementById("accts"+i).innerHTML
	bdf.BType.selectedIndex=document.getElementById("btype"+i).innerHTML
	
	document.getElementById("CalcBAmt").innerHTML=camt
	bdf.Exception.checked=(camt!=bamt)
	bdf.BAmt.disabled=(camt==bamt)
	bdf.BAmt.value=bamt

	cc=document.getElementById("ContractClients")
	ac=document.getElementById("AssocClients")
	if ((bdf.OType.selectedIndex=document.getElementById("otype"+i).innerHTML)==1) {
		bdf.BClts1.selectedIndex=document.getElementById("bclts"+i).innerHTML
		cc.style.display=''
		ac.style.display='none'
	} else {
		bdf.BClts2.selectedIndex=document.getElementById("bclts"+i).innerHTML
		cc.style.display='none'
		ac.style.display=''
	}

	document.getElementById("AClts").innerHTML=document.getElementById("aclts"+i).innerHTML
	document.getElementById("CltsSO").innerHTML=document.getElementById("cltsSO"+i).innerHTML
	document.getElementById("CltsYR").innerHTML=document.getElementById("cltsYR"+i).innerHTML
	bdf.Notes.value=document.getElementById("notes"+i).innerHTML.fromHTML().replace(/&nbsp;/g,"").trim()
	setCookie("scrollTop",document.body.scrollTop)
	e.display="block"
	bdf.Org.focus();
}
function changeOType(){
	if (bdf.OType.selectedIndex==1) {
		cc.style.display=''
		ac.style.display='none'
	} else {
		cc.style.display='none'
		ac.style.display=''
	}
	recalcBilling()
}
function recalcBilling(){
	t=bdf.BType.selectedIndex
	o=bdf.OType.selectedIndex
	if (o==1)
		c=bdf.BClts1.selectedIndex
	else
		c=bdf.BClts2.selectedIndex
	a=parseInt(document.getElementById("Accts").innerHTML)

	base=350
	if (o==1)	// Contract Lobbyists
		amt=base*c
	else		// Organizations
		amt=base*((c==1)? 1 : 10/7)*((a>1)? 1.5 : 1)
	// amt*=(t==0)? 1.5 : 1
	amt=Math.round(amt)

	document.getElementById("CalcBAmt").innerHTML=amt
	b=bdf.BAmt.value
	bdf.Exception.checked=(amt!=b)
	bdf.BAmt.disabled=(amt==b)
}
function verifyException(){
	bdf.BAmt.disabled=!bdf.BAmt.disabled
	if (bdf.BAmt.disabled) bdf.BAmt.value=document.getElementById("CalcBAmt").innerHTML
}
function submitBillings(){
	bdf.BAmt.disabled=false
	bdf.submit()
}
</script>
</head>

<body onload='init()' onclick='hideDetail(arguments[0])' class=bkg04 style='margin:10'>
<form id=BillingSummaryForm>
<table width=100% border=0 cellpadding=0 cellspacing=1 class=det00>
<col width=175><col width=125><col span=6 width=60 align=center><col style='padding-left:5'>
<tr class=hdg24>
<td onclick='sortBy("Organization")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'><br>Organization</div></td>
<td onclick='sortBy("[Contact]")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'><br>Contact Name</div></td>
<td onclick='sortBy("[Org Type]")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Org<br>Type</div></td>
<td onclick='sortBy("[Org Accounts] DESC")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Org<br>Accounts</div></td>
<td onclick='sortBy("[Billing Type],[Org Type] DESC")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Billing<br>Type</div></td>
<td onclick='sortBy("[Billing Amount] DESC")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Billing<br>Amount</div></td>
<td onclick='sortBy("[Billing Clients] DESC")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Billing<br>Clients</div></td>
<td onclick='sortBy("[SO-Clients] DESC")'><div onmouseover='colHover(this,1)' onmouseout='colHover(this,0)'>Actual<br>Clients</div></td>
<td style='cursor:default'><br>Billing Notes</td>
</tr>
<%
	strSQLOrder = Request.Cookies("LegiTrak")("BillingOrderField")
	If strSQLOrder = "" Then
		strSQLOrder = "[Billing Type],[Org Type] DESC,Organization"
	ElseIf strSQLOrder <> "Organization" Then
		strSQLOrder = strSQLOrder & ",Organization"
	End If

	strSQLJoin = "[Client List]"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Customer List] ON [Client List].CustomerID=[Customer List].CustomerID)"
	strSQLJoin = "(" & strSQLJoin & " INNER JOIN [Organization List] ON [Customer List].OrganizationID=[Organization List].OrganizationID)"
	strSQL = _
		"(SELECT O.OrganizationID," & _
		" SUM(CASE WHEN ISNULL(CL.[Year-Round],0)=0 THEN 1 ELSE 0 END) AS [SO-Clients]," & _
        " SUM(CASE WHEN CL.[Year-Round]=1 THEN 1 ELSE 0 END) AS [YR-Clients] " & _
		"FROM [Client List] CL" & _
		" INNER JOIN [Customer List] C ON CL.CustomerID = C.CustomerID" & _
		" INNER JOIN [Organization List] O ON C.OrganizationID = O.OrganizationID " & _
		"GROUP BY O.OrganizationID" & _
		") AS CC"

	strSQLJoin = "[Organization List] O"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN " & strSQL & "  ON O.OrganizationID=CC.OrganizationID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] PriCust ON O.[Primary CustomerID]=PriCust.CustomerID)"
	strSQLJoin = "(" & strSQLJoin & " LEFT JOIN [Customer List] ON O.OrganizationID = [Customer List].OrganizationID)"
	strSQL = _
		"SELECT" & _
		" O.OrganizationID, O.Organization," & _
		" PriCust.[Contact First Name]+' '+PriCust.[Contact Last Name] AS Contact," & _
		" PriCust.Password," & _
		" COUNT([Customer List].CustomerID) AS [Org Accounts]," & _
		" O.[Billing Type]," & _
		" O.[Org Type]," & _
		" O.[Billing Amount]," & _
		" O.[Billing Clients]," & _
		" ISNULL(CC.[SO-Clients],0) AS [SO-Clients], ISNULL(CC.[YR-Clients],0) AS [YR-Clients]," & _
		" O.[Billing Notes] " & _
		"FROM " & strSQLJoin
	strSQL = strSQL & _
		"GROUP BY" & _
		" O.OrganizationID, O.Organization," & _
		" PriCust.[Contact First Name]," & _
		" PriCust.[Contact Last Name]," & _
		" PriCust.[Password]," & _
		" O.[Billing Type]," & _
		" O.[Billing Amount]," & _
		" O.[Billing Clients]," & _
		" CC.[SO-Clients], CC.[YR-Clients]," & _
		" O.[Billing Notes]," & _
		" O.[Org Type]" & _
		"ORDER BY " & strSQLOrder
	Set rstReport=Server.CreateObject("ADOR.Recordset")
	rstReport.Open strSQL, strConnection

	' Calculate summary statistics:
	' 1st Index(0-3): Year-Round, Session-Only, Prospects, Other 
	' 2nd Index(0-2): Lobbyists, Associations, Total
	' 3rd Index(0-4): Orgs, Accounts, BClients, SO-Clients+YR-Clients, Billing Amount
	Dim aSummary(3,2,4)
	
	i=0
	Do Until rstReport.EOF

		x = rstReport("Billing Type")
		y = rstReport("Org Type")
		aSummary(x,y,0) = aSummary(x,y,0)+1
		aSummary(x,y,1) = aSummary(x,y,1)+rstReport("Org Accounts")
		aSummary(x,y,2) = aSummary(x,y,2)+rstReport("Billing Clients")
		aSummary(x,y,3) = aSummary(x,y,3)+rstReport("SO-Clients")+rstReport("YR-Clients")
		aSummary(x,y,4) = aSummary(x,y,4)+rstReport("Billing Amount")
		aSummary(x,2,0) = aSummary(x,2,0)+1
		aSummary(x,2,1) = aSummary(x,2,1)+rstReport("Org Accounts")
		aSummary(x,2,2) = aSummary(x,2,2)+rstReport("Billing Clients")
		aSummary(x,2,3) = aSummary(x,2,3)+rstReport("SO-Clients")+rstReport("YR-Clients")
		aSummary(x,2,4) = aSummary(x,2,4)+rstReport("Billing Amount")

		Select Case rstReport("Billing Type")
			Case 0: strBType = "YR"
			Case 1: strBType = "SO"
			Case 2: strBType = "P"
			Case 3: strBType = "O"
		End Select
		Select Case rstReport("Org Type")
			Case 0: strOType = "A"
			Case 1: strOType = "C"
		End Select

'	Calculate Standard Billing Price
		Base = 750
		CalcBill = 0
		cltsSO = rstReport("SO-Clients")
		cltsYR = rstReport("YR-Clients")
        If (cltsSO + cltsYR) < rstReport("Billing Clients") Then
            cltsSO = rstReport("Billing Clients") - cltsYR
        End If
		If strBType <> "SO" Then
			cltsYR = cltsSO+cltsYR
			cltsSO = 0
		End If
		If strOType="C" Then
			If (cltsYR >= 4) Then
				CalcBill = Base*1.5*4
			ElseIf (4-cltsYR) < cltsSO Then
				CalcBill = Base*1.5*cltsYR + base*(4-cltsYR)
			Else
				CalcBill = Base*1.5*cltsYR + base*cltsSO
			End If
		Else
			If cltsYR = 0 Then
				If cltsSO <= 2 Then
					CalcBill = Base
				Else
					CalcBill = Base*(8/3)
				End If
			ElseIf cltsYR = 1 Then
				If cltsSO < 2 Then
					CalcBill = Base*1.5
				ElseIf cltsSO = 2 Then
					CalcBill = Base*1.5 + Base
				Else
					CalcBill = Base*(8/3)
				End If
			ElseIf cltsYR = 2 Then
				If cltsSO = 0 Then
					CalcBill = Base*1.5
				Else
					CalcBill = Base*(8/3)
				End If
			Else
				CalcBill = Base*1.5*(8/3)
			End If
			If rstReport("Org Accounts") > 1 Then CalcBill = CalcBill*1.5
		End If		
        CalcBill = Round(CalcBill, 0)

		If rstReport("Billing Type") > 1 Then
			strStyle = " style='color:#808080'"
		Else
			strStyle = ""
		End If
		If rstReport("Password") = "changeme" Then
			strPassword = " style='color:red'"
		Else
			strPassword = ""
		End If
		If rstReport("Billing Type") <= 1 And rstReport("Billing Amount") <> CalcBill Then
			strBillDiff = " style='background-color:blue;color:white;font-weight:bold'"
		Else
			strBillDiff = ""
		End If
		If rstReport("Org Type") = 0 Then ' Association
			If rstReport("Billing Clients") < 2 Then
				MaxAllowed = 2
			Else
				MaxAllowed = 50 ' unlimited
			End If
		Else ' Lobbyists
			If rstReport("Billing Clients") < 4 Then
				MaxAllowed = rstReport("Billing Clients")
			Else
				MaxAllowed = 50 ' unlimited
			End If
		End If
		If  rstReport("Billing Type") <= 1 And (cltsSO+cltsYR) > MaxAllowed Then
			strCltDiff = " style='background-color:red;color:white;font-weight:bold'"
		Else
			strCltDiff = ""
		End If

		Response.Write _
			"<tr valign=top>" & _
			"<td id=org" & i & " style='cursor:pointer' onMouseOver='colHover(this,1)' onMouseOut='colHover(this,0)'" & _
			" onclick='selectDetail(" & i & ")'" & ">" & rstReport("Organization") & "</td>" & _
			"<td id=con" & i & strStyle & strPassword & ">" & rstReport("Contact") & "</td>" & _
			"<td" & strStyle & ">" & strOType & "</td>" & _
			"<td id=accts" & i & strStyle & ">" & rstReport("Org Accounts") & "</td>" & _
			"<td" & strStyle & ">" & strBType & "</td>" & _
			"<td id=bamt" & i & strStyle & strBillDiff & ">" & rstReport("Billing Amount") & "</td>" & _
			"<td id=bclts" & i & strStyle & ">" & rstReport("Billing Clients") & "</td>" & _
			"<td id=aclts" & i & strStyle & strCltDiff & ">" & cltsSO+cltsYR & "</td>"
		Response.Write _
			"<td" & strStyle & "><span id=notes" & i & ">" & rstReport("Billing Notes") & "</span>" & _
			"<span style='display:none'>" & _
			"<span id=camt" & i & ">" & CalcBill & "</span>" & _
			"<span id=orgid" & i & ">" & rstReport("OrganizationID") & "</span>" & _
			"<span id=otype" & i & ">" & rstReport("Org Type") & "</span>" & _
			"<span id=btype" & i & ">" & rstReport("Billing Type") & "</span>" & _
			"<span id=cltsSO" & i & ">" & cltsSO & "</span>" & _
			"<span id=cltsYR" & i & ">" & cltsYR & "</span>" & _
			"</span></td></tr>"
		rstReport.MoveNext
		i=i+1
	Loop
	rstReport.Close
	Set rstReport = Nothing
%>
</table>
</form>

<table border=0 cellpadding=0 cellspacing=0 class=det00 style='padding:0 10 0 0'>
<col width=125 align=right><col width=125><col width=75 align=right>
<col width=75 align=right><col width=75 align=right>
<col width=75 align=right><col width=75 align=right>
<tr class=hdg24>
<td></td><td></td>
<td><br>Organizations</td>
<td><br>Accounts</td>
<td>Billing<br>Clients</td>
<td>Actual<br>Clients</td>
<td>Annual<br>Billings</td>
</tr>
<%
	' Summary statistics:
	' 1st Index(0-3): Year-Round, Session-Only, Prospects, Other 
	' 2nd Index(0-2): Lobbyists, Associations, Total
	' 3rd Index(0-4): Orgs, Accounts, BClients, SO-Clients+YR-Clients, Billing Amount
	For i = 0 to 3
		Response.Write "<tr valign=bottom><td><br><b>"
		Select Case i
			Case 0: Response.Write "Year-Round:"
			Case 1: Response.Write "Session-Only:"
			Case 2: Response.Write "Prospects:"
			Case 3: Response.Write "Other:"
		End Select
		Response.Write "</b></td>"

		For j = 0 to 2
			Response.Write "<td><b>"
			Select Case j
				Case 0: Response.Write "Associations"
				Case 1: Response.Write "Lobbyists"
				Case 2: Response.Write "Total"
			End Select
			Response.Write "</b></td>"

			For k = 0 to 4
				Response.Write "<td"
				If j=1 Then Response.Write " style='border-bottom:1px solid black'"
				Response.Write ">" & FormatNumber(aSummary(i,j,k),0,,-1) & "</td>"
			Next 'k
			If j < 2 Then
				Response.Write "</tr><td></td>"
			Else
				Response.Write "</tr>"
			End If
		Next 'j
		If i = 1 Then
			Response.Write _
				"<tr><td colspan=2></td>" & _
				"<td colspan=5 style='border-bottom:3px double black'>&nbsp;</td></tr>"
			Response.Write _
				"<tr><td></td><td bgcolor=#E0E0E0><b>Billing Total</b></td>"
			For k = 0 to 4
				Response.Write "<td bgcolor=#E0E0E0>" & FormatNumber(aSummary(0,2,k)+aSummary(1,2,k),0,,-1) & "</td>"
			Next 'k
		End If
	Next 'i
%>
</table>

<div id=BillingDetails class=div1A style="
z-index:2;display:none;position:absolute;left:15;padding:5 0;
height:250;width:95%;overflow:hidden;
top:expression(document.body.offsetHeight+document.body.scrollTop-265);">

<form id=BillingDetailForm action="maint-billing.asp" method=post>
<input type=hidden name=UpdateBilling value=True>
<input type=hidden name=OrgID>

<table width=100% border=0 cellspacing=0 cellpadding=0 class=hdg10 style='padding-left:10'>
<col width=175 align=right>
<tr><td>Organization:</td><td><input name=Org type=text style='width:300'></td></tr>
<tr><td>Contact:</td><td><span id=Con></span></td></tr>
<tr><td>Organization Accounts:</td><td><span id=Accts></span></td></tr>
<tr><td>Billing Type:</td><td><select name=BType style='width:150' onchange='recalcBilling()'>
<option value=0>Year-Round (Annual)
<option value=1>Session-Only
<option value=2>Prospect
<option value=3>Other
</select></td></tr>
<tr><td>Organization Type:</td><td><select name=OType style='width:150' onchange='changeOType()'>
<option value=0>Association
<option value=1>Contract Lobbyist
</select></td></tr>
<tr><td>Billing Clients:</td><td><span style='width:110'>
<span id=ContractClients>
<select name=BClts1 onchange='recalcBilling()'>
<option value=0>Disabled
<option value=1>1
<option value=2>2
<option value=3>3
<option value=4 selected>Unlimited
</select>
</span>
<span id=AssocClients style='display:none'>
<select name=BClts2 onchange='recalcBilling()'>
<option value=0>Disabled
<option value=1>1-2
<option value=2 selected>Unlimited
</select>
</span>
</span>
Actual Clients: &nbsp; <span style='width:30' id=AClts></span>(Session-Only: <span id=CltsSO></span>, Year-Round: <span id=CltsYR></span>)</td></tr>
<tr><td>Billing Amount:</td><td>
<span id=CalcBAmt style='width:110'></span><input type=checkbox name=Exception onclick='verifyException()'>
Billing Exception, Amount: &nbsp; <input name=BAmt type=text  style='width:50' disabled></td></tr>
<tr valign=top><td>Notes:</td><td><textarea style='width:97%' rows=3 name=Notes></textarea></td></tr>

<tr><td></td><td valign=bottom><br>
<input type=button class=btn61 onclick='submitBillings()' value=Submit><span style='width:200'>&nbsp;</span>
<input type=button class=btn61 onclick='hideDetail(arguments[0],1)' value=Cancel>
</td></tr>
</table>

</form>
</div>

</body>
</html>