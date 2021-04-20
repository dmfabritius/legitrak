<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	strRedirect = ""

' SEE IF WE WANT TO LOG OUT
	If Request.Cookies("LegiTrak")("Logout") = "1" Then
		Response.Cookies("LegiTrak") = ""
		Response.Cookies("LegiTrak").Expires= Date-1
		strSQL = "SELECT * FROM [Customer List] WHERE CustomerID=" & CustomerID
		Set rsResult=Server.CreateObject("ADOR.Recordset")
		rsResult.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		rsResult("In Use") = False
		rsResult.Update
		rsResult.Close
		Set rsResult = Nothing
		strRedirect = "top.document.location.href=top.document.location.href"
	End If

' SIGN-ON FORM VARIABLES
	strUsername = TweakQuote(Request.Form("Username"))
	strPassword = TweakQuote(Request.Form("Password"))
	strErrorDisplay = ";display:none"

' IF WE'VE GOT A USERNAME, TRY TO LOG IN
	If Len(strUsername) <> 0 Then
		Set cxnSQL = Server.CreateObject("ADODB.Connection")
		cxnSQL.Open strConnection

		strSQL = "SELECT * FROM [System Status]"
		Set rsResult=cxnSQL.Execute(strSQL)
		SessionStatus = rsResult("Session Status")
		rsResult.Close

		If SessionStatus = 3 Then
			strSQLWhere = " AND (O.[Billing Type] <> 1 OR CC.YR=1)"
		Else
			strSQLWhere = ""
		End If

		strSQL = _
			"(SELECT CC.CustomerID, MAX(CL.[Year-Round]) YR " & _
			"FROM [Customer Clients] CC INNER JOIN [Client List] CL ON CC.ClientID=CL.ClientID " & _
			"GROUP BY CC.CustomerID) CC"
		strSQL = _
			"SELECT" & _
			" O.[Billing Type], C.CustomerID, C.Password, C.[Customer Company Name]," & _
			" C.DefPriority, C.DefPosition " & _
			"FROM [Customer List] C" & _
			" INNER JOIN [Organization List] O ON C.OrganizationID = O.OrganizationID" & _
			"  LEFT JOIN " & strSQL & " ON C.CustomerID=CC.CustomerID " & _
			"WHERE C.[Username]='" & strUsername & "'" & _
			" AND C.[Password]='" & strPassword & "'" & _
			" AND O.[Billing Clients] > 0" & _
			strSQLWhere
		Set rsResult = cxnSQL.Execute(strSQL)

		bolSignon = True
		If rsResult.EOF Then
			bolSignon = False ' If the record wasn't found, then bad password
		Else
			' The SQL call is case-insensitive, so we check that here
			If strPassword <> rsResult("Password") Then bolSignon = False
		End If

		If bolSignon Then
			CustomerID = rsResult("CustomerID")
			Response.Cookies("LegiTrak")("CustomerID") = Encrypt(CustomerID)
			Response.Cookies("LegiTrak")("CustomerName") = rsResult("Customer Company Name")
			Response.Cookies("LegiTrak")("DefPriority") = rsResult("DefPriority")
			Response.Cookies("LegiTrak")("DefPosition") = rsResult("DefPosition")
			Response.Cookies("LegiTrak")("SessionStatus") = SessionStatus
			If rsResult("Billing Type") = 1 Then
				Response.Cookies("LegiTrak")("SessionOnly") = "True"
			Else
				Response.Cookies("LegiTrak")("SessionOnly") = "False"
			End If
			strSQL = _
				"UPDATE [Customer List] SET" & _
				" [In Use]=1," & _
				" [Session Started]='" & Now & "'," & _
				" [IP Address]='" & Request.ServerVariables("REMOTE_ADDR") & "' " & _
				"WHERE CustomerID=" & CustomerID
			cxnSQL.Execute strSQL, , adExecuteNoRecords
			strRedirect = _
				"parent.subheading.document.getElementById(""custName"").innerHTML='" & rsResult("Customer Company Name") & "';" & _
				"parent.menu.location.href='menu-logged-in.asp';" & _
				"parent.contents.location.href='contents.asp';" & _
				"parent.details.location.href='customer.htm';" & _
				"top.document.getElementById(""nav"").rows=""190,*"";"
		Else
			strErrorDisplay = ""
		End If

		Set rsResult = Nothing
		Set cxnSQL = Nothing
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function init(){
	parent.contents.location.href="blank.htm";<%=strRedirect%>
	sizePix()
	document.getElementById("SignOn").Username.focus()
}
</script>
</head>
<body onload='init()' onresize='sizePix()' class=bkg04 style='overflow:hidden'>
<div style='margin:50 0 25 50'>
<form id=SignOn method=post action="signon.asp">
<table width=500 border=0 cellpadding=0 cellspacing=0 class=hdg24 style='padding-right:7;background-color:transparent'>
<col width=150 align=right><col width=350>
<tr><td align=left colspan=2 style='font-size:14pt;padding-bottom:30'>
Please enter your account information</td></tr>
<tr><td>Customer Sign-On:</td><td><input type=text name=Username size=46></td></tr>
<tr><td>Password:</td><td><input type=password name=Password size=46></td></tr>
<tr><td colspan=2 align=right style='padding-right:87'><br><br><input type=submit value=Submit class=btn61></td></tr>
</table>
</form>
</div>

<div class=div16 style='height:80;width:85%;padding:30<%=strErrorDisplay%>'>
<span class=det10>Invalid customer name or password.  Please try again.</span>
</div>

<div id=d1 style="z-index:-1;position:absolute;top:0;left:0;height:85%;margin:25 0 0 50;overflow:hidden;
 filter:progid:DXImageTransform.Microsoft.BasicImage(Opacity=0.15)">
<div id=d2 style="height:100%;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='img/wa.gif', sizingMethod='scale')"></div>
</div>
</body>
</html>