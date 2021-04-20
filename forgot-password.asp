<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
' QUEUE UP REQUESTED CUSTOMER PASSWORD REPORT
	If Request.Form("PasswordRequested") = "True" Then

		strFirst = TweakQuote(Request.Form("FirstName"))
		strLast = TweakQuote(Request.Form("LastName"))
		strZip = Replace(Request.Form("CustomerZip"),"'","")

		Set cxnSQL = CreateObject("ADODB.Connection")
		cxnSQL.Open strConnection
		strCommand= _
			"SELECT * FROM [Customer List] WHERE" & _
			" [Contact First Name]='" & strFirst & "' AND" & _
			" [Contact Last Name]='" & strLast & "' AND" & _
			" [Zip]='" & strZip & "'"

		Set rsCustomer = cxnSQL.Execute(strCommand)

		RequestResults = "Customer not found, please re-enter."
		If Not rsCustomer.EOF Then
			RequestResults = "Your request was submitted successfully."
			strCommand = _
				"INSERT INTO [Report Queue]" & _
				"(ReportID, CustomerID, ClientID, [Report Status], [Effective Date])" & _
				"VALUES (6, " & rsCustomer("CustomerID") & ", 0, 'REQUESTED', GETDATE())"
			cxnSQL.Execute strCommand, , adExecuteNoRecords
		End If
		rsCustomer.Close
		Set rsCustomer = Nothing

		cxnSQL.Close
		Set cxnSQL = Nothing
	Else
		HideResults = ";display:none"
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function init(){
    top.contents.location.href="blank.htm"
	sizePix()
	document.getElementById("Pwd").FirstName.focus()
}
</script>
</head>
<body onload='init()' onresize='sizePix()' class=bkg04>

<div style='margin:50 0 25 50'>
<form id=Pwd method=post action='forgot-password.asp'>
<input type=hidden name=PasswordRequested value=True>
<table width=500 border=0 cellpadding=0 cellspacing=0 class=hdg24 style='padding-right:7;background-color:transparent'>
<col width=150 align=right><col width=350>
<tr><td align=left colspan=2 style='font-size:14pt;padding-bottom:30'>
Please enter your account information to have<br>your password e-mailed to you</td></tr>
<tr><td>First Name:</td><td><input type=text name=FirstName size=46></td></tr>
<tr><td>Last Name:</td><td><input type=text name=LastName size=46></td></tr>
<tr><td>Zip Code:</td><td><input type=text name=CustomerZip size=46></td></tr>
<tr><td colspan=2 style='padding-right:87'><br><br><input type=submit value=Submit class=btn61></td></tr>
</table>
</form>
</div>

<div class=div16 style='height:80;width:85%;padding:30<%=HideResults%>'>
<span class=det10><%=RequestResults%></span>
</div>

<div id=d1 style="z-index:-1;position:absolute;top:0;left:0;height:85%;margin:25 0 0 50;overflow:hidden;
 filter:progid:DXImageTransform.Microsoft.BasicImage(Opacity=0.15)">
<div id=d2 style="height:100%;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='img/wa.gif', sizingMethod='scale')"></div>
</div>

</body>
</html>