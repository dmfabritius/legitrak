<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' KEYWORDS
	strSQL = "SELECT * FROM [Client Keywords] WHERE ClientID=" & ClientID
	Set rsKeywords=Server.CreateObject("ADOR.Recordset")
	rsKeywords.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
	Const MAXKEYWORDS = 20
	Dim keywords(20), weights(20)
	
	If Request.Form("UpdateKeywords") = "True" Then
		Do Until rsKeywords.EOF
			rsKeywords.Delete
			rsKeywords.Update
			rsKeywords.MoveNext
		Loop
		NewKeyword = False
		For i = 1 to MAXKEYWORDS
			keywords(i) = Trim(Request.Form("keyword" & i))
			weights(i)  = Request.Form("weight"  & i)
			If Len(weights(i)) = 0 Then weights(i) = 0
			If Not IsNumeric(weights(i)) Then weights(i) = 0
			If keywords(i) <> "" and weights(i) <> 0 Then
				NewKeyword = True
				rsKeywords.AddNew
				rsKeywords("ClientID") = ClientID
				rsKeywords("Keyword") = keywords(i)
				rsKeywords("Weight") = weights(i)
				rsKeywords.Update
				rsKeywords.MoveNext
			End If
			keywords(i) = ""
			weights(i) = ""
		Next
		If NewKeyword Then rsKeywords.MoveFirst
	End If

	' Load the client's keywords
	i = 0
	Do Until rsKeywords.EOF or i = MAXKEYWORDS
		i = i + 1
		keywords(i) = rsKeywords("Keyword")
		weights(i) = rsKeywords("Weight")
		rsKeywords.MoveNext
	Loop
	KeywordCount = i

	rsKeywords.Close
	Set rsKeywords = Nothing
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function verifyWeight(e,i){
	var o=document.getElementsByName("weight"+i)[0]
	var k=parseInt(o.value)
	if (isNaN(k)) k=-1
	if (k<0 || k>100){
		w()
		o.value=0
		o.select()
		return false
	} else {
		mark(e)
		return true
	}
}
function w(){
	alert(
		"Please use weight values between 1 and 100.\n"+
		"You can use a weight of zero to delete a keyword."
	)
}
function searchDigests(){
	setCookie("KeywordClient","<%=Encrypt(ClientID)%>")
	selectMenu("mnu3","allclts-search.htm")
}
</script>
</head>
<body onload='selectTab(1)' class=bkg04 style='margin:10 0 0 20'>
<form action="client-keywords.asp" method=post>
<input type="hidden" name=UpdateKeywords value=True>
<table border=0 cellpadding=0 cellspacing=1 style='padding-left:5'>
<tr align=center class=hdg24>
<td>Keyword</td><td><span style='cursor:help' onclick='w()'>Weight</span></td>
<td></td>
<td>Keyword</td><td><span style='cursor:help' onclick='w()'>Weight</span></td>
</tr>
<%
	For i = 0 to (MAXKEYWORDS/2)-1
		Response.Write "<tr>"
		For j = 0 to 1
			Response.Write _
				"<td>" & _
				"<input size=20 type=text onchange='mark(this)'" & _
				" name=keyword" & i+j*(MAXKEYWORDS/2)+1 & _
				" tabindex=" & 2*(i+j*(MAXKEYWORDS/2))+1 & _
				" value='" & Replace(keywords(i+j*(MAXKEYWORDS/2)+1),"'","&#39") & "'>" & _
				"</td>"
			Response.Write _
				"<td>" & _
				"<input size=4 type=text onchange='return verifyWeight(this," & i+j*(MAXKEYWORDS/2)+1 & ")'" & _
				" name=weight" & i+j*(MAXKEYWORDS/2)+1 & _
				" tabindex=" & 2*(i+j*(MAXKEYWORDS/2)+1) & _
				" value=" & weights(i+j*(MAXKEYWORDS/2)+1) & ">" & _
				"</td>"
			If j = 0 Then
				Response.Write "<td width=25></td>"
			End If
		Next
		Response.Write "</tr>"
	Next
%>
<tr><td colspan=5 height=55 valign=bottom align=center>
<div style='height:25'><input type=submit style='width:160' value="Update Keywords"></div>
<input type=button style='width:160' value="Search Digests" onclick='searchDigests()'>
</td></tr>
</table>
</form>
</body>
</html>
