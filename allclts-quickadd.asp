<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<!--#include virtual="includes/allclts-quickadd-post.asp"-->
<%
' If we come from the browse or search page, we'll get a bill number
	BillNumber = CInt("0" & Request.Cookies("LegiTrak")("QuickAddBill"))
	If BillNumber < 1000 Or BillNumber > 9999 Then BillNumber=""
	Response.Cookies("LegiTrak")("QuickAddBill") = ""
	
' LOAD CLIENT LIST
	strSQL = "(SELECT * FROM [Client Specific Bill Info] WHERE [Bill Number]=0" & BillNumber & ") CS"
	SQLjoin = "[Customer Clients] CC"
	SQLjoin = "(" & SQLjoin & " INNER JOIN [Client List] CL ON CC.ClientID = CL.ClientID)"
	SQLjoin = "(" & SQLjoin & " LEFT JOIN " & strSQL & " ON CC.ClientID = CS.ClientID) "
	strSQL = _
		"SELECT" & _
		" CL.ClientID, CL.[Client Company Name], CS.Notes," & _
		" ISNULL(CS.PriorityNum,0) AS PriorityNum," & _
		" ISNULL(CS.PositionNum,0) AS PositionNum " & _
		"FROM " & SQLjoin & _
		"WHERE CC.CustomerID=" & CustomerID & _
		" ORDER BY CL.[Short Company Name]"
	Set rsClients=Server.CreateObject("ADOR.Recordset")
	rsClients.Open strSQL, strConnReadOnly
	aClients = rsClients.GetRows()
	intClientCount = UBound(aClients,2)
	rsClients.Close
	Set rsClients = Nothing

' LOAD BILL NUMBER FROM BROWSE PAGE IF WE CAME FROM THERE
	If BillNumber <> "" Then
		strSQL = "SELECT Title, Companion FROM [Daily Status] WHERE [Bill Number]=" & BillNumber
		Set rsDailyStatus=Server.CreateObject("ador.Recordset")
		rsDailyStatus.Open strSQL, strConnReadOnly
		strTitle = rsDailyStatus("Title")
		If Not IsNull(rsDailyStatus("Companion")) Then
			strCompBill = rsDailyStatus("Companion")
			If strCompBill >= 5000 Then
				strChamber = "SB "
			Else
				strChamber = "HB "
			End If
		Else
			strCompBill = ""
		End If
		rsDailyStatus.Close
		Set rsDailyStatus = Nothing
	Else
		BillNumber=""
	End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
var cltCount=<%=intClientCount%>
var bdf
var h=parent.data.location.href.substr(parent.data.location.href.lastIndexOf("/")+1)

function init(){<%
	If strRedirect <> "" Then
		Response.Write strRedirect
	Else
		Response.Write _
			"selectTab(0);" & _
			"bdf=document.getElementById('BillDetailForm');" & _
			"bdf.Title.value=""" & strTitle & """;" & _
			"bdf.digCom.value=' Load\nDigest';" & _
			"if (bdf.Bill.value=='') bdf.Bill.focus(); else bdf.Comments.focus();"
	End If
%>}
function cancelAdd() {
	if (h!="allclts-quickadd.asp") {
		if (h=="allclts-tracking.asp") selectTab(1)
		if (h=="allclts-browse.asp") selectTab(2)
		if (h=="allclts-search.asp") selectTab(3)
		parent.document.getElementById("fs").rows="20,*,0"
	} else {
		bdf.reset()
		document.body.scrollTop=0
		bdf.Bill.style.backgroundColor="#FFFFFF"
		bdf.Bill.focus()
	}
}
function verifyBill(){
	i=c=0
	while ((o=document.getElementsByName("Clt_"+i++)).length!=0) c+=o[0].checked? 1 : 0
	if (c==0 && !confirm("No tracking lists have been selected.  Click OK if you want to abandon your changes.")) return false
	
	if (bdf.Bill.value.length==0||!isBill(bdf.Bill)){
		alert("Please enter a bill number.")
		document.body.scrollTop=0
		bdf.Bill.style.backgroundColor="#FFFFFF"
		bdf.Bill.focus()
		return false
	} else {
		if (h!="allclts-quickadd.asp") {
			if (h=="allclts-tracking.asp") {
				bdf.action="allclts-tracking.asp"
				bdf.target="data"
			} else {
				if (h=="allclts-browse.asp") selectTab(2)
				if (h=="allclts-search.asp") selectTab(3)
				parent.document.getElementById("fs").rows="20,*,0"
				bdf.action="allclts-quickadd-post.asp"
			}
		}	
		return true
	}
}
</script>
</head>
<body class=bkg03 onload='init()'>
<form id=BillDetailForm action="allclts-quickadd.asp" method=post onsubmit='return verifyBill()'>
<input type=hidden name=QuickAdd value=True>
<input type=hidden name=ClientCount value=<%=intClientCount%>>
<table id=Bills width=100% border=0 cellspacing=2 cellpadding=0 class=hdg29 style='padding-left:10;border:4px solid #FFFFFF;border-bottom-width:0'>
<col width=90 align=right>
<tr><td><b>Bill Number:</b></td><td>
<input name=Bill type=text size=4 maxlength=4 style='font-weight:bold' value='<%=BillNumber%>' onchange='return isBill(this)'>
<input name=Title type=text readonly tabindex=50 class=hdg29 style='width:300;border-width:0;margin:0 0 2 5'>
</td></tr><tr><td valign=top><b>Comments:</b>
<iframe id=DigestFrame style='display:none'></iframe>
<br><input name=digCom type=button onclick='dig2Com()' style='font-size:10px;width:40px;height:35px;margin-right:13px'></td>
<td><textarea name=Comments rows=4 style='width:100%'></textarea></td>
</tr></table>

<table width=100% border=0 cellspacing=4 cellpadding=0 class=det00>
<col align=center>
<tr class=hdg29><td>Sel</td><td>Client Name</td><td>Notes</td><td>Priority</td><td>Position</td></tr>
<%

	DefPri = Request.Cookies("LegiTrak")("DefPriority")
	DefPos = Request.Cookies("LegiTrak")("DefPosition")
	
	Dim thePri(4), thePos(6)
	For i = 0 to intClientCount
		If aClients(3,i) <> 0 Then
			thePri(aClients(3,i)) = " selected"
			thePos(aClients(4,i)) = " selected"
		Else
			thePri(DefPri) = " selected"
			thePos(DefPos) = " selected"
		End If
	
		Response.Write "<tr class=bkg04>"
		Response.Write "<td style='padding:0 5'><input type=checkbox name=Clt_" & i & " value=True>"
		Response.Write "<input type=hidden name=CltNum_" & i & " value=" & Encrypt(aClients(0,i)) & "></td>"
		Response.Write "<td>" & aClients(1,i)
		If aClients(3,i) <> 0 Then
			Response.Write "&nbsp; <b>(tracking)</b>"
		End If		
		Response.Write "</td>"
		Response.Write "<td><input name=Notes_" & i & " type=text maxlength=20 value='" & aClients(2,i) & "'></td>"
		Response.Write "<td><select name=Pri_" & i & ">" & _
			"<option value=1" & thePri(1) & ">High" & _
			"<option value=2" & thePri(2) & ">Medium" & _
			"<option value=3" & thePri(3) & ">Low" & _
			"<option value=4" & thePri(4) & ">TBD</select></td>"
		Response.Write "<td><select name=Pos_" & i & " style='width:83'>" & _
			"<option value=1" & thePos(1) & ">Support" & _
			"<option value=2" & thePos(2) & ">Oppose" & _
			"<option value=3" & thePos(3) & ">Concerns" & _
			"<option value=4" & thePos(4) & ">Neutral" & _
			"<option value=5" & thePos(5) & ">Monitor" & _
			"<option value=6" & thePos(6) & ">-Blank-</select></td>"
		Response.Write "</tr>"

		thePri(aClients(3,i)) = ""
		thePos(aClients(4,i)) = ""
		thePri(DefPri) = ""
		thePos(DefPos) = ""
	Next
%>
<tr class=bkg04 valign=middle style='height:35'><td colspan=5 align=center>
<input type=submit value=Submit>
<span style='width:200;text-align:right'><input type=button value="Cancel" onclick='cancelAdd()'></span>
</td></tr>
</table>

<div class=bkg04 style='position:relative;height:100%;margin:0 4;padding:10 6'>
<%
	If strCompBill <> "" Then Response.Write _
		"<table class=det00><tr valign=top>" & _
		"<td><input type=checkbox name=Companion value=" & strCompBill & "></td>" & _
		"<td>Also track the companion bill, " & strChamber & strCompBill & _
		", applying the same attributes and comments.</td>" & _
		"</tr></table>"
%>
</div>

</form>
</body>
</html>
