<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' Attempt to load saved values
	BillStart = CInt("0" & Request.Cookies("LegiTrak")("BillStart"))
	BillEnd = CInt("0" & Request.Cookies("LegiTrak")("BillEnd"))
	DigestStart = CInt("0" & Request.Cookies("LegiTrak")("DigestStart"))
	DigestEnd = CInt("0" & Request.Cookies("LegiTrak")("DigestEnd"))
	Source = CInt("0" & Request.Cookies("LegiTrak")("Source"))

	KeywordClient = Decrypt(Request.Cookies("LegiTrak")("KeywordClient"))
	strQuery = Trim(Request.Cookies("LegiTrak")("SearchQuery"))
	Dim AndOrSel(2), EditionSel(2), TableSel(1), ColSel(2)
	intFilterAndOr = CInt("0" & Request.Cookies("LegiTrak")("FilterAndOr"))
	AndOrSel(intFilterAndOr) = " selected"
	intFilterEdition = CInt("0" & Request.Cookies("LegiTrak")("FilterEdition"))
	EditionSel(intFilterEdition) = " selected"
	intFilterTable = CInt("0" & Request.Cookies("LegiTrak")("FilterTable"))
	TableSel(intFilterTable) = " selected"
	intFilterCol = CInt("0" & Request.Cookies("LegiTrak")("FilterCol"))
	ColSel(intFilterCol) = " selected"
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<style>u{color:blue;cursor:pointer}</style>
<script src="js/bts.js"></script>
<script src="js/allclts-search.js"></script>
</head>

<body onload='init()' class=bkg03>

<form id=AdHocForm method=post action='allclts-search.asp' onsubmit='submitFilters()' style='margin:0'>
<table id=Bills width=100% border=0 cellspacing=4 cellpadding=0>
<col width=114 align=center><col width=125 align=center>
<tr class=hdg29>
<td>Keywords&nbsp; <input type=button value=Edit style='height:18;font-size:10;cursor:pointer' onclick='editKeywords()'></td>
<td>Combine Using</td>
<td align=center>Search string &nbsp; <input type=button value=Help style='height:18;font-size:10;cursor:help' onclick='queryHelp()'></td></tr>
<tr valign=top class=hdg29>
<td><select name=KeywordClient onchange='verifyCltKey()'>
<option value='' selected>-None-
<%
	SQLstmt = _
		"SELECT [Client List].*" & _
		" FROM [Customer Clients] INNER JOIN [Client List]" & _
		" ON [Customer Clients].ClientID = [Client List].ClientID" & _
		" WHERE [Customer Clients].CustomerID=" & CustomerID & _
		" ORDER BY [Short Company Name]"
	Set rsClients=Server.CreateObject("ador.Recordset")
	rsClients.Open SQLstmt, strConnReadOnly
		Do
			Response.Write "<option value='" & Encrypt(rsClients("ClientID")) & "'"
			If KeywordClient = rsClients("ClientID") Then
				Response.Write " selected"
			End If
			Response.Write ">" & rsClients("Short Company Name")
			rsClients.MoveNext
		Loop Until rsClients.EOF
	rsClients.Close
	Set rsClients = Nothing
%>
</select></td>
<td><select name=AndOr>
<option value=0<%=AndOrSel(0)%>>OR
<option value=1<%=AndOrSel(1)%>>AND
<option value=2<%=AndOrSel(2)%>>AND NOT
</select></td>
<td><textarea name=SearchQuery rows=2 style='width:100%' onfocus='enableOptions()' onblur='verifyQuery()'><%=strQuery%></textarea></td>
</tr>

<tr class=hdg29>
<td>Edition: <select name=SearchEdition>
<option value=0<%=EditionSel(0)%>>Both
<option value=1<%=EditionSel(1)%>>0
<option value=2<%=EditionSel(2)%>>1
</select></td>
<td>Digest: <select name=SearchTable onchange='verifySearchTable()'>
<option value=0<%=TableSel(0)%>>All
<option value=1<%=TableSel(1)%>>Current
</select></td>
<td>&nbsp; Data fields: <select name=SearchColumn>
<option value=0<%=ColSel(0)%>>Both
<option value=1<%=ColSel(1)%>>Digest Title
<option value=2<%=ColSel(2)%>>Digest Text
</select></td>
</tr>

<tr class=shd29><td colspan=3>
<span style='position:relative;top:-2'>To search the digests, click </span>
<input type=submit value=Submit>
</td></tr>

</table>
</form>

<form id=BrowseBillsForm style='margin:0'>
<%
	strIsAbout = ""

' WEIGHTED KEYWORDS
	If KeywordClient <> 0 Then
		' Load client keywords into an array
		strSQL = "SELECT Keyword, Weight FROM [Client Keywords] WHERE ClientID=" & KeywordClient
		Set rsResults = CreateObject("ADOR.Recordset")
		rsResults.Open strSQL, strConnReadOnly
		If Not rsResults.EOF Then
			aKeywords = rsResults.GetRows()
			' Build weighted keyword sub-expression
			strIsAbout = "ISABOUT("
			For i = 0 To UBound(aKeywords, 2)
				strIsAbout = strIsAbout & _
					"""" & aKeywords(0, i) & _
					""" weight(" & aKeywords(1, i) / 100 & "),"
			Next 'i
			strIsAbout = Left(strIsAbout, Len(strIsAbout) - 1) & ")"
'	    Else
'			strMessage = "No bills matching the current search criteria were found."
		End If
		rsResults.Close
		Set rsResults = Nothing
	End If

' ADVANCED SEARCH EXPRESSION
	If strQuery <> "" Then
		If strIsAbout = "" Then
			strIsAbout = strQuery
		Else
			Select Case intFilterAndOr
				Case 0: strAndOr = " OR "
				Case 1: strAndOr = " AND "
				Case 2: strAndOr = " AND NOT "
			End Select
			strIsAbout = strIsAbout & strAndOr & "(" & strQuery & ")"
		End If
	End If

	If strIsAbout <> "" Then

		If intFilterTable = 0 Then
			strTable1 = "[Supplements]"
			strTable2 = "[Supplements with Unique Bill Numbers]"
		Else
			strTable1 = "[Current Supplement]"
			strTable2 = "[Current Supplement]"
		End If
	
		Select Case intFilterCol
			Case 0: strColumn = "*"
			Case 1: strColumn = "[Long Title]"
			Case 2: strColumn = "[Body]"
		End Select

		Select Case intFilterEdition
			Case 0: strSQLWhere = ""
			Case 1: strSQLWhere = " WHERE S.Edition=0"
			Case 2: strSQLWhere = " WHERE S.Edition=1"
		End Select

	    strSQLJoin = _
   	 	"CONTAINSTABLE(" & _
			strTable1 & "," & _
    		strColumn & ",'" & _
	    	strIsAbout & "') R"
		strSQLJoin = "(" & strSQLJoin & " INNER JOIN " & strTable2 & " S" & _
			" ON R.[KEY] = S.SupplementID)"
		strSQLJoin = "(" & strSQLJoin & "  LEFT JOIN [Daily Status] D" & _
			" ON S.[Bill Number] = D.[Bill Number])"
		strSQL = "SELECT" & _
			" R.RANK," & _
			" S.[Bill Number], D.Status, D.House, D.Location, S.[Prime Sponsor], S.[Long Title]" & _
			" FROM  " & strSQLJoin & strSQLWhere & _
			" ORDER BY R.[RANK] DESC, D.[Bill Number]"
		Set rsBills = CreateObject("ADOR.Recordset")
		On Error Resume Next
		rsBills.Open strSQL, strConnReadOnly

		If rsBills.State = 0 Then ' Recordset will be open (State=1) unless there is a query syntax error
			strMessage = "Syntax error in query.  Please try again."
		ElseIf rsBills.EOF Then
			strMessage = "No bills matching the current search criteria were found."
		Else
			strMessage = ""
			Response.Write _
				"<table id=Bills border=0 cellspacing=4 cellpadding=0 width=100% class=det00>"

			Do Until rsBills.EOF
				If IsNull(rsBills("House")) Or Trim(rsBills("House")) = "" Then
					strHouse=""
				Else
					strHouse=rsBills("House") & ", "
				End If
				Response.Write _
					"<tr class=bkg04 valign=top><td class=lnk70 align=right onclick='quickAdd(" & rsBills("Bill Number") & ")'>" & _
					Replace(rsBills("Status") & rsBills("Bill Number")," ","") & "</td>" & _
					"<td>" & rsBills("Long Title") & "</td>" & _
					"<td><div style='width:80'>" & rsBills("Prime Sponsor") & "</div></td>" & _
					"<td><div style='width:110'>" & strHouse & rsBills("Location") & "</div></td>" & _
					"<td class=lnk40 onclick='lnk(arguments[0],""" & rsBills("Status") & """)'>" & _
					"<u>D</u>_<u>F</u>_<u>A</u></td></tr>"
				rsBills.MoveNext
			Loop
			Response.Write "</table>"
		End If
		rsBills.Close
		Set rsBills = Nothing
	Else
		strMessage = "No bills matching the current search criteria were found."
	End If

	Response.Write _
		"<div id=Bottom class=hdg14 style='position:relative;height:100%;margin:0 4;padding:40'>" & _
		strMessage & "</div>"
%>    
</form>
</body>
</html>
