<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
	Set cxnSQL = CreateObject("ADODB.Connection")
	cxnSQL.Open strConnection

' REGISTER RATINGS REPORT
	If Request.Form("RptCustom") = "14" Then
		EClts = Request.Form("CltAutoReg")
		If Len(EClts) = 0 Then Response.End
		aClts = Split(EClts,",")
		strClts = ""
		For i = 0 to UBound(aClts)
			strClts = strClts & DeCrypt(aClts(i)) & ","
		Next 'i
		strClts = Left(strClts,Len(strClts)-1)
		strSQL = _
			"UPDATE [Client List] SET [Automatic Register]=0 WHERE ClientID IN " & _
			"(SELECT ClientID FROM [Customer Clients] WHERE CustomerID=" & CustomerID & ")"
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		strSQL = _
			"UPDATE [Client List] SET [Automatic Register]=1 WHERE ClientID IN (" & strClts & ")"
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		Response.End
	End If

' CONTRIBUTION MATRIX REPORT
	If Request.Form("RptCustom") = "29" Then
        strCType = Request.Form("ContribType")
        If strCType <> "R" And strCType <> "A" Then strCType="B"
		strSQL = _
			"UPDATE [Customer List] SET ContribMatrix='" & strCType & "' WHERE CustomerID=" & CustomerID
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		EClts = Request.Form("CltCM")
		If Len(EClts) = 0 Then Response.End
		aClts = Split(EClts,",")
		strClts = ""
		For i = 0 to UBound(aClts)
			strClts = strClts & DeCrypt(aClts(i)) & ","
		Next 'i
		strClts = Left(strClts,Len(strClts)-1)
		strSQL = _
			"UPDATE [Client List] SET ContribMatrix=0 WHERE ClientID IN " & _
			"(SELECT ClientID FROM [Customer Clients] WHERE CustomerID=" & CustomerID & ")"
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		strSQL = _
			"UPDATE [Client List] SET ContribMatrix=1 WHERE ClientID IN (" & strClts & ")"
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		Response.End
	End If

' DAILY/WEEKLY REPORTS

	ClientID = Decrypt(Request.Form("RptCustom"))
	RptID = CInt(Request.Form("Rpt"))

	If RptID = 0 And ClientID = 0 Then RptID = 1 ' default daily report for customers
	If RptID = 0 And ClientID <> 0 Then RptID = 2 ' default weekly report for client lists

	Select Case RptID
		Case 17, 19: CustomID = 1
		Case 18, 20: CustomID = 2
		Case Else  : CustomID = 0
	End Select
	If Request.Form("Auto") = "0" Then
	    intAuto = 0
	Else
	    intAuto = 1
	End If
    strRptFmt = Trim(Request.Form("RptFormat"))
    If strRptFmt <> ".html" And strRptFmt <> ".htm" And strRptFmt <> ".xls" Then strRptFmt = ".doc"
	intStyle = CInt("0" & Request.Form("RptStyle"))
	If intStyle <> 0 Then
		strStyle = " [Report Style]=" & intStyle & ","
	Else
		strStyle = ""
	End If
	RptSec = Split(Request.Form("RptSec"),",")
	If Request.Form("RptCustom") = "0" Then

' SAVE CUSTOMER PREFERENCES AND CUSTOM SECTION
		strSQL = _
			"UPDATE [Customer List] SET" & _
			" [Active Report]=" & RptID & "," & _
			" [Automatic Daily]=" & intAuto & "," & _
			strStyle & _
			" [Report Format]='" & strRptFmt & "' " & _
			"WHERE CustomerID=" & CustomerID
		cxnSQL.Execute strSQL, , adExecuteNoRecords
		If CustomID <> 0 And RptID < 24 Then
			strSQL = _
				"DELETE FROM [Customer Custom Reports] WHERE" & _
				" CustomerID=" & CustomerID & " AND" & _
				" CustomID=" & CustomID
			cxnSQL.Execute strSQL, , adExecuteNoRecords
			For i = 3 to 6
				If RptSec(i) <> 0 Then
					strSQL = "INSERT INTO [Customer Custom Reports] VALUES (" & _
						CustomerID & "," & CustomID & "," & CInt(RptSec(i)) & ")"
					cxnSQL.Execute strSQL, , adExecuteNoRecords
				End If
			Next 'i
		End If
	Else

' SAVE CUSTOMER'S CLIENT PREFERENCES AND CUSTOM SECTION
		If ClientID = 0 Then Response.End

		intRptPri = CInt(Request.Form("RptPri"))
		If intRptPri >=0 or intRptPri <=4 Then
			strPri = ", [Report Priority]=" & intRptPri
		Else
			strPri = ""
		End If
		If Request.Cookies("LegiTrak")("RptCustAA") = "True" Then
			aClt = Split(Request.Form("Clients"),",")
			strSQLWhere = " IN (SELECT ClientID FROM [Customer Clients] WHERE CustomerID=" & CustomerID & ")"
			Response.Cookies("LegiTrak")("RptCustAA") = ""
		Else
			Dim aClt(1)
			aClt(1) = Request.Form("RptCustom")
			strSQLWhere = "=" & ClientID
		End If

		strSQL = _
			"UPDATE [Client List] SET" & _
			" [Active Report]=" & RptID & "," & _
			" [Automatic Weekly]=" & intAuto & "," & _
			strStyle & _
			" [Report Comments Header]='" & TweakQuote(Trim(Request.Form("SH"))) & "'," & _
			" [Report Comments]='" & TweakQuote(Trim(Request.Form("SC"))) & "'," & _
			" [Report Format]='" & strRptFmt & "'" & _
			strPri & _
			" WHERE ClientID" & strSQLWhere
		cxnSQL.Execute strSQL, , adExecuteNoRecords

		If CustomID <> 0 And RptID < 24 Then
			strSQL = _
				"DELETE FROM [Client Custom Reports] WHERE" & _
				" CustomID=" & CustomID & " AND" & _
				" ClientID" & strSQLWhere
			cxnSQL.Execute strSQL, , adExecuteNoRecords

			For i = 1 to UBound(aClt)
				ClientID = Decrypt(aClt(i))
				For j = 0 to 2
					If ClientID <> 0 And RptSec(j) <> 0 Then
						strSQL = "INSERT INTO [Client Custom Reports] VALUES (" & _
							ClientID & "," & CustomID & "," & CInt(RptSec(j)) & ")"
						cxnSQL.Execute strSQL, , adExecuteNoRecords
					End If
				Next 'j
			Next 'i
		End If

		strSQL = _
			"UPDATE [Customer List] SET" & _
			" [Report Comments Header]='" & TweakQuote(Trim(Request.Form("GH"))) & "'," & _
			" [Report Comments]='" & TweakQuote(Trim(Request.Form("GC"))) & "' " & _
			"WHERE CustomerID=" & CustomerID
		cxnSQL.Execute strSQL, , adExecuteNoRecords
	End If
	
	cxnSQL.Close
 	Set cxnSQL = Nothing

	Response.End
%>