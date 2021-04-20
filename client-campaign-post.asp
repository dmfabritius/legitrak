<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<!--#include virtual="includes/security-check.asp"-->
<%
' UPDATE CLIENT CANDIDATE COMMENTS
	If Request.Form("UpdateCand") = "True" Then
		CandidateID = CInt(Request.Form("CandID"))
		PoliticianID = CInt(Request.Form("PolID"))

		' We either want to add/update the comments for a single client, or all of the
		' clients for a customer.  The easiest way I could think to do this was to use
		' an "outer loop" query to drive the list of client ID's.
		If Request.Form("ApplyToAll") = "All" Then
			strSQL = "SELECT ClientID FROM [Customer Clients] WHERE CustomerID=" & CustomerID
		Else
			strSQL = "SELECT ClientID FROM [Client List] WHERE ClientID=" & ClientID
		End If
		Set rsClients=Server.CreateObject("ADOR.Recordset")
		rsClients.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		Do
			strSQL = _
				"SELECT * FROM [Client Politician Comments] " & _
				"WHERE" & _
				" [ClientID]=" & rsClients("ClientID") & " AND" & _
				" [PoliticianID]=" & PoliticianID
			Set rsContribCom=Server.CreateObject("ADOR.Recordset")
			rsContribCom.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
			If rsContribCom.EOF Then
				rsContribCom.AddNew
				rsContribCom("ClientID") = rsClients("ClientID")
				rsContribCom("PoliticianID") = PoliticianID
			End If
			rsContribCom("Comments") = Request.Form("Com")

			' For the client we're really working with, update the other fields
			If rsClients("ClientID") = ClientID Then
				If Trim(Request.Form("PriRec")) = "" Then
					rsContribCom("Primary Rec") = NULL
				Else
					rsContribCom("Primary Rec") = CInt(Request.Form("PriRec"))
				End If
				If Trim(Request.Form("GenRec")) = "" Then
					rsContribCom("General Rec") = NULL
				Else
					rsContribCom("General Rec") = CInt(Request.Form("GenRec"))
				End If
				If Trim(Request.Form("Group")) = "" Then
					rsContribCom("Group") = NULL
				Else
					rsContribCom("Group") = CInt(Request.Form("Group"))
				End If
			End If
			
			rsContribCom.Update
			rsContribCom.Close
			rsClients.MoveNext
		Loop Until rsClients.EOF
		rsClients.Close
		Set rsClients = Nothing
		Set rsContribCom = Nothing

' UPDATE CLIENT CANDIDATE CONTRIBUTIONS
		strSQL = _
			"SELECT * FROM [Contributions] " & _
			"WHERE ClientID=" & ClientID & " AND CandidateID=" & CandidateID
		Set rsContribs=Server.CreateObject("ADOR.Recordset")
		rsContribs.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		Do Until rsContribs.EOF
			rsContribs.Delete
			rsContribs.Update
			rsContribs.MoveNext
		Loop

		C = Split(Request.Form("C"),",")
		For i = 0 to 5
			If i < 3 Then
				intPrimary = 1
			Else
				intPrimary = 0
			End If
			dtDate = C(i*2)
			intAmount  = C(i*2+1)
			If Len(intAmount) = 0 Then intAmount = 0
			If Not IsNumeric(intAmount) Then intAmount = 0
			If IsDate(dtDate) And intAmount <> 0 Then
				rsContribs.AddNew
				rsContribs("ClientID") = ClientID
				rsContribs("CandidateID") = CandidateID
				rsContribs("Date") = dtDate
				rsContribs("Amount") = intAmount
				rsContribs("Primary") = intPrimary
				rsContribs.Update
				rsContribs.MoveNext
			End If
		Next
	Else
		Response.Cookies("LegiTrak")("scrollTop") = ""
	End If


	If Request.Form("UpdateMult") = "True" Then
		aIDs = Split(Request.Form("CandsToUpdate"),";")

' ADD/UPDATE MULTIPLE CANDIDATE (POLITICIAN) COMMENTS
		strPriRec = Request.Form("PriRec")
		If Not IsNumeric(strPriRec) Then
			intPriRec = NULL
		Else
			intPriRec = CInt(strPriRec)
		End If
		strGenRec = Request.Form("GenRec")
		If Not IsNumeric(strGenRec) Then
			intGenRec = NULL
		Else
			intGenRec = CInt(strGenRec)
		End If
		strGroup = Request.Form("Group")
		If Not IsNumeric(strGroup) Then
			intGroup = NULL
		Else
			intGroup = CInt(strGroup)
		End If

		aPolIDs = Split(aIDs(0),",")
		strWhere = ""
		For i = 1 to UBound(aPolIDs)
    		If IsNumeric(aPolIDs(i)) Then
			    strWhere = strWhere & " OR P.PoliticianID=" & aPolIDs(i)
			End If
		Next 'i
		strWhere = Right(strWhere,Len(strWhere)-4)

		strSQL = "(SELECT * FROM [Client Politician Comments] WHERE ClientID=" & ClientID & ") CC"
		strJoin = "Politicians P LEFT JOIN " & strSQL & " ON P.PoliticianID=CC.PoliticianID"
		strSQL = "SELECT P.PoliticianID PolID, CC.* FROM " & strJoin &	" WHERE " & strWhere
		Set rsCC=Server.CreateObject("ADOR.Recordset")
		rsCC.Open strSQL, strConnection, adOpenDynamic, adLockPessimistic
		Do Until rsCC.EOF
			If IsNull(rsCC("ClientID")) Then
				PolID = rsCC("PolID")
				rsCC.AddNew
				rsCC("ClientID") = ClientID
				rsCC("PoliticianID") = PolID
			End If
			If strPriRec <> "" Then rsCC("Primary Rec") = intPriRec
			If strGenRec <> "" Then rsCC("General Rec") = intGenRec
			If strGroup <> "" Then rsCC("Group") = intGroup
			rsCC.Update
			rsCC.MoveNext
		Loop
		rsCC.Close
		Set rsCC = Nothing

' ADD CONTRIBUTIONS FOR MULTIPLE CANDIDATES
		strSQL = "INSERT INTO Contributions (ClientID,[Date],Amount,[Primary],CandidateID) VALUES (" & ClientID & ","
		
		Set cmdSQL = CreateObject("ADODB.Connection")
		cmdSQL.Open strConnection

        dtPriDate = CDate(Request.Form("PriDate"))
        intPriAmt = CInt(Request.Form("PriAmt"))
		aCandPri = Split(aIDs(1),",")
		strPri = "'" & dtPriDate & "'," & intPriAmt & ",1,"
		For i = 1 to UBound(aCandPri)
		    If IsNumeric(aCandPri(i)) Then
    			strCommand = strSQL & strPri & aCandPri(i) & ")"
	    		cmdSQL.Execute strCommand, , adExecuteNoRecords
	    	End If
		Next 'i

        dtGenDate = CDate(Request.Form("GenDate"))
        intGenAmt = CInt(Request.Form("GenAmt"))
		aCandGen = Split(aIDs(2),",")
		strGen = "'" & dtGenDate & "'," & intGenAmt & ",0,"
		For i = 1 to UBound(aCandGen)
		    If IsNumeric(aCandGen(i)) Then
			    strCommand = strSQL & strGen & aCandGen(i) & ")"
			    cmdSQL.Execute strCommand, , adExecuteNoRecords
			End If
		Next 'i

		Set cmdSQL = Nothing

	End If

	Response.End
%>
