<%@ Language="VBScript" %><!--#include virtual="includes/common.asp"-->
<%
	If CustomerID <> 1 And CustomerID <> 267 Then Response.Redirect "errors/403-17.htm"

	If Request.Form("PDCverify") = "True" Then

	    Set cxnSQL = CreateObject("ADODB.Connection")
          cxnSQL.ConnectionTimeout = 600
	    cxnSQL.Open strConnection

	    aPDCnew = Split(Request.Form("PDCnew"),",")
	    aPDCname = Split(Request.Form("PDCName"),",")
	    aPolID = Split(Request.Form("PolID"),",")
	    PolIDCnt = UBound(aPolID)

        strDelID = ""
	    For i = 0 to PolIDCnt
            If Len(Trim(aPDCName(i))) = 0 Then
                strDelID = strDelID & "," & aPolID(i)
            Else
                strCommand = ""
                If aPDCnew(i) = 1 Then
		        strCommand = _
			        " LastName=SUBSTRING(LastName,1,1)+LOWER(SUBSTRING(LastName,2,LEN(LastName)-1))," & _
			        " FirstName=SUBSTRING(FirstName,1,1)+LOWER(SUBSTRING(FirstName,2,LEN(FirstName)-1)),"
                End If
		        strCommand = _
			        "UPDATE [Politicians] SET" & strCommand & _
			        " PDCName='" & Trim(Replace(aPDCName(i),"'","''")) & "'," & _
			        " PDCverified=1" & _
			        " WHERE PoliticianID=" & aPolID(i)
		        cxnSQL.Execute strCommand, , adExecuteNoRecords
            End If

	    Next 'i

        If strDelID <> "" Then
            strCommand = "DELETE FROM Politicians WHERE PoliticianID IN (" & Mid(strDelID,2) & ")"
            cxnSQL.Execute strCommand, , adExecuteNoRecords
        End If
	
    End If
%>
<html>
<head>
<link rel=stylesheet href="styles.css" type="text/css">
<script src="js/bts.js"></script>
<script>
function okay(){
    return confirm("All the records with a blank PDC Name will be deleted, the remaining records will be marked as verified. Also note, if there are a lot of records this will take a *LONG TIME*, so please be patient. Are you sure?")
}
</script>
</head>
<body onload='selectTab(5)' class=bkg03>
<form id=PDCForm method=post action='maint-pdc.asp'>
<br>
<div style="text-align:center;margin:10px 20px">
<span class=det00>Cut and paste PDC Name from duplicate/new records (those with UPPERCASE NAMES) to existing Politician records.<br>
Leave the PDC Name blank to delete that record from the database.</span>
<br><br>
<input type=submit value='Submit' onclick='return okay()'>
<input type=hidden name=PDCverify value=True>
</div><br>
<%
    strSQL = "(SELECT PoliticianID, MAX([Year]) AS [Year], Party FROM [Candidate Details] GROUP BY PoliticianID, Party) CD"
    strSQL2 = "SELECT DISTINCT LastName FROM Politicians WHERE ISNULL(PDCVerified, 0)=0"

	strSQL = _
	    "SELECT P.PoliticianID, P.LastName, P.FirstName, P.PDCname, P.PDCverified, CD.Party, ISNULL(CD.[Year],9999) [Year] " & _
        "FROM Politicians P LEFT OUTER JOIN " & strSQL & " ON P.PoliticianID = CD.PoliticianID " & _
        "WHERE ISNULL(P.PDCVerified,0)=0 OR P.LastName IN (" & strSQL2 & ")" & _
        "ORDER BY P.LastName, P.FirstName"
	Set rsPDC=Server.CreateObject("ADOR.Recordset")
	rsPDC.Open strSQL, strConnection
	
	If Not rsPDC.EOF Then
		Response.Write _
			"<table width=100% cellspacing=4 cellpadding=0 class=det00 style='cursor:default'>" & _
			"<col width=200px><col width=250px>"
		Response.Write _
            "<tr class=hdg24><td>Politician record</td><td>PDC Name</td><td>Verified</td><td colspan=10>&nbsp;</td></tr>"
		Do Until rsPDC.EOF

            PDCnew=0
            If rsPDC("Lastname") = UCASE(rsPDC("Lastname")) Then PDCnew = 1
			Response.Write _
			    "<tr valign=top class=bkg04>" & _
			    "<td>" & rsPDC("Lastname") & ", " & rsPDC("Firstname") & " (" & rsPDC("Party") & ") " & rsPDC("Year") & "</td>" & _
			    "<td><input name=PolID type=hidden value=" & rsPDC("PoliticianID") & "><input name=PDCname value=""" & rsPDC("PDCName") & """ style='width:100%'></td>" & _
			    "<td>&nbsp;<input name=PDCnew type=hidden value=" & PDCnew & ">" & rsPDC("PDCverified") & "</td>" & _
			    "</tr>"
			rsPDC.MoveNext
		Loop
		Response.Write "</table>"
	Else
		Response.Write "No unverified politicians."
	End If
	rsPDC.Close
	Set rsPDC = Nothing
%>
</form>
</body>
</html>
