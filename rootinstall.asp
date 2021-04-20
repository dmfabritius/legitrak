<HTML>
<HEAD>
<TITLE>Installing A Root Certificate</TITLE>
<BR>Root Certificate Authority Installation
<BR>
<BR>

<%@ LANGUAGE="VBScript"%>
<%
Set fs = CreateObject("Scripting.FileSystemObject")
Set MyFile = fs.OpenTextFile("d:\websites\legitrak\wwwroot\legitrak001.base64.cer", 1)

Output = ""

Do While MyFile.AtEndOfStream <> true
  line = Chr(34) & MyFile.ReadLine & Chr(34)
  If MyFile.AtEndOfStream <> true then
    line = line & " & _" & Chr(10)
  End If
  Output = Output & line
Loop

MyFile.Close

Set MyFile = Nothing
Set fs = Nothing
%>

<SCRIPT language="VBSCRIPT">
on error resume next
Dim Str, CEnroll

Set CEnroll = CreateObject("CEnroll.CEnroll.1")
Str = <% Response.Write Output %>

CEnroll.installPKCS7(Str)

Set CEnroll = Nothing
</SCRIPT>
</HEAD>
</HTML>