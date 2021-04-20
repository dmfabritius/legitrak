<%
	Response.Buffer = True
	Response.CacheControl = "no-cache"
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","no-cache"
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1

	strConnection ="PROVIDER=SQLOLEDB;SERVER=.\SQLEXPRESS;DATABASE=btsdb;UID=xxx;PASSWORD=xxx"
	strConnReadOnly ="PROVIDER=SQLOLEDB;SERVER=.\SQLEXPRESS;DATABASE=btsdb;UID=xxx;PASSWORD=xxx"

	' ADO constants
	Const adOpenStatic = 3		' This returns an unchanging "snapshot" of the data
	Const adOpenDynamic = 2		' This lets us see changes made by other users
	Const adLockPessimistic = 2	' This lets us modify the data
	Const adUseClient = 3		' This lets us see the Recordcount
	Const adExecuteNoRecords = 128	' Indicates that we don't want any records returned

	Function Encrypt(ByVal inp)
		Encrypt = Hex(inp*7817)
	End Function

	Function Decrypt(ByVal inp)
		Dim r, h, d

		Decrypt = 0
		If Len(inp) <> 0 Then
			inp = CDbl("&h" & Trim(inp))
			d = inp/7817
  			If d = Int(d) Then Decrypt = d
		End If
	End Function

	Function RevStr(ByVal s)
		Dim i

		RevStr = ""
		For i = Len(s) To 1 Step -1
			RevStr = RevStr & Mid(s,i,1)
		Next 'i
	End Function
	
	Function MakeHTML(ByVal inp)
'		MakeHTML=Replace(Replace(Replace(Replace(inp,"&","&amp;"),"<","&lt;"),">","&gt;"),vbCrLf,"<br>")
    	MakeHTML=""
		If Not IsNull(inp) Then MakeHTML=Replace(Replace(inp,"&","&amp;"),vbCrLf,"<br>")
	End Function
	
	Function TweakQuote(ByVal inp)
	    TweakQuote=""
	    If Not IsNull(inp) Then TweakQuote=Replace(inp,"'","�")
	End Function

	' This is the customer who's signed in
	' and the client they're working with
	CustomerID = Decrypt(Request.Cookies("LegiTrak")("CustomerID"))
	ClientID = Decrypt(Request.Cookies("LegiTrak")("ClientID"))
	CustomerName = Request.Cookies("LegiTrak")("CustomerName")
%>