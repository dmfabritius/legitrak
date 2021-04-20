var ahf
function init(){
	selectTab(3)
	parent.document.getElementById("fs").rows="20,*,0"
	document.body.scrollTop=0
	ahf=document.getElementById("AdHocForm")
	verifyQuery()
	verifySearchTable()
}
function quickAdd(b){
	setCookie("QuickAddBill",b)
	parent.qa.location.href="allclts-quickadd.asp"
	parent.document.getElementById("fs").rows="20,0,*"
}
function submitFilters(){
	setCookie("KeywordClient",ahf.KeywordClient.value)
	setCookie("SearchQuery",ahf.SearchQuery.innerHTML)
	setCookie("FilterAndOr",ahf.AndOr.selectedIndex)
	if (ahf.SearchTable.selectedIndex==0)
		setCookie("FilterEdition",ahf.SearchEdition.selectedIndex)
	else
		setCookie("FilterEdition",0)
	setCookie("FilterTable",ahf.SearchTable.selectedIndex)
	setCookie("FilterCol",ahf.SearchColumn.selectedIndex)
}
function enableOptions(){
	if (ahf.KeywordClient.value==0) return true
	ahf.AndOr.disabled=false
}
function disableOptions(){
	ahf.AndOr.disabled=true
}
function verifyCltKey(){
	if (ahf.KeywordClient.value==0)
		disableOptions()
	else
		verifyQuery()
}
function verifyQuery(){
	if (ahf.SearchQuery.innerHTML.trim() != '')
		enableOptions()
	else {
		disableOptions()
	}
}
function verifySearchTable(){
	if (ahf.SearchTable.value==0)
		ahf.SearchEdition.disabled=false
	else
		ahf.SearchEdition.disabled=true
}
function queryHelp(){
	window.open("query-help.htm",null,
		"top=25,left=25,height=600,width=900,status=yes,"+
		"toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes")
}
function editKeywords(){
	if (ahf.KeywordClient.value==0)
		alert("Please select a client to edit their keywords.")
	else {
		k=ahf.KeywordClient
		name=k.options[k.selectedIndex].innerHTML.trim().replace(/<BR>|<br>/,"")
		c=top.contents.document.getElementById("ClientMenu")
		for(i=0;i<c.rows.length;i+=2) if (c.rows[i].cells[0].childNodes[0].childNodes[0].innerHTML==name) break
		menuSelect(null,"client-keywords.htm",i)
	}
}