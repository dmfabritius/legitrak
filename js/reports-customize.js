var H,E,S,F,A,R,C,RS

function init(){
	selectTab(2)
	H=SR.document.getElementsByName("H")
	E=SR.document.getElementsByName("E")
	S=SR.document.getElementsByName("S")

	F=document.getElementById("RptForm")
	A=document.getElementsByName("ActiveRpt")
	R=document.getElementsByName("Rpt")
	C=document.getElementsByName("Choices")
	RS=document.getElementsByName("RptSec")

	activeGrp=document.getElementById("ActiveGrp")
	autoGrp=document.getElementById("AutoGrp")

	q=document.location.search.substr(1)
	updateActive(null,F.RptCustom.value=((q=="")? 0 : q))
}
function updateActive(r,c){ // r=reportID, c=cust/cltID
	if (r!=null)
		mark(activeGrp) // when switching among default, cstm#1, & cstm #2, mark that field
	else
		activeGrp.style.backgroundColor=""
	if (!c) c=F.RptCustom.value
	if (c==14||c==29) { // Auto Reg & Contrib Matrix
		r=F.RptCustom.value
		c=0
	}
	a=document.getElementById("R"+c).innerHTML.split(";")
	p=a[0].split(",")
	if (r==null) r=p[2] // if not given reportID, use value from preferences
	r=(r!=0)? r : (c==0)? 1 : 2 // if reportID is zero, use default daily/weekly

	if (r<24) { // non-Excel reports can choose a style
		updateStyle(null,p[0])
		F.RptStyle.disabled=(r==14)
	} else { // Excel reports use style 5
		updateStyle(null,5)
		F.RptStyle.disabled=true
	}
	F.RptStyle.style.backgroundColor=""
	//F.RptFormat.value=p[1] //keeps overwriting rpt format when user is trying to modify active report
	F.RptFormat.style.backgroundColor=""
	if (r!=14&&r!=29) { // don't fetch preference for Auto Reg or Contrib Matrix
		F.Auto[0].disabled=F.Auto[1].disabled=false
		F.Auto[(p[3]==1)? 0 : 1].checked=true
	} else {
		F.Auto[1].checked=true
		F.Auto[0].disabled=F.Auto[1].disabled=true
	}
	autoGrp.style.backgroundColor=""

	dw=0
	document.getElementById("ApplyToAll").style.display="none"
	F.ApplyToAll.checked=false
	F.RptPri.value=0
	if (r==2||r==19||r==20||r==24) { // weekly reports
		document.getElementById("ApplyToAll").style.display=""
		C[1].style.display="none"
		C[2].style.display="none"
		C[3].style.display="none"
		H[1].style.display="none"
		H[2].style.display="none"
		RSoffset=Soffset=-1
		F.RptPri.value=p[4]
		F.SH.value=document.getElementById("H"+c).innerHTML
		F.SH.style.backgroundColor=""
		F.SC.value=document.getElementById("C"+c).innerHTML
		F.SC.style.backgroundColor=""
		updateComm()
		E[7].style.display="none"
	} else if (r==1||r==17||r==18||r==25||r==27) { // daily reports
		C[0].style.display="none"
		C[2].style.display="none"
		C[3].style.display="none"
		H[0].style.display="none"
		H[2].style.display="none"
		dw=1
		RSoffset=2
		Soffset=5
		for(i=0;i<4;i++) E[i].style.display="none"
	} else if (r==14) { // register report
		C[0].style.display="none"
		C[1].style.display="none"
		C[2].style.display=""
		C[3].style.display="none"
		H[0].style.display="none"
		H[1].style.display="none"
		H[2].style.display=""
	} else if (r==29) { // contribution matrix
		C[0].style.display="none"
		C[1].style.display="none"
		C[2].style.display="none"
		C[3].style.display=""
		H[0].style.display="none"
		H[1].style.display="none"
		H[2].style.display="none"
	}

	F.RptFormat.disabled=false
	s=null
	d=false
	if (r==1) {
		s=new Array(null,1,2,3,4)
		d=true
	} else if (r==2) {
		s=new Array(null,8,9,10)
		d=true
	} else if (r==17||r==19) {
		s=a[1].split(",")
	} else if (r==18||r==20) {
		s=a[2].split(",")
	}
	if (!s) {
		C[dw].style.display="none"
		H[dw].style.display="none"
		F.RptFormat.value='.xls '
		if (r==14) F.RptFormat.value='.html'
		if (r==14||r==29) F.RptFormat.disabled=true
		for(i=0;i<=7;i++) E[i].style.display="none"
		E[4].style.display=""
		switch(r*1){
			case 24: sec=5; break;
			case 25: sec=11; break;
			case 27: sec=12; break;
			case 29: sec=13; break;
			case 14: sec=14; break;
		}
		E[4].innerHTML=S[sec].innerHTML
	} else {
		C[dw].style.display=""
		H[dw].style.display=""
		for(i=1;i<=4;i++){
			n=(s[i]==null)? 0 : s[i]*1
			RS[i+RSoffset].value=n
			RS[i+RSoffset].disabled=false
			if (n!=0) {
				E[i+3].style.display=""
				E[i+3].innerHTML=S[RS[i+RSoffset].selectedIndex+Soffset].innerHTML
			} else
				E[i+3].style.display="none"
			RS[i+RSoffset].disabled=d
			RS[i+RSoffset].style.backgroundColor=""
		}
	}
	for (i=0;i<R.length;i++) R[i].checked=(R[i].value==r)
	for (i=0;i<A.length;i++) A[i].style.display="none"
	switch(r*1){
		case 2: case 19: case 20: A[0].style.display=""; break;
		case 14: case 24: case 29: A[1].style.display=""; break;
		case 1: case 17: case 18: A[2].style.display=""; break;
		case 25: case 27: A[3].style.display=""; break;
	}
	F.RptPri.style.backgroundColor=""
	updatePri(F.RptPri,1)
}
function updateComm(e){
	E[0].innerHTML=F.GH.value
	E[1].innerHTML=F.GC.value.substr(0,200)
	E[2].innerHTML=F.SH.value
	E[3].innerHTML=F.SC.value.substr(0,200)
	for (i=0;i<4;i++) E[i].style.display=(E[i].innerHTML=="")? "none" : ""
}
function updateSec(a,e,i){
	if (e.selectedIndex==0)
		E[i].style.display="none"
	else {
		E[i].style.display=""
		E[i].innerHTML=S[e.selectedIndex+a].innerHTML
	}
	mark(e)
}
function updateStyle(e,v){
	for (i=0;i<SR.document.styleSheets.length;i++) SR.document.styleSheets[i].disabled=true
	if (v==null) {v=e.value; mark(e)}
	v=(v==0)? 1 : v
	SR.document.styleSheets[v-1].disabled=false
	if (v!=5) F.RptStyle.value=v
}
function updateFormat(e){
	mark(e)
	f=e.value
	r=activeReport()
	if (f=='.xls ') {
		if (r==2||r==19||r==20)
			updateActive(24)
		else
			updateActive(25)
	} else {
		if (r==24)
			updateActive(2)
		else if (r==25||r==27)
			updateActive(1)
	}
	e.value=f
}
function updatePri(e,nomark){
	P=SR.document.getElementsByName("P")
	for (i=0;i<P.length;i++) {
		P[i].style.display=(e.value==0)? "" : "none"
	}
	p=(e.value!=0)? e.options[e.selectedIndex].innerHTML : "High"
	SR.document.getElementById("PriDesc").innerHTML=p+" Priority Bills"
	if (!nomark) mark(e)
}
function activeReport(){
	for (i=0;i<R.length;i++) if (R[i].checked) break
	return R[i].value*1
}
function updateCache(){
	c=F.RptCustom.value
	if (c==14||c==29) return
	rpt=activeReport()
	prefs=F.RptStyle.value+","+F.RptFormat.value+","+rpt+","+((F.Auto[0].checked)? 1 : 0)
	if (rpt==2||rpt==19||rpt==20) prefs+=","+F.RptPri.value
	cstm=0
	switch (rpt) {
		case 17: case 19: cstm=1; break;
		case 18: case 20: cstm=2; break;
	}
	if (c==0) {
		cust=document.getElementById("R0")
		a=cust.innerHTML.split(";")
		if (cstm!=0){
			a[cstm]=""
			for(i=3;i<=6;i++) a[cstm]+=","+RS[i].value
		}
		cust.innerHTML=prefs+";"+a[1]+";"+a[2]
	} else {
		for(i=0,sects="";i<=2;i++) sects+=","+RS[i].value
		if(!F.ApplyToAll.checked){
			clts=new Array(null,c)
		} else {
			clts=F.Clients.value.split(",")
			setCookie("RptCustAA","True")
		}
		for(i=1;i<clts.length;i++){
			clt=document.getElementById("R"+clts[i])
			a=clt.innerHTML.split(";")
			if (cstm!=0) a[cstm]=sects				
			clt.innerHTML=prefs+";"+a[1]+";"+a[2]
			document.getElementById("H"+clts[i]).innerHTML=F.SH.value
			document.getElementById("C"+clts[i]).innerHTML=F.SC.value
		}
	}
	F.ApplyToAll.checked=false
	
	// un-mark form fields
	for(i=0;i<F.elements.length;i++) F.elements(i).style.backgroundColor=""
	activeGrp.style.backgroundColor=""
	autoGrp.style.backgroundColor=""
}