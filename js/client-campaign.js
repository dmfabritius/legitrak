var detailActive = 0
var csf,cdf,cdd,mdd

function init(){
	csf=document.getElementById("CandSummaryForm")
	cdf=document.getElementById("CandDetailForm")
	cdd=document.getElementById("CandDetails")
	mdd=document.getElementById("MultDetails")
	
	selectTab(3)
	var scrollTop = getCookie("scrollTop")
	if (scrollTop!="") document.body.scrollTop=scrollTop
}
function scr(){
	s=cdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-210
	s=mdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-155
}
function sortBy(f){
	o=""
	if (f=="Dst") o="ISNULL(C.DistrictID,99),"
	if (f=="Grp") o="ISNULL(CC.[Group],32767),"
	if (f=="Rec") o="CASE WHEN CC.TotRec=0 THEN 99999 ELSE CC.TotRec END,"
	if (f=="Act") o="CASE WHEN ISNULL(AC.Actual,0)=0 THEN 99999 ELSE AC.Actual END,"
	setCookie("CandOrderField",o+"C.[LastName], C.[FirstName]")
	window.location.href="client-campaign.asp"
}
function updateFilter(FilterByRace){
	setCookie("FilterName",(FilterByRace)? 0 : csf.FilterName.value)
	setCookie("FilterYear",csf.FilterYear.value)
	setCookie("FilterRace",csf.FilterRace.value)
	setCookie("FilterStatus",(FilterByRace)? 0 : csf.FilterStatus.value)
	setCookie("FilterParty",csf.FilterParty.value)
	setCookie("FilterDst",csf.FilterDst.value)
	setCookie("FilterGrp",csf.FilterGrp.value)
	setCookie("FilterRec",csf.FilterRec.value)
	setCookie("FilterAct",csf.FilterAct.value)
	setCookie("scrollTop","")
	window.location.href="client-campaign.asp"
}
function hideDetail(e,override) {
	e=(!e) ? event.srcElement : e.target
	while (e.parentNode!=null && !/Details/.test(e.id)) e=e.parentNode
	if (/Details/.test(e.id) && !override) return
	if (detailActive==-1)
		detailActive=1;
	else if (detailActive) {
		cdd.style.display="none"
		mdd.style.display="none"
	}
}
function selectDetail(i,c,p){
	detailActive=-1
	setCookie("scrollTop",document.body.scrollTop)
	cdf.Index.value=i
	cdf.CandID.value=c
	cdf.PolID.value=p
	nam=document.getElementById("nam"+i)
	cdf.CandName.value=nam.innerHTML+((nam.style.color=="gray")? " - Withdrawn" : "")
	cdf.Com.value=document.getElementById("com"+i).innerHTML.fromHTML().trim()
	amts=document.getElementById("amt"+i).innerHTML.split(",")
	C=document.getElementsByName("C")
	for (a=0;a<12;a++) C[a].value=amts[a]
	cdf.PriRec.value=amts[12]
	cdf.GenRec.value=amts[13]
	cdf.ElecYear.value=amts[14]
	cdf.Group.value=document.getElementById("grp"+i).innerHTML.toLowerCase().replace(/\&nbsp;/,"")
	cdd.style.top=document.body.clientHeight+document.body.scrollTop-220
	cdd.style.display = "block"
	cdf.Com.focus()
}
function submitDetail(){
	i=cdf.Index.value
	cdf.Com.value=cdf.Com.value.trim()
	document.getElementById("com"+i).innerHTML=cdf.Com.value.toHTML()
	C=document.getElementsByName("C")
	for (a=0,act=0,amts="";a<6;a++) {
		n=parseInt(C[a*2+1].value)
		if (!isNaN(n) && n!=0) {
			amts+=C[a*2].value+","+n+","
			act+=n
		} else
			amts+=",,"
	}
	document.getElementById("amt"+i).innerHTML=amts+cdf.PriRec.value+","+cdf.GenRec.value+","+cdf.ElecYear.value
	document.getElementById("act"+i).innerHTML=(act!=0) ? act : "&nbsp;"
	document.getElementById("rec"+i).innerHTML=((r=cdf.PriRec.value*1+cdf.GenRec.value*1)!=0) ? r : "&nbsp;"
	document.getElementById("grp"+i).innerHTML=(cdf.Group.value.trim()=="")? "&nbsp;" : cdf.Group.value
	cdd.style.display="none"
}

function toggleDet(i,b){
	s=document.getElementById("s"+i).cells[1].childNodes[0].style
	d=document.getElementById("d"+i).style
	cd=document.getElementById("com"+i)
	c=cd.style
	if (b==null) b=(!/b1/.test(s.backgroundImage)) ? 1 : 0
	s.backgroundImage = "url(img/b"+b+".gif)"
	if (b==1) {
		d.display=""
		c.overflow="visible"
		c.height="100%"
		cd.parentNode.rowSpan=2
	} else {
		d.display="none"
		c.overflow="hidden"
		c.height="14px"
		cd.parentNode.rowSpan=1
	}
}
function allDet(o){
	o.style.backgroundImage = "url(img/b"+(b=(!/b1/.test(o.style.backgroundImage)) ? 1 : 0)+".gif)"
	for (i=0;i<csf.CandCount.value;i++) toggleDet(i,b)
}
function gotoRace(t,r){
	o=csf.FilterRace
	if (t==1) for(i=0;i<o.options.length;i++) if (o.options[i].text==r) f=o.options[i].value
	if (t==2) for(i=0;i<o.options.length;i++) if (o.options[i].id==r) f=o.options[i].value
	setCookie("FilterRace",f)
	setCookie("FilterName",0)
	setCookie("scrollTop","")
	window.location.href="client-campaign.asp"
}
function gotoDistrict(d){
	setCookie("FilterDst",d)
	setCookie("FilterName",0)
	setCookie("scrollTop","")
	window.location.href="client-campaign.asp"
}
function selectCands(e){
	var c=null
	if (e.title=="Select All"){
		if (confirm("Would you like to select all of the candidates?\n(Withdrawn candidates will be skiped.)")){
			e.title="Un-Select All"
			c=true
		}
	} else {
		if (confirm("Would you like to un-select all of the candidates?\n(Withdrawn candidates will be skiped.)")){
			e.title="Select All"
			c=false
		}
	}
	if (c!=null)
		for (i=0;i<csf.CandCount.value;i++) {
			o=document.getElementById("chk"+i)
			if (o.parentNode.parentNode.cells[1].childNodes[1].style.color=="") o.checked=c
		}
}
function selectMult(){
	detailActive=-1
	for (i=0;i<csf.CandCount.value;i++)	if (document.getElementById("chk"+i).checked) break
	if (i==csf.CandCount.value){
		alert("To update the attributes of multiple candidates, first select them using the checkboxes.")
		return
	}
	mdd.style.top=document.body.clientHeight+document.body.scrollTop-155
	mdd.style.display="block"
}
function submitMult(){
	x =(pr=csf.PriRec.value)==""
	x&=(gr=csf.GenRec.value)==""
	x&=(gp=csf.Group.value)==""
	x&=(pd=csf.PriDate.value)==""
	x&=(pa=csf.PriAmt.value)=="" 
	x&=(gd=csf.GenDate.value)==""
	x&=(ga=csf.GenAmt.value)==""
	if (x) alert("No changes requested.")
	else {
		for(i=0,p="",cp="",cg="";i<csf.CandCount.value;i++)
			if(document.getElementById("chk"+i).checked) {
				id=document.getElementById("nam"+i).parentNode.innerHTML.split(",")
				p+=","+id[2].substring(0,id[2].indexOf(")"))
				amts=document.getElementById("amt"+i).innerHTML.split(",")
				if (pr!="") amts[12]=pr
				if (gr!="") amts[13]=gr
				document.getElementById("rec"+i).innerHTML=((r=amts[12]*1+amts[13]*1)!=0) ? r : "&nbsp;"
				if (gp!="") document.getElementById("grp"+i).innerHTML=(gp.trim()=="")? "&nbsp;" : gp
				if (pd!="" && pa!="" && pa!=0) {
					for (j=0;j<3;j++)
						if (amts[j*2]=="") {
							amts[j*2]=pd
							amts[j*2+1]=pa
							okp=true
							break
						}
					if (j!=3) cp+=","+id[1]
					else alert(
						document.getElementById("nam"+i).innerHTML+
						" already has 3 primary contributions.  The new contribution was NOT added for this candidate.")
				}
				if (gd!="" && ga!="" && ga!=0) {
					for (j=3;j<6;j++)
						if (amts[j*2]=="") {
							amts[j*2]=gd
							amts[j*2+1]=ga
							break
						}
					if (j!=6) cg+=","+id[1]
					else alert(
						document.getElementById("nam"+i).innerHTML+
						" already has 3 general contributions.  The new contribution was NOT added for this candidate.")
				}
				for (a=0,amt="",act=0;a<12;a++) {
					amt+=amts[a]+","
					if (a%2==1) act+=amts[a]*1
				}
				document.getElementById("amt"+i).innerHTML=amt+amts[12]+","+amts[13]+","+amts[14]
				document.getElementById("act"+i).innerHTML=(act!=0) ? act : "&nbsp;"
				document.getElementById("chk"+i).checked=false
			}
		csf.CandsToUpdate.value=p+";"+cp+";"+cg
		csf.submit()
		csf.reset()
		document.getElementById("Sel").title="Select All"
		mdd.style.display="none"
	}	
}
