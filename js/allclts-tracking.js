var detailActive=0
var bsf,bdf,tdd,mdd
var aClts=new Array()
var ho

function init() {
	selectTab(1)
	bsf=document.getElementById("BillSummaryForm")
	bdf=document.getElementById("BillDetailForm")
	tdd=document.getElementById("TrackingDetails")
	mdd=document.getElementById("MultipleDetails")
	loadClts()

	bdf.digCom.value=" Load\nDigest"
	parent.document.getElementById("fs").rows="20,*,0"
	scrollTop=getCookie("scrollTop")
	if (scrollTop!="") document.body.scrollTop=scrollTop
	scr()
}
function loadClts(){
	c=top.contents.document.getElementById("ClientMenu")
	for(var i=0;i<c.rows.length-1;i+=2){
		o=c.rows[i].cells[0].childNodes[0].childNodes[0]
		aClts[o.innerHTML]=new Array(o.id,o.title)
	}
	for (i=0;i<bsf.BillCount.value;i++){
		o=document.getElementById("name"+i)
		o.title=aClts[o.innerHTML][1]
	}
}
function quickAdd(){
	parent.qa.location.href="allclts-quickadd.asp"
	parent.document.getElementById("fs").rows="20,0,*"
}
function scr(){
	checkHeight()
	s=tdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-ho
	s=mdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-155
}
function sortBy(f) {
	o=""
	d="C.[Bill Number],CL.[Short Company Name]"
	if (f=="Position") o="C.[PositionNum],"
	if (f=="Priority") o="C.[Dead],C.[PriorityNum],"
	if (f=="Location") o="CASE WHEN LTRIM(RTRIM(D.House))='' THEN D.Location ELSE D.House+', '+D.Location END,"
	if (f=="Client") d="CL.[Short Company Name],C.[Bill Number]"
	setCookie("OrderField2",o+d)
	window.location.href="allclts-tracking.asp"
}
function updateFilter(){
	setCookie("FilterPri",bsf.FilterPri.value)
	setCookie("FilterPos",bsf.FilterPos.value)
	setCookie("FilterLoc",bsf.FilterLoc.value)
	window.location.href="allclts-tracking.asp"
}
function hideDetail(e,override) {
	e = (!e) ? event.srcElement : e.target
	while (e.parentNode!=null && !/Details/.test(e.id)) e=e.parentNode
	if (/Details/.test(e.id) && !override) return
	if (detailActive==-1)
		detailActive=1;
	else if (detailActive) {
		tdd.style.display="none"
		mdd.style.display="none"
	}
}
function selectDetail(i) {
	checkHeight()
	detailActive=-1
	setCookie("scrollTop",document.body.scrollTop)
	bdf.Index.value=i
	n=document.getElementById("name"+i).innerHTML
	bdf.UpdateClientID.value=aClts[n][0]
	bdf.ClientName.value=aClts[n][1]
	bdf.Bill.value=document.getElementById("bill"+i).innerHTML
	bdf.Title.value=document.getElementById("title"+i).innerHTML.fromHTML()
	bdf.Pri.value=document.getElementById("pri"+i).innerHTML
	bdf.Pos.value=document.getElementById("pos"+i).innerHTML
	bdf.Dead.checked=(document.getElementById("dead"+i).innerHTML=="True")
	n=document.getElementById("notes"+i).innerHTML.fromHTML()
	bdf.Notes.value=n.substr(1,n.length-2)
	bdf.Comments.value=document.getElementById("com"+i).innerHTML.fromHTML().trim()
	tdd.style.top=document.body.clientHeight+document.body.scrollTop-ho
	tdd.style.display="block"
	bdf.Comments.focus()
}
function submitDetail(){
	i=bdf.Index.value
	bdf.Notes.value=(n=bdf.Notes.value.trim())
	bdf.Comments.value=bdf.Comments.value.trim()
	if (bdf.Bill.value!=document.getElementById("bill"+i).innerHTML || bdf.Delete.checked) {
		bdf.target=""
		bdf.action="allclts-tracking.asp"
		setCookie("scrollTop",document.body.scrollTop)
	} else {
		p=document.getElementById("pri"+i)
		d=document.getElementById("dead"+i)
		p.innerHTML=bdf.Pri.value
		if (bdf.Dead.checked) {
			d.innerHTML="True"
			p.style.backgroundColor=myStyles[".bkg08"].backgroundColor
			p.style.padding="0 5"
		} else {
			d.innerHTML="False"
			p.style.backgroundColor="transparent"
		}
		document.getElementById("pos"+i).innerHTML=bdf.Pos.value
		document.getElementById("notes"+i).innerHTML=(n!=="")? "("+n.toHTML()+")" : ""
		document.getElementById("com"+i).innerHTML=bdf.Comments.value.toHTML()
		bdf.target="post"
		bdf.action="allclts-tracking-post.asp"
		tdd.style.display="none"
	}
}
function selectBills(e,c){
	if (e.title=="Select All"){
		if (confirm("Would you like to select all of the bills?")){
			e.title="Un-Select All"
			c=true
		}
	} else {
		if (confirm("Would you like to un-select all of the bills?")){
			e.title="Select All"
			c=false
		}
	}
	if (c!=null) for (i=0;i<bsf.BillCount.value;i++) document.getElementById("chk"+i).checked=c
}
function selectMult(){
	if ((cnt=bsf.BillCount.value)>150) alert("This may take a few seconds, click OK to proceed...")
	detailActive=-1
	a=new Array()
	for (i=0;i<cnt;i++)
		if (document.getElementById("chk"+i).checked)
			a.push(aClts[document.getElementById("name"+i).innerHTML][0]+","+document.getElementById("bill"+i).innerHTML)
	if (a.length==0){
		alert("To update the attributes of multiple bills, first select them using the check boxes.")
		return
	}
	a.sort()
	for(i=0,b="",t="";i<a.length;i++) {
		s=a[i].split(",")
		if (s[0]!=t) b+=";"+(t=s[0])
		b+=","+s[1]
	}
	bsf.BillsToUpdate.value=b
	mdd.style.top=document.body.clientHeight+document.body.scrollTop-155
	mdd.style.display="block"
}
function submitMult(){
	if (!bsf.Delete.checked &&
		(bsf.Notes.value=bsf.Notes.value.trim())=="" &&
		bsf.Pri.selectedIndex==0 &&
		bsf.Pos.selectedIndex==0 &&
		bsf.Dead.selectedIndex==0){
			alert("No changes requested.")
	} else if (bsf.Delete.checked) {
		bsf.target=""
		bsf.action="allclts-tracking.asp"
		bsf.submit()
	} else {
		bsf.target="post"
		bsf.action="allclts-tracking-post.asp"
		bc=myStyles[".bkg08"].backgroundColor
		pri=bsf.Pri.options[bsf.Pri.value].innerHTML.trim()
		pos=bsf.Pos.options[bsf.Pos.value].innerHTML.trim()
		for(i=0;i<bsf.BillCount.value;i++)
			if(document.getElementById("chk"+i).checked) {
				p=document.getElementById("pri"+i)
				d=document.getElementById("dead"+i)
				if (bsf.Pri.value!=0) p.innerHTML=pri
				if (bsf.Dead.value==1) {
					d.innerHTML="True"
					p.style.backgroundColor=bc
					p.style.padding="0 5"
				} else if (bsf.Dead.value!=-1){
					d.innerHTML="False"
					p.style.backgroundColor="transparent"
				}
				if (bsf.Pos.value!=0) document.getElementById("pos"+i).innerHTML=pos
				if (bsf.Notes.value!="") document.getElementById("notes"+i).innerHTML="("+bsf.Notes.value.toHTML()+")"
				document.getElementById("chk"+i).checked=false
			}
		bsf.submit()
		bsf.reset()
		document.getElementById("Sel").title="Select All"
		mdd.style.display="none"
	}
}
function checkHeight(){
	r=Math.max(Math.min((document.body.clientHeight-380)/18,12),0)
	document.getElementsByName("Comments")[0].rows=3+r
	tdd.style.height=220+18*r
	ho=235+18*r
}