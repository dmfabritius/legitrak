var detailActive=0
var bsf,bdf,tdd,mdd
var ho

function init() {
	selectTab(0)
	bsf=document.getElementById("BillSummaryForm")
	bdf=document.getElementById("BillDetailForm")
	tdd=document.getElementById("TrackingDetails")
	mdd=document.getElementById("MultipleDetails")

	bdf.digCom.value=" Load\nDigest"
	scrollTop=getCookie("scrollTop")
	if (scrollTop!="") document.body.scrollTop=scrollTop
	if ((b=getCookie("ClientBill"))!=""){
		setCookie("ClientBill","")
		var bt=document.getElementById("Bills")
		for(i=0;i<bt.rows.length;i++)
			if (bt.rows[i].cells[1].innerHTML==b) break
		if (i<bt.rows.length){
			//Bills.rows(i).scrollIntoView()
			selectDetail(i-2)
			detailActive=1
		}
	}
}
function scr(){
	checkHeight()
	s=tdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-ho
	s=mdd.style
	if (s.display!="none") s.top=document.body.clientHeight+document.body.scrollTop-155
}
function sortBy(f){
	o=""
	if (f=="Pos") o="C.[PositionNum],"
	if (f=="Pri") o="C.[Dead],C.[PriorityNum],"
	if (f=="Location") o="CASE WHEN LTRIM(RTRIM(D.House))='' THEN D.Location ELSE D.House+', '+D.Location END,"
	setCookie("OrderField",o+"C.[Bill Number]")
	window.location.href="client-tracking.asp"
}
function updateFilter(){
	setCookie("FilterPri",bsf.FilterPri.value)
	setCookie("FilterPos",bsf.FilterPos.value)
	setCookie("FilterLoc",bsf.FilterLoc.value)
	window.location.href="client-tracking.asp"
}
function hideDetail(e,override) {
	e=(!e) ? event.srcElement : e.target
	while (e.parentNode!=null && !/Details/.test(e.id)) e=e.parentNode
	if (/Details/.test(e.id) && !override) return
	if (detailActive==-1)
		detailActive=1;
	else if (detailActive) {
		tdd.style.display="none"
		mdd.style.display="none"
	}
}
function addDetail() {
	checkHeight()
	detailActive=-1
	setCookie("scrollTop",document.body.scrollTop)
	bdf.reset()
	tdd.style.top=document.body.clientHeight+document.body.scrollTop-ho
	tdd.style.display="block"
	bdf.Pri.selectedIndex=getCookie("DefPriority")*1-1
	bdf.Pos.selectedIndex=getCookie("DefPosition")*1-1
	bdf.Bill.focus()
}
function selectDetail(i) {
	checkHeight()
	detailActive=-1
	setCookie("scrollTop",document.body.scrollTop)
	bdf.Index.value=i
	bdf.Bill.value=document.getElementById("bill"+i).innerHTML
	bdf.Title.value=document.getElementById("title"+i).innerHTML.fromHTML()
	bdf.Pri.value=document.getElementById("pri"+i).innerHTML
	bdf.Pos.value=document.getElementById("pos"+i).innerHTML.replace(/\&nbsp;/g,"")
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
	if (i==-1 || bdf.Bill.value!=document.getElementById("bill"+i).innerHTML || bdf.Delete.checked) {
		bdf.target=""
		bdf.action="client-tracking.asp"
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
		document.getElementById("notes"+i).innerHTML=(n!=="") ? "("+n.toHTML()+")" : ""
		document.getElementById("com"+i).innerHTML=bdf.Comments.value.toHTML()
		bdf.target="post"
		bdf.action="client-tracking-post.asp"
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
	if (c!=null) for (i=0,b=document.getElementsByName("ckbx");i<b.length;i++) b[i].checked=c
}
function selectMult(){
	detailActive=-1
	for(i=0,b=document.getElementsByName("ckbx"),c=false;i<b.length;i++) c|=b[i].checked
	if (!c){
		alert("To update the attributes of multiple bills, first select them using the checkboxes.")
		return
	}
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
	} else {
		if (bsf.Delete.checked) {
			bsf.target=""
			bsf.action="client-tracking.asp"
			bsf.submit()
		} else {
			bsf.target="post"
			bsf.action="client-tracking-post.asp"
			bc=myStyles[".bkg08"].backgroundColor
			pri=bsf.Pri.options[bsf.Pri.value].innerHTML.trim()
			pos=bsf.Pos.options[bsf.Pos.value].innerHTML.trim()
			for (i=0,b=document.getElementsByName("ckbx");i<b.length;i++)
				if (b[i].checked) {
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
				}
			bsf.submit()
			bsf.reset()
			document.getElementById("Sel").title="Select All"
			mdd.style.display="none"
		}
	}
}
function checkHeight(){
	r=Math.max(Math.min((document.body.clientHeight-380)/18,12),0)
	document.getElementsByName("Comments")[0].rows=3+r
	tdd.style.height=200+18*r
	ho=215+18*r
}