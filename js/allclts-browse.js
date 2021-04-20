var bf,bh,dh
function init(){
	selectTab(2)
	parent.document.getElementById("fs").rows="20,*,0"
	document.body.scrollTop=0
	bf=document.getElementById("BillsForm")
	bh=document.getElementById("BillHead")
	dh=document.getElementById("DigestHead")
	if (getCookie("Source")=='1') {
		selDigests()
		bf.DigestStart.focus()
	} else {
		selBills()
		bf.BillStart.select()
	}
}
function selBills(){
	bh.className="hdg29"
	bf.BillStart.disabled=false
	bf.BillEnd.disabled=false

	dh.className="hdg89"
	bf.DigestStart.disabled=true
	bf.DigestEnd.disabled=true
}
function selDigests(){
	dh.className="hdg29"
	bf.DigestStart.disabled=false
	bf.DigestEnd.disabled=false

	bh.className="hdg89"
	bf.BillStart.disabled=true
	bf.BillEnd.disabled=true
}
function allBills() {
	bf.BillStart.value=bf.minBill.value
	bf.BillStart.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	bf.BillEnd.value=bf.maxBill.value
	bf.BillEnd.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	selBills()
} 
function allDigests() {
	bf.DigestStart.value=bf.minSup.value
	bf.DigestStart.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	bf.DigestEnd.value=bf.maxSup.value
	bf.DigestEnd.style.backgroundColor=myStyles[".hdg29"].backgroundColor
	selDigests()
} 
function submitFilters(){
	setCookie("BillStart",bf.BillStart.value)
	setCookie("BillEnd",bf.BillEnd.value)
	setCookie("DigestStart",bf.DigestStart.value)
	setCookie("DigestEnd",bf.DigestEnd.value)
	setCookie("FilterSponsor",bf.FilterSponsor.selectedIndex)
	setCookie("FilterComm",bf.FilterComm.selectedIndex)
	setCookie("FilterLevel",bf.FilterLevel.selectedIndex)
	if (!bf.BillStart.disabled)
		setCookie("Source",0)
	else
		setCookie("Source",1)
}
function quickAdd(b){
	setCookie("QuickAddBill",b)
	parent.qa.location.href="allclts-quickadd.asp"
	parent.document.getElementById("fs").rows="20,0,*"
}