var MSIE = /MSIE/.test(window.navigator.userAgent)

function loadStyles(){
	if (document.styleSheets.length==0) return
	s=new Array()
	if (document.styleSheets[0].rules)
		css=document.styleSheets[0].rules
	else
		css=document.styleSheets[0].cssRules
	for(var i=0;i<css.length;i++) s[css[i].selectorText]=css[i].style
	return s
}
var myStyles=loadStyles()

String.prototype.trim=function(){
    return this.replace(/(^[\s,\n]*)|([\s,\n]*$)/g, "")
}
String.prototype.toHTML=function(){
    return this.replace(/\&/g,"&amp;").replace(/\n/g,"<BR>")
}
String.prototype.fromHTML=function(){
    return this.replace(/\&amp;/g,"&").replace(/<BR>|<br>/g,"\n")
}
function setCookie(name, value) {
//	document.cookie=name+"="+escape(value).replace(/\+/g,"%2B")
	c=/LegiTrak=([^;]+)/.exec(document.cookie)
	c=(c!=null)? c[1] : ""
	t1=c.split("&")
	for(i=0;i<t1.length;i++)
		if ((t2=t1[i].split("="))[0]==name) {
			t1[i]=t2[0]+"="+escape(value).replace(/\+/g,"%2B")
			break
		}
	if (i==t1.length)
		value=c+((c!="")? "&" : "" )+name+"="+escape(value).replace(/\+/g,"%2B")	// new sub-cookie
	else
		value=t1.toString().replace(/,/g,"&")	// replace existing sub-cookie
	document.cookie="LegiTrak="+value+";path=/"
}
function getCookie(name) {
	if ((c=/LegiTrak=([^;]+)/.exec(document.cookie))==null) return ""
	re=new RegExp(name+"=([^&]+)")
	value=re.exec(c[1])
	return (value!=null)? unescape(value[1].replace(/\+/g,"%20")) : ""
}
function isBill(e){
	b=parseInt(e.value)
	if(isNaN(b)||b<1000||b>9999){
		alert('Please enter a bill number between 1000 and 9999.')
		return false
	}
	mark(e)
	return true
}
function isEmail(e,o){
	if (e.value.trim().length==0) {
		if (!o) { // address is optional when true
			alert("Please enter an e-mail address.")
			return false
		}
		return true
	}
	ok=true
	a=/^([^<>()[\]\\.,;:\s@\'\"]+(\.[^<>()[\]\\.,;:\s@\'\"]+)*)@(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,})$/
	d=/\.(com|net|org|gov|edu)$/
	e1=e.value.split(",")
	e3=""
	for (i=0;i<e1.length;i++) {
		e2=e1[i].split(";")
		for (j=0;j<e2.length;j++) {
			t=e2[j].trim()
			if (t.length!=0)
				if (a.test(t)) {
					e3+="; "+t
					if (!d.test(t)) alert(
						"Warning: The domain name of the e-mail address you entered, "+
						t+",\nis not one of the most common domains (com, net, org, gov, and edu).\n"+
						"Please be sure you entered your address correctly.")
				} else {
					alert(
						"The e-mail address you entered, "+t+", is invalid.\n"+
						"Please re-type your address.")
					ok=false
				}
		}
	}
	if(ok){
		e.value=e3.substr(2)
		mark(e)
		return true
	}
	return false
}
function isDate(e,d){
	if (a=e.value.match(/^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{2,4})$/)) {
		if (a[5]<1900) a[5]=a[5]*1+2000
		d=new Date(a[5],a[1]-1,a[3])
		if (isNaN(d)||d.getFullYear()!=a[5]||d.getMonth()!=a[1]-1||d.getDate()!=a[3]) d=null
	} else if (a=e.value.match(/^(\d{1,2})(\/|-)(\d{1,2})$/)) {
		a[5]=new Date().getFullYear()
		d=new Date(a[5],a[1]-1,a[3])
		if (isNaN(d)||d.getFullYear()!=a[5]||d.getMonth()!=a[1]-1||d.getDate()!=a[3]) d=null
	}
	if (!d){
		alert("Please enter a validate date.")
		e.focus()
		return false
	}
	e.value=a[1]+"/"+a[3]+"/"+a[5]
	return true
}
function colHover(e,h) {
	s=e.style
	if (h) {
		s.colorSave=s.color
		s.color=myStyles[".hdg29"].backgroundColor
		s.bgSave=s.backgroundColor
		s.backgroundColor=myStyles[".hdg29"].color
	} else {
		s.color=s.colorSave
		s.backgroundColor=s.bgSave
	}
}
function menuHover(e,h) {
	while (e.tagName!="TD") e=e.parentNode
	if (e.style.color==myStyles[".mnu61"].color) return
	img=document.getElementById(e.id+"b")
	if (h) {
		e.style.color=myStyles[".mnu61"].backgroundColor
		e.style.backgroundColor=myStyles[".bkg0A"].backgroundColor
		img.src="img/m2.gif"
	} else {
		e.style.color=myStyles[".mnu29"].color
		e.style.backgroundColor=myStyles[".mnu29"].backgroundColor
		img.src="img/m0.gif"
	}
}
function menuSelect(e,url,i){
	if (i==null) {
		d=document
	} else {
		e=top.contents.document.getElementById("ClientMenu").rows[i].cells[0]
		e.scrollIntoView()
		d=top.contents.document
	}
	while (e.tagName!="TD") e=e.parentNode
	
	if ((m=getCookie("menuItem"))=="") m="mnu1"
	if (e.id==m) return
	menuUnSelect(m)

	e.className="mnu61"
	e.style.color=myStyles[".mnu61"].color
	e.style.backgroundColor=myStyles[".mnu61"].backgroundColor
	d.getElementById(e.id+"b").src="img/m1.gif"

	setCookie("menuItem",e.id)
	sh=top.subheading.document
	if (/clt/.test(e.id)) {
		o=e.childNodes[0].childNodes[0]
		setCookie("ClientID",o.id)
		setCookie("ClientName",o.title)
		sh.getElementById("sep").innerHTML="&#153;&#151;"
		sh.getElementById("cltName").innerHTML=o.title
		if (/client/.test(top.details.location.href))
			top.details.data.location.href=top.details.data.location.href
		else
//			top.details.location.href= (!url) ? "client.htm" : url
			top.details.location.href= (!url) ? ((getCookie("SessionStatus")!=3)? "client.htm" : "client-campaign.htm") : url
	} else {
		sh.getElementById("cltName").innerHTML=""
		sh.getElementById("sep").innerHTML=""
		top.details.location.href=url
	}
}
function menuUnSelect(m){
	if (!m || m=="") return
	if (/clt/.test(m))
		o=top.contents
	else
		o=top.menu
	o.document.getElementById(m+"b").src="img/m0.gif"
	o=o.document.getElementById(m)
	o.className="mnu29"
	o.style.color=myStyles[".mnu29"].color
	o.style.backgroundColor=myStyles[".mnu29"].backgroundColor
}
function selectMenu(m,url){
	menuUnSelect(getCookie("menuItem"))

	setCookie("menuItem",m)
	o=top.menu.document
	e=o.getElementById(m)

	e.className="mnu61"
	e.style.color=myStyles[".mnu61"].color
	e.style.backgroundColor=myStyles[".mnu61"].backgroundColor
	o.getElementById(e.id+"b").src="img/m1.gif"

	sh=top.subheading.document
	sh.getElementById("cltName").innerHTML=""
	sh.getElementById("sep").innerHTML=""
	top.details.location.href=url
}
function tabHover(e,h) {
	if (e.style.color==myStyles[".tab61"].color) return
	if (h) {
		e.style.color=myStyles[".tab61"].backgroundColor
		e.style.backgroundColor=myStyles[".bkg0A"].backgroundColor
		i=2
	} else {
		e.style.color=myStyles[".tab29"].color
		e.style.backgroundColor=myStyles[".tab29"].backgroundColor
		i=0
	}
	o=parent.tabs.document
	tab=e.id.substr(1)
	o.images[tab*2].src="img/tl"+i+".gif"
	o.images[tab*2+1].src="img/tr"+i+".gif"
}
function tabSelect(e,url){
	if (e.style.color==myStyles[".tab61"].color) return
	parent.data.location.href=url
}
function selectTab(tab){
	o=parent.tabs.document

	for (i=0;d=o.getElementById("d"+i);i++) d.style.zIndex=9-i
	for (i=0;t=o.getElementById("t"+i);i++) {
		t.className="tab29"
		t.style.color=myStyles[".tab29"].color
		t.style.backgroundColor=myStyles[".tab29"].backgroundColor
	}

	for (i=0;i<o.images.length;i+=2) {
		o.images[i].src="img/tl0.gif"
		o.images[i+1].src="img/tr0.gif"
	}

	o.getElementById("d"+tab).style.zIndex=10
	o.getElementById("t"+tab).className="tab61"
	o.getElementById("t"+tab).style.color=myStyles[".tab61"].color
	o.getElementById("t"+tab).style.backgroundColor=myStyles[".tab61"].backgroundColor
	o.images[tab*2].src="img/tl1.gif"
	o.images[tab*2+1].src="img/tr1.gif"
}
function mark(e) {
	e.style.backgroundColor=myStyles[".hdg29"].backgroundColor
}
function lnk(e,s){
	e = (!e) ? event.srcElement : e.target
	if (e.tagName!="U") return
	t=e.innerHTML
	b=parseInt(e.parentNode.parentNode.childNodes[1].innerHTML)
	if (isNaN(b)){
		b=e.parentNode.parentNode.childNodes[0].innerHTML
		b=b.substr(b.length-4)
	}
	if (t=="A")
		u="http://apps.leg.wa.gov/billinfo/summary.aspx?bill="+b

	else {
		y=new Date().getFullYear()
		y=(y%2) ? y+"-"+(y+1).toString().substr(2) : (y-1)+"-"+y.toString().substr(2)
		c=(b>4999) ? "senate" : "house"
		x=""
		if (r=/[2-9]*S/.exec(s)) x+="-s"+((r.toString().length==2) ? r.toString().substr(0,1) : "")
		if (e.innerHTML=="D") {
			x+=".dig"
			d="digests/"+c
		} else {
			d="bills/"+c+" bills"
			if (r=/[2-9]*E/.exec(s)) x+=".e"+((r.toString().length==2) ? r.toString().substr(0,1) : "")
		}
		u="http://lawfilesext.leg.wa.gov/biennium/"+y+"/pdf/"+d+"/"+b+x+".pdf"

	}
	window.open(u,'new') 
}
function sizePix(){
	if (!document.getElementById("d2").filters) return
	dw=document.body.offsetWidth*0.9
	dh=h=document.body.offsetHeight*0.9
	w=dh*1.53
	if(w>dw){h=h*(dw/w);w=dw}
	if(h>dh){w=w*(dh/h);h=dh}
	d1.style.width=w
	d1.style.height=h
}
function dig2Com(){
	if (!isBill(bdf.Bill)) {
		return false
	} else {
		setCookie("digest","~")
		document.getElementById('DigestFrame').src="digest.asp?bill="+bdf.Bill.value
		if (!MSIE) alert("Click OK to replace comments with the bill digest.")
		while ((d=getCookie("digest"))=="~") if (!MSIE) break;
		bdf.Comments.value=d
	}
}
