function GetAbsoluteTop(elem)
{
	var topPosition = 0;

	while (elem)
	{
		if (elem.tagName == 'BODY')
		{
			break;
		}
		topPosition += elem.offsetTop;
		elem = elem.offsetParent;
	}
	return topPosition;
}

function GetAbsoluteLeft(elem)
{
	var leftPosition = 0;

	while (elem)
	{
		if (elem.tagName == 'BODY')
		{
			break;
		}
		leftPosition += elem.offsetLeft;
		elem = elem.offsetParent;
	}

	return leftPosition;
}

function CheckParent(oElement, oParent)
{
	var vParentFound = false;
	
	while (oElement)
	{
		if (oElement.tagName == 'BODY')
		{
			break;
		}
		else
		{
			
			if (oElement.id == oParent.id)
			{
				vParentFound = true;
				break;
			}
		}	
		oElement= oElement.parentElement;
	}
	return vParentFound;

}

function OpenPreview(oWindow, vURL) {

	oWindow.open(vURL,null,'scrollbars=yes,height=540,width=660,status=no,toolbar=no,menubar=no,location=no');

}

function OpenHelp() 
{
	var oDocument = window.document
	if (oDocument.mtHelpSection)
		var vURL = './help/?Area=' + oDocument.mtHelpSection;
	else
		var vURL = './help';
		
	window.open(vURL, null, 'scrollbars=yes,height=440,width=740,status=no,toolbar=no,menubar=no,location=no');
}

function OpenReport(IDInforme, ShowFilters, Filters) 
{

	if (ShowFilters)
		window.open('rptPreview.asp?IDInforme=' + IDInforme + '&ShowFilters=&' + Filters,null,'scrollbars=yes,height=550,width=750,status=no,toolbar=no,menubar=no,location=no');
	else
		window.open('rptPreview.asp?IDInforme=' + IDInforme + '&' + Filters,null,'scrollbars=yes,height=550,width=750,status=no,toolbar=no,menubar=no,location=no');

}

function Print(oWindow) {

	oWindow.print();

}


function CloseWindow(oWindow) {

	oWindow.close;
	
}

function MenuNada(){
	return false;
}

function Refresh(oWindow) {

	oWindow.location.reload();

}

function GetDateAR(vDate)
{
	var oDate = new Date(vDate);
	return oDate.getDate() + '/' + (oDate.getMonth() + 1) + '/' + oDate.getYear();
}


function MM_findObj(n, d) { 
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_validateForm() { 
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  var lang = getQuerystring('lang');
  var validemail, mustnumber, numberfrom, isrequired, errorocurred;
  if (lang == 'esp') { validemail = ' tiene que ser un correo electronico valido.\n'; mustnumber = ' tiene que ser un numero.\n'; numberfrom = ' tiene que ser un numero desde '; isrequired = ' es requerido.\n'; errorocurred = 'El siguiente error ha ocurrido:\n'; }
  else { validemail = ' must be a valid email address.\n'; mustnumber = ' must be number.\n'; numberfrom = ' it have to be a number from '; isrequired = ' is required.\n'; errorocurred = 'The following error ocurred:\n'; }
  
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
          if (p < 1 || p == (val.length - 1)) errors += '- ' + nm + validemail;
      } else if (test!='R') {
      if (isNaN(val)) errors += '- ' + nm + mustnumber;
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val < min || max < val) errors += '- ' + nm + numberfrom + min + ' to ' + max + '.\n';
      } 
  } 
} else if (test.charAt(0) == 'R') errors += '- ' + nm + isrequired;
}
}if (errors) { alert(errorocurred + errors); return 1; } else {return 0;}
  document.MM_returnValue = (errors == '');
}


function MM_nbGroup(event, grpName) { //v3.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : args[i+1];
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    if ((nbArr = document[grpName]) != null)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = args[i+1];
      nbArr[nbArr.length] = img;
  } }
}

function MM_preloadimages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadimages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}



// 2 funcion

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);




// 3 funcion (inicio.htm)
function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}



function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}



var blank = "blank.gif";

topedge = 0;  // location of news box from top of page
leftedge = 0;  // location of news box from left edge
boxheight = 600;  // height of news box
boxwidth = 180;  // width of news box
scrollheight = 1600; // total height of all data to be scrolled

function scrollnews(cliptop) {
if (document.layers) {
newsDiv = document.news;
newsDiv.clip.top = cliptop;
newsDiv.clip.bottom = cliptop + boxheight;
newsDiv.clip.left = 0;
newsDiv.clip.right = boxwidth + leftedge;
newsDiv.left = leftedge;
newsDiv.top = topedge - cliptop;
}
else {
newsDiv = news.style;
newsDiv.clip = "rect(" + cliptop + "px " + (boxwidth + leftedge) + "px " + (cliptop + boxheight) + "px 0px)";
newsDiv.pixelLeft = leftedge;
newsDiv.pixelTop = topedge - cliptop;
}
cliptop = (cliptop + 1) % (scrollheight + boxheight);
newsDiv.visibility='visible';
setTimeout("scrollnews(" + cliptop + ")", 15);
}


function openwin(url,target,data) 
{
window.open(url,target,data);
}

function reloadpage(frm)
{
frm.submit();
}

//<!-- Chequea las fechas -->
function checkdate(objName) {
var datefield = objName;
if (chkdate(objName) == false) {
datefield.select();
alert("La Fecha es inválida.  Ingresela Nuevamente.");
datefield.focus();
return false;
}
else {
return true;
   }
}

function chkdate(objName) {
//var strDatestyle = "US"; //United States date style
var strDatestyle = "EU";  //European date style
var strDate;
var strDateArray;
var strDay;
var strMonth;
var strYear;
var intday;
var intMonth;
var intYear;
var booFound = false;
var datefield = objName;
var strSeparatorArray = new Array("-"," ","/",".");
var intElementNr;
var err = 0;
var strMonthArray = new Array(12);
strMonthArray[0] = "Jan";
strMonthArray[1] = "Feb";
strMonthArray[2] = "Mar";
strMonthArray[3] = "Apr";
strMonthArray[4] = "May";
strMonthArray[5] = "Jun";
strMonthArray[6] = "Jul";
strMonthArray[7] = "Aug";
strMonthArray[8] = "Sep";
strMonthArray[9] = "Oct";
strMonthArray[10] = "Nov";
strMonthArray[11] = "Dec";
strDate = datefield.value;
if (strDate.length < 1) {
return true;
}
for (intElementNr = 0; intElementNr < strSeparatorArray.length; intElementNr++) {
if (strDate.indexOf(strSeparatorArray[intElementNr]) != -1) {
strDateArray = strDate.split(strSeparatorArray[intElementNr]);
if (strDateArray.length != 3) {
err = 1;
return false;
}
else {
strDay = strDateArray[0];
strMonth = strDateArray[1];
strYear = strDateArray[2];
}
booFound = true;
   }
}
if (booFound == false) {
if (strDate.length>5) {
strDay = strDate.substr(0, 2);
strMonth = strDate.substr(2, 2);
strYear = strDate.substr(4);
   }
}
if (strYear.length == 2) {
strYear = '20' + strYear;
}
// US style
if (strDatestyle == "US") {
strTemp = strDay;
strDay = strMonth;
strMonth = strTemp;
}
intday = parseInt(strDay, 10);
if (isNaN(intday)) {
err = 2;
return false;
}
intMonth = parseInt(strMonth, 10);
if (isNaN(intMonth)) {
for (i = 0;i<12;i++) {
if (strMonth.toUpperCase() == strMonthArray[i].toUpperCase()) {
intMonth = i+1;
strMonth = strMonthArray[i];
i = 12;
   }
}
if (isNaN(intMonth)) {
err = 3;
return false;
   }
}
intYear = parseInt(strYear, 10);
if (isNaN(intYear)) {
err = 4;
return false;
}
if (intMonth>12 || intMonth<1) {
err = 5;
return false;
}
if ((intMonth == 1 || intMonth == 3 || intMonth == 5 || intMonth == 7 || intMonth == 8 || intMonth == 10 || intMonth == 12) && (intday > 31 || intday < 1)) {
err = 6;
return false;
}
if ((intMonth == 4 || intMonth == 6 || intMonth == 9 || intMonth == 11) && (intday > 30 || intday < 1)) {
err = 7;
return false;
}
if (intMonth == 2) {
if (intday < 1) {
err = 8;
return false;
}
if (LeapYear(intYear) == true) {
if (intday > 29) {
err = 9;
return false;
}
}
else {
if (intday > 28) {
err = 10;
return false;
}
}
}
if (strDatestyle == "US") {
datefield.value = strMonthArray[intMonth-1] + " " + intday+" " + strYear;
}
else {
datefield.value = intday + " " + strMonthArray[intMonth-1] + " " + strYear;
}
return true;
}

function LeapYear(intYear) {
if (intYear % 100 == 0) {
if (intYear % 400 == 0) { return true; }
}
else {
if ((intYear % 4) == 0) { return true; }
}
return false;
}

function doDateCheck(from, to) {
if (Date.parse(from.value) <= Date.parse(to.value)) {
alert("Las Fechas son válidas.");
}
else {
if (from.value == "" || to.value == "") 
    alert("Ambas Fechas Tienen que ser ingresadas.");
    else 
    alert("La fecha tiene que ser posterior.");
   }
}
// fin de chequeo de fechas //


function CountLeft(field, count, max) 
{ 
  		if (field.value.length > max) field.value = field.value.substring(0, max); 
  		else  
  			 count.value = max - field.value.length; 
}

function getQuerystring(key, default_) {
    if (default_ == null) default_ = "";
    key = key.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
    var regex = new RegExp("[\\?&]" + key + "=([^&#]*)");
    var qs = regex.exec(window.location.href);
    if (qs == null)
        return default_;
    else
        return qs[1];
}
