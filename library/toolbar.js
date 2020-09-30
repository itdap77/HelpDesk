var scFlag = false;
var scrollcount = 0;
var Strict_Compat = false;
var ToolBar_Supported = false;
var Frame_Supported   = false;
var doImage = doImage;
var TType = TType;

if (navigator.userAgent.indexOf("MSIE")    != -1 && 
	navigator.userAgent.indexOf("Windows") != -1 && 
	navigator.appVersion.substring(0,1) > 3)
{
	ToolBar_Supported = true;
	if (document.compatMode == "CSS1Compat")
		Strict_Compat = true;
}

if(doImage == null)
{
	var a= new Array();
	a[0] = prepTrackingString(window.location.hostname,7);
	if (TType == null)
	{	
		a[1] = prepTrackingString('PV',8);
	}
	else
	{
		a[1] = prepTrackingString(TType,8);
	}
	a[2] = prepTrackingString(window.location.pathname,0);
	if( '' != window.document.referrer)
	{
		a[a.length] = prepTrackingString(window.document.referrer,5);
	}
	
	if (navigator.userAgent.indexOf("SunOS") == -1 && navigator.userAgent.indexOf("Linux") == -1)
	{
		buildIMG(a);
	}
}	

if (ToolBar_Supported)
{
	var newLineChar = String.fromCharCode(10);
	var char34 = String.fromCharCode(34);
	var LastMSMenu = "";
	var LastICPMenu = "";
	var CurICPMenu = "";
	var IsMSMenu = false;
	var IsMenuDropDown = true;
	var HTMLStr;
	var TBLStr;
	var x = 0;
	var y = 0;
	var x2 = 0;
	var y2 = 0;
	var x3 = 0;
	var MSMenuWidth;
	var ToolbarMinWidth;
	var ToolbarMenu;
	var ToolbarBGColor;
	var ToolbarLoaded = false;
	var aDefMSColor  = new Array(3);
	var aDefICPColor = new Array(3);
	var aCurMSColor  = new Array(3);
	var aCurICPColor = new Array(3);
	var MSFont;
	var ICPFont;
	var MSFTFont;
	var ICPFTFont;
	var MaxMenu = 30;
	var TotalMenu = 0;
	var arrMenuInfo = new Array(30);
	var bFstICPTBMenu = true;
	var bFstICPFTMenu = true;
	
	var redirecturl = "<script language='Javascript' src='/library/include/ctredir.js'></script>";
	document.write(redirecturl);
	
	// Output style sheet and toolbar ID
	document.write("<SPAN ID='StartMenu' STYLE='display:none;'></SPAN>");

	// Build toolbar template
	HTMLStr  = "<DIV ID='idToolbar'     STYLE='background-color:white;width:100%;'>";
	HTMLStr += "<DIV ID='idRow1'        STYLE='position:relative;height:20px;'>";
	HTMLStr += "<DIV ID='idICPBanner'   STYLE='position:absolute;top:0px;left:0px;height:60px;width:250px;overflow:hidden;vertical-align:top;'><!--BEG_ICP_BANNER--><!--END_ICP_BANNER--></DIV>";
	HTMLStr += "<DIV ID='idMSMenuCurve' STYLE='position:absolute;top:0px;left:250px;height:20px;width:18px;overflow:hidden;vertical-align:top;'><IMG SRC='/library/toolbar/images/curve.gif' BORDER=0></DIV>";
	HTMLStr += "<DIV ID='idMSMenuPane'  STYLE='position:absolute;top:0px;left:250px;height:20px;width:10px;background-color:black;float:right;' NOWRAP><!--MS_MENU_TITLES--></DIV>";
	HTMLStr += "</DIV>";
	HTMLStr += "<DIV ID='idRow2' STYLE='position:relative;left:250px;height:40px;'>";
	HTMLStr += "<DIV ID='idADSBanner'   STYLE='position:absolute;top:0px;left:0px;height:40px;width:200px;vertical-align:top;overflow:hidden;'><!--BEG_ADS_BANNER--><!--END_ADS_BANNER--></DIV>";
	HTMLStr += "<DIV ID='idMSCBanner'   STYLE='position:absolute;top:0px;left:180px;height:40px;width:112px;vertical-align:top;overflow:hidden;' ALIGN=RIGHT><!--BEG_MSC_BANNER--><!--END_MSC_BANNER--></DIV>";
	HTMLStr += "</DIV>";
	HTMLStr += "<DIV ID='idRow3' STYLE='position:relative;height:20px;width:100%'>";
	HTMLStr += "<DIV ID='idICPMenuPane' STYLE='position:absolute;top:0px;left:0px;height:20px;background-color:black;' NOWRAP><!--ICP_MENU_TITLES--></DIV>";
	HTMLStr += "</DIV>";
	HTMLStr += "</DIV>";
	HTMLStr += 	"<SCRIPT TYPE='text/javascript'>" + 
				"   var ToolbarMenu = StartMenu;" + 
				"</SCRIPT>" + 
				"<DIV WIDTH=100%>";		

	// Define event handlers
	window.onresize  = resizeToolbar;
	window.onscroll  = scrollbaroptions;

	// Intialize global variables
	ToolbarBGColor	= "#1478EB";					// toolbar background color
	
	if (Strict_Compat)
	{
		MSFont  = "bold x-small Arial";
		ICPFont = "bold x-small Verdana";
	}
	else
	{
		MSFont  = "xx-small Verdana";
		ICPFont = "bold xx-small Verdana";
	}
	
	aDefMSColor[0]	= aCurMSColor[0]  = "#000000";	// bgcolor;
	aDefMSColor[1]	= aCurMSColor[1]  = "white";	// text font color
	aDefMSColor[2]  = aCurMSColor[2]  = "#FFCC00";	// mouseover font color
	
	aDefICPColor[0]	= aCurICPColor[0] = "#1478EB";	// bgcolor;
	aDefICPColor[1] = aCurICPColor[1] = "white";	// text font color
	aDefICPColor[2] = aCurICPColor[2] = "#FFCC00";	// mouseover font color
}

function drawToolbar()
{
	HTMLStr += "</DIV>";
	document.write(HTMLStr);
	ToolbarLoaded = true;

	MSMenuWidth     = Math.max(idMSMenuPane.offsetWidth, (200+112));
	ToolbarMinWidth = (250+18) + MSMenuWidth;

	idToolbar.style.backgroundColor     = ToolbarBGColor;
	idMSMenuPane.style.backgroundColor  = aDefMSColor[0];
	idICPMenuPane.style.backgroundColor = aDefICPColor[0];
	resizeToolbar();

	for (i = 0; i < TotalMenu; i++) 
	{
		thisMenu = document.all(arrMenuInfo[i].IDStr);
		if (thisMenu != null)
		{
			if (arrMenuInfo[i].IDStr == LastMSMenu && arrMenuInfo[i].type == "R")
			{
				//Last MSMenu has to be absolute width
				arrMenuInfo[i].type = "A";
				arrMenuInfo[i].unit = 200;
			}
			if (arrMenuInfo[i].type == "A")
				thisMenu.style.width = arrMenuInfo[i].unit + 'px';
			else 
				thisMenu.style.width = Math.round(arrMenuInfo[i].width * arrMenuInfo[i].unit) + 'em';
		}
	}
}

function resizeToolbar()
{
	scFlag = false;
	scrollcount = 0;
	if (ToolBar_Supported == false) return;

	w = Math.max(ToolbarMinWidth, document.body.clientWidth) - ToolbarMinWidth;
	if (document.all("idMSMenuCurve"))
	{	
		idMSMenuCurve.style.left  = (250+w) + 'px';
		idMSMenuPane.style.left   = (250+w+18) + 'px';
		idMSMenuPane.style.width  = MSMenuWidth  + 'px';
		idADSBanner.style.left    = (w+18)  + 'px';
		idMSCBanner.style.left    = (w+18+200)  + 'px';
		idMSCBanner.style.width   = (MSMenuWidth - 200)  + 'px';
		idICPMenuPane.style.width = ToolbarMinWidth + w  + 'px';
	}
}

function setICPBanner(Gif,Url,AltStr)
{	
	setBanner(Gif,Url,AltStr,"<!--BEG_ICP_BANNER-->","<!--END_ICP_BANNER-->");
}

function setBanner(BanGif, BanUrl, BanAltStr, BanBegTag, BanEndTag)
{
	begPos = HTMLStr.indexOf(BanBegTag);
	endPos = HTMLStr.indexOf(BanEndTag) + BanEndTag.length;
	SubStr = HTMLStr.substring(begPos, endPos);
	SrcStr = "";
	if (BanUrl != "")
		SrcStr += "<A Target='_top' HREF='" + formatURL(BanUrl, BanGif) + "'>";
	SrcStr += "<IMG SRC='" + BanGif + "' ALT='" + BanAltStr + "' BORDER=0>";
	if (BanUrl != "")
		SrcStr += "</A>";
	SrcStr = BanBegTag + SrcStr + BanEndTag;
	HTMLStr = HTMLStr.replace(SubStr, SrcStr);	
}

function setSubMenuWidth(MenuIDStr, WidthType, WidthUnit)
{
	var fFound = false;
	if (TotalMenu == MaxMenu)
	{
		alert("Unable to process menu. Maximum of " + MaxMenu + " reached.");
		return;
	}
	
	for (i = 0; i < TotalMenu; i++)
		if (arrMenuInfo[i].IDStr == MenuIDStr)
		{
			fFound = true;
			break;
		}

	if (!fFound)
	{
		arrMenuInfo[i] = new menuInfo(MenuIDStr);
		TotalMenu += 1;
	}

	if (!fFound && WidthType.toUpperCase().indexOf("DEFAULT") != -1)
	{
		arrMenuInfo[i].type = "A";
		arrMenuInfo[i].unit = 160;
	}
	else
	{
		arrMenuInfo[i].type = (WidthType.toUpperCase().indexOf("ABSOLUTE") != -1)? "A" : "R";
		arrMenuInfo[i].unit = WidthUnit;
	}
}

// This function creates a menuInfo object instance.
function menuInfo(MenuIDStr)
{
	this.IDStr = MenuIDStr;
	this.type  = "";
	this.unit  = 0;
	this.width = 0;
	this.count = 0;
}

function updateSubMenuWidth(MenuIDStr)
{
	for (i = 0; i < TotalMenu; i++)
		if (arrMenuInfo[i].IDStr == MenuIDStr)
		{
			if (arrMenuInfo[i].width < MenuIDStr.length) 
				arrMenuInfo[i].width = MenuIDStr.length;
			arrMenuInfo[i].count = arrMenuInfo[i].count + 1;
			break;
		}
}

function addICPMenu(MenuIDStr, MenuDisplayStr, MenuHelpStr, MenuURLStr)
{ 	
	if (addICPMenu.arguments.length > 4)
		TargetStr = addICPMenu.arguments[4];
	else
		TargetStr = "_top";
	tempID = "ICP_" + MenuIDStr;
	addMenu(tempID, MenuDisplayStr, MenuHelpStr, MenuURLStr, TargetStr, true); 
	bFstICPTBMenu=false;		
}

function addMSMenu(MenuIDStr, MenuDisplayStr, MenuHelpStr, MenuURLStr)
{	
	TargetStr = "_top";
	tempID = "MS_" + MenuIDStr;
	addMenu(tempID, MenuDisplayStr, MenuHelpStr, MenuURLStr, TargetStr, false); 
	LastMSMenu = tempID;
}

function addMenu(MenuIDStr, MenuDisplayStr, MenuHelpStr, MenuURLStr, TargetStr, bICPMenu)
{
	cFont   = bICPMenu? ICPFont : MSFont;
	cColor0 = bICPMenu? aDefICPColor[0] : aDefMSColor[0];
	cColor1 = bICPMenu? aDefICPColor[1] : aDefMSColor[1];
	cColor2 = bICPMenu? aDefICPColor[2] : aDefMSColor[2];
	tagStr  = bICPMenu? "<!--ICP_MENU_TITLES-->" : "<!--MS_MENU_TITLES-->";
	
	MenuStr = newLineChar;
	if ((bICPMenu == false && LastMSMenu != "") || (bICPMenu == true && bFstICPTBMenu==false))
		MenuStr += "<SPAN STYLE='font:" + cFont + ";color:" + cColor1 + "'>|&nbsp;</SPAN>"; 
	MenuStr += "<A TARGET='" + TargetStr + "' TITLE='" + MenuHelpStr + "'" +
			"   ID='AM_" + MenuIDStr + "'" +
			"   STYLE='text-decoration:none;cursor:hand;font:" + cFont + ";background-color:" + cColor0 + ";color:" + cColor1 + ";'";
	if (MenuURLStr != "")
	{
		if (bICPMenu)
			MenuStr += " HREF='" + formatURL(MenuURLStr, ("ICP_" + MenuDisplayStr)) + "'";
		else
			MenuStr += " HREF='" + formatURL(MenuURLStr, ("MS_" + MenuDisplayStr)) + "'";
	}
	else
		MenuStr += " HREF='' onclick='window.event.returnValue=false;'";
	MenuStr += 	" onmouseout="  + char34 + "mouseMenu('out' ,'" + MenuIDStr + "'); hideMenu();" + char34 + 
				" onmouseover=" + char34 + "mouseMenu('over','" + MenuIDStr + "'); doMenu('"+ MenuIDStr + "');" + char34 + ">" +
				"&nbsp;" + MenuDisplayStr + "&nbsp;</a>";
	MenuStr += tagStr;
	HTMLStr = HTMLStr.replace(tagStr, MenuStr);	
	setSubMenuWidth(MenuIDStr,"default",0);
}

function addICPSubMenu(MenuIDStr, SubMenuStr, SubMenuURLStr)
{	
	if (addICPSubMenu.arguments.length > 3)
		TargetStr = addICPSubMenu.arguments[3];
	else
		TargetStr = "_top";
	tempID = "ICP_" + MenuIDStr;
	addSubMenu(tempID,SubMenuStr,SubMenuURLStr,TargetStr,true); 
}

function addMSSubMenu(MenuIDStr, SubMenuStr, SubMenuURLStr)
{	
	TargetStr = "_top";
	tempID = "MS_" + MenuIDStr;
	addSubMenu(tempID,SubMenuStr,SubMenuURLStr,TargetStr,false); 
}

function addSubMenu(MenuIDStr, SubMenuStr, SubMenuURLStr, TargetStr, bICPMenu)
{
	cFont   = bICPMenu? ICPFont : MSFont;
	cColor0 = bICPMenu? aDefICPColor[0] : aDefMSColor[0];
	cColor1 = bICPMenu? aDefICPColor[1] : aDefMSColor[1];
	cColor2 = bICPMenu? aDefICPColor[2] : aDefMSColor[2];
	
	var MenuPos = MenuIDStr.toUpperCase().indexOf("MENU");
	if (MenuPos == -1) { MenuPos = MenuIDStr.length; }
	InstrumentStr = MenuIDStr.substring(0 , MenuPos) + "|" + SubMenuStr;
	URLStr        = formatURL(SubMenuURLStr, InstrumentStr);

	var LookUpTag  = "<!--" + MenuIDStr + "-->";
	var sPos = HTMLStr.indexOf(LookUpTag);
	if (sPos <= 0)
	{
		HTMLStr += newLineChar + newLineChar + "<SPAN ID='" + MenuIDStr + "'";
		HTMLStr += 	" STYLE='display:none;position:absolute;width:160px;background-color:" + cColor0 + ";padding-top:0px;padding-left:0px;padding-bottom:20px;z-index:9px;'";
		HTMLStr += "onmouseout='hideMenu();'>";		
		if (Frame_Supported == false || bICPMenu == false)
		HTMLStr += "<HR  STYLE='position:absolute;left:0px;top:0px;color:" + cColor1 + "' SIZE=1>";
		HTMLStr += "<DIV STYLE='position:relative;left:0px;top:8px;'>";
	}

	TempStr = newLineChar +
				"<A ID='AS_" + MenuIDStr + "'" +
				"   STYLE='text-decoration:none;cursor:hand;font:" + cFont + ";color:" + cColor1 + "'" +
				"   HREF='" + URLStr + "' TARGET='" + TargetStr + "'" +
				" onmouseout="  + char34 + "mouseMenu('out' ,'" + MenuIDStr + "');" + char34 + 
				" onmouseover=" + char34 + "mouseMenu('over','" + MenuIDStr + "');" + char34 + ">" +
				"&nbsp;" + SubMenuStr + "</A><BR>" + LookUpTag;
	if (sPos <= 0)
		HTMLStr += TempStr + "</DIV></SPAN>";
	else
		HTMLStr = HTMLStr.replace(LookUpTag, TempStr);	

	updateSubMenuWidth(MenuIDStr);	
}

function addICPSubMenuLine(MenuIDStr)
{	
	tempID = "ICP_" + MenuIDStr;
	addSubMenuLine(tempID,true);
}

function addMSSubMenuLine(MenuIDStr)
{	
	tempID = "MS_" + MenuIDStr;
	addSubMenuLine(tempID,false);
}

function addSubMenuLine(MenuIDStr, bICPMenu)
{
	var LookUpTag = "<!--" + MenuIDStr + "-->";
	var sPos = HTMLStr.indexOf(LookUpTag);
	if (sPos > 0)
	{
		cColor  = bICPMenu? aDefICPColor[1] : aDefMSColor[1];
		TempStr = newLineChar + "<HR STYLE='color:" + cColor + "' SIZE=1>" + LookUpTag;
		HTMLStr = HTMLStr.replace(LookUpTag, TempStr);
	}
}

function mouseMenu(id, MenuIDStr) 
{
	IsMSMenu   = (MenuIDStr.toUpperCase().indexOf("MS_") != -1);
	IsMouseout = (id.toUpperCase().indexOf("OUT") != -1);

	if (IsMouseout)
	{
		color = IsMSMenu? aDefMSColor[1] : aDefICPColor[1];
		if (MenuIDStr == CurICPMenu && aCurICPColor[1] != "") 
			color = aCurICPColor[1];
	}
	else
	{
		color = IsMSMenu? aDefMSColor[2] : aDefICPColor[2];
		if (MenuIDStr == CurICPMenu && aCurICPColor[2] != "") 
			color = aCurICPColor[2];
	}
	window.event.srcElement.style.color = color;
}

function doMenu(MenuIDStr) 
{
	var thisMenu = document.all(MenuIDStr);
	if (ToolbarMenu == null || thisMenu == null || thisMenu == ToolbarMenu) 
	{
		window.event.cancelBubble = true;
		return false;
	}
	// Reset dropdown menu
	window.event.cancelBubble = true;
	ToolbarMenu.style.display = "none";
	ToolbarMenu = thisMenu;
	IsMSMenu = (MenuIDStr.toUpperCase().indexOf("MS_") != -1);

	// Set dropdown menu display position
	x  = window.event.srcElement.offsetLeft +
	 	 window.event.srcElement.offsetParent.offsetLeft;

	if (MenuIDStr == LastMSMenu)
		x += (window.event.srcElement.offsetWidth - thisMenu.style.posWidth);
	x2 = x + window.event.srcElement.offsetWidth;
	y  = (IsMSMenu)? 
		 (idRow1.offsetHeight) :
		 (idRow1.offsetHeight + idRow2.offsetHeight + idRow3.offsetHeight);
		 
	thisMenu.style.top  = y;
	thisMenu.style.left = x;
	thisMenu.style.clip = "rect(0 0 0 0)";
	thisMenu.style.display = "block";
	thisMenu.style.zIndex = 102;

	// delay 2 millsecond to allow the value of ToolbarMenu.offsetHeight be set
	window.setTimeout("showMenu()", 2);
	return true;
}

function showMenu() 
{
	if (ToolbarMenu != null) 
	{ 
		IsMenuDropDown = (Frame_Supported && IsMSMenu == false)? false : true;
		if (IsMenuDropDown == false)
		{
			y = (y - ToolbarMenu.offsetHeight - idRow3.offsetHeight);
			if (y < 0) y = 0;
			ToolbarMenu.style.top = y;
		}
		y2 = y + ToolbarMenu.offsetHeight;

		ToolbarMenu.style.clip = "rect(auto auto auto auto)";
		x2 = x + ToolbarMenu.offsetWidth;
	}
}

function hideMenu()
{
	if (ToolbarMenu != null && ToolbarMenu != StartMenu) 
	{
		// Don't hide the menu if the mouse move between the menu and submenus
		cY = event.clientY + document.body.scrollTop;
		cX = event.clientX; 
		if (document.body.offsetWidth > x && scFlag) {
			cX = x + 9;
		}
		if ( (cX >= (x+5) && cX<=x2) &&
			((IsMenuDropDown == true  && cY > (y-10) && cY <= y2)      ||
			(IsMenuDropDown == false && cY >= y     && cY <= (y2+10)) ))
		{
			window.event.cancelBubble = true;
			return; 
		}
		ToolbarMenu.style.display = "none";
		ToolbarMenu = StartMenu;
		window.event.cancelBubble = true;
	}
}

function formatURL(URLStr, InstrumentStr)
{
	return URLStr;
}

function scrollbaroptions()
{
	scrollcount ++;
	if (scrollcount < 3)
	{
		scFlag = true;
	}else{
		scrollcount = 0;
		scFlag = false;
	}	
}

if (ToolBar_Supported != null && ToolBar_Supported == true)
{
	setICPBanner("/homepage/gif/bnr-microsoft.gif","/isapi/gomscom.asp?target=/","Microsoft Home") ;

	addMSMenu("ProductsMenu", "All Products", "","/isapi/gomscom.asp?target=/catalog/default.asp?subid=22");
	addMSSubMenu("ProductsMenu","Downloads","/isapi/gomscom.asp?target=/downloads/");
	addMSSubMenu("ProductsMenu","MS Product Catalog","/isapi/gomscom.asp?target=/catalog/default.asp?subid=22");
	addMSSubMenu("ProductsMenu","Microsoft Accessibility","/isapi/gomscom.asp?target=/enable/");
	addMSSubMenuLine("ProductsMenu");
	addMSSubMenu("ProductsMenu","Servers","/isapi/gomscom.asp?target=/servers/");
	addMSSubMenu("ProductsMenu","Developer Tools","/isapi/gomsdn.asp?target=/vstudio/");
	addMSSubMenu("ProductsMenu","Office","/isapi/gomscom.asp?target=/office/");
	addMSSubMenu("ProductsMenu","Windows","/isapi/gomscom.asp?target=/windows/");
	addMSSubMenu("ProductsMenu","MSN","http://www.msn.com/");

	addMSMenu("SupportMenu", "Support", "","http://www.microsoft.com/support");
	addMSSubMenu("SupportMenu","Knowledge Base","http://support.microsoft.com/search/");
	addMSSubMenu("SupportMenu","Developer Support","http://msdn.microsoft.com/support/");
	addMSSubMenu("SupportMenu","IT Pro Support"," http://www.microsoft.com/technet/support/");
	addMSSubMenu("SupportMenu","Product Support Options","http://www.microsoft.com/support");
	addMSSubMenu("SupportMenu","Service Providers","http://directory.microsoft.com/resourcedirectory/services.aspx");

	addMSMenu("SearchMenu", "Search", "","/isapi/gosearch.asp?target=/us/default.asp");					
	addMSSubMenu("SearchMenu","Search Microsoft.com","/isapi/gosearch.asp?target=/us/default.asp");
	addMSSubMenu("SearchMenu","MSN Web Search","http://search.msn.com/");

	addMSMenu("MicrosoftMenu", "Microsoft.com Guide", "","/isapi/gomscom.asp?target=/");
	addMSSubMenu("MicrosoftMenu","Microsoft.com Home","/isapi/gomscom.asp?target=/");
	addMSSubMenu("MicrosoftMenu","MSN Home","http://www.msn.com/");
	addMSSubMenuLine("MicrosoftMenu");
	addMSSubMenu("MicrosoftMenu","Contact Us","/isapi/goregwiz.asp?target=/regwiz/forms/contactus.asp");
	addMSSubMenu("MicrosoftMenu","Events","/isapi/gomscom.asp?target=/usa/events/default.asp");
	addMSSubMenu("MicrosoftMenu","Newsletters","/isapi/goregwiz.asp?target=/regsys/pic.asp?sec=0");
	addMSSubMenu("MicrosoftMenu","Profile Center","/isapi/goregwiz.asp?target=/regsys/pic.asp");
	addMSSubMenu("MicrosoftMenu","Training & Certification","http://www.microsoft.com/isapi/gomscom.asp?target=/traincert");
	addMSSubMenu("MicrosoftMenu","Free E-mail Account","http://www.hotmail.com/");

	addICPMenu("HomeMenu", "Microsoft Home", "Microsoft Home","/isapi/gomscom.asp?target=/");
	addICPMenu("MSNMenu", "MSN Home", "MSN Home","http://www.msn.com/");
	addICPMenu("SubscribeMenu", "Subscribe", "","http://www.microsoft.com/isapi/goregwiz.asp?target=/regsys/pic.asp");
	addICPSubMenu("SubscribeMenu","Newsletters","http://www.microsoft.com/isapi/goregwiz.asp?target=/regsys/pic.asp");
	addICPSubMenu("SubscribeMenu","Software","/isapi/gomscom.asp?target=/licensing/");
	addICPMenu("ProfileMenu","Manage Your Profile", "Manage Your Profile","http://www.microsoft.com/isapi/goregwiz.asp?target=/regsys/pic.asp");
}

function prepTrackingString(ts, type)
{
	var rArray;
	var rString;
	var pName = '';
	if (0 == type)
	{
		pName = 'p=';
		rString = ts.substring(1);
		rArray = rString.split('/');
	}
	if (1 == type)
	{
		pName = 'qs=';
		rString = ts.substring(1);
		rArray = rString.split('&');		
	}
	if (2 == type)
	{
		pName = 'f=';
		rString = escape(ts);
		return pName + rString;
	}
	if (3 == type)
	{
		pName = 'tPage=';
		rString = escape(ts);
		return pName+rString;
	}
	if (4 == type)
	{
		pName = 'sPage=';
		rString = escape(ts);
		return pName + rString;
	}
	if (5 == type)
	{
		pName = 'r=';
		rString = escape(ts);
		return pName + rString;
	}
	if (6 == type)
	{
		pName = 'MSID=';
		rString = escape(ts);
		return pName + rString;
	}
	if (7 == type)
	{
		pName = 'source=';
		rString = ts.toLowerCase();
		if(rString.indexOf("microsoft.com") != -1)
		{
			rString = rString.substring(0,rString.indexOf("microsoft.com"));
			if('' == rString)
				rString = "www";
			else
				rString = rString.substring(0,rString.length -1);
		}
		return pName + rString;
	}
	if (8 == type)
	{
		pName = 'TYPE=';
		rString = escape(ts);
		return pName + rString;
	}
	rString = '';
	if(null != rArray)
	{
		if(0 == type)
			for( j=0; j < rArray.length - 1; j++)
				rString += rArray[j] + '_';  
		else
			for( j=0; j < rArray.length  ; j++)
				rString += rArray[j] + '_';  
	}
	rString = rString.substring(0, rString.length - 1);  	 	
	return pName + rString;
}

function buildIMG(pArr)
{
	var TG = '<LAYER visibility="hide"><div style="display:none;"><IMG src="' + location.protocol + '//c.microsoft.com/trans_pixel.asp?';
	for(var i=0; i<pArr.length; i++)
	{
		if(0 == i)
			TG +=  pArr[i];
		else
			TG += '&' + pArr[i];
	}
	TG +='" height="0" width="0" hspace="0" vspace="0" Border="0"></div></layer>';
	document.writeln(TG);
}