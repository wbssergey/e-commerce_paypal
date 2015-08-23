
//Do Browser Detection------------------------------------------------------------------
iAppType=0; IsIE=false; IsN4=false; IsN6=false; IsMac=false

iAppVs=(navigator.appVersion.indexOf('MSIE')>0)?navigator.appVersion.indexOf('MSIE')+5:0
iAppVs=parseFloat(navigator.appVersion.substr(iAppVs,3))
sAppAgent=navigator.userAgent.toLowerCase()

if (document.all){iAppType=1; IsIE=true}					//ie
else if (document.layers){iAppType=2; IsN4=true}			//nn4
else if (document.getElementById){iAppType=3; IsN6=true}	//nn6,mozilla
IsMac=(sAppAgent.indexOf('mac')!=-1)						//mac
//alert(' iAppType: '+iAppType +'\n iAppVs: '+iAppVs +'\n sAppAgent: '+sAppAgent)
//--------------------------------------------------------------------------------------

function IsEnterKey(e) {//keycode: 13=Enter
	kcode=IsIE?event.keyCode:(IsN4||IsN6)?e.which:0; return (kcode==13)
 }

function rtrim(str) { while(str.substr(str.length-1)==' '){str=str.substring(0,str.length-1)};return str }
function ltrim(str) { while(str.substr(0,1)==' '){str=str.substring(1,str.length)};return str }
function trim(str)  { return ltrim(rtrim(str)) }

	
function delRecord(url,vrb,vlu)
	{
	wrn='permanently delete record ' +vlu +' ??'			
	if(confirm(wrn)) { goForm(url,'Delete',vrb,vlu) }
	}
	
function resetRecord(vrb,vlu)
	{
	fm=document.getElementById('fUpdate')
	fm.intention.value='Reset'
	fm(vrb).value=vlu
	fm.submit()
	}

function goForm(url,itn,vrb,vlu)
	{
	var fm=document.forms.fGO; fm.action=url; fm.intention.value=itn
	
	for(var i=0;i<fm.elements.length;i++) { e=fm.elements[i]
		if(e.name!='intention'){e.name=vrb; e.value=vlu; fm.submit(); return false}
		}
	return false
	}

function autoAnswer(ctr,ctx,vlu)
	{	// ctr=control	ctx=context vlu=0=getValFromEle
	fld=document.getElementById('fld'+ctx)
	if(vlu==0){vlu=document.getElementById('vlu'+ctx).value}
	chk=document.getElementById('chk'+ctx)
	switch(ctr) {
		case 'chk': if(chk.checked==true){fld.value=vlu} else{fld.value=''}; break
		case 'fld': chk.checked=(fld.value==vlu); break
		}
	}

function openWin(url,tgt,wdt,hgt)
	{
	prp  = 'resizable=yes,scrollbars=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,'
	prp += 'width='+wdt+',height='+hgt+',top=50,left=150'
	
	win=window.open(url,tgt,prp)
	if(url.indexOf('://')==-1){win.focus()}
	}		

function pleaseWait(msg,bgclr,txclr,frmID) {
	var htm=''
	htm += '<html><head><title>Please Wait</title><link rel=styleSheet href='+siteRoot+'i_style.css></head><body>'
	htm += '<br><br><br><br><table class=datatablepage width=630 align=left valign=top cellpadding=2 cellspacing=1 border=0><tr>'
	htm += '<td style=background-color:'+bgclr+';color:'+txclr+';text-align:center>'
	htm += '<img src="'+siteRoot+'images/plswait.gif" alt="Please Wait"><br><br>'
	htm +=  msg
	htm += '<br><br><b><i>Please Wait &nbsp;.... </i></b><br>'
	htm += '</td></tr></table></body></html>'
	
	if(fm=document.getElementById(frmID)) {
		if(!bgclr==''){document.body.style.backgroundColor=bgclr}
		if(!txclr==''){document.body.style.color=txclr}
		fm.submit()
		document.body.innerHTML=htm
		return false
	 }
 }

function toggleFormDisable(frmNm,bSts) {
	//frmNm=form name	bSts=disable true/false
	var i,s,fm; s=(typeof(bSts)=='boolean'?bSts:false)
	if(!(fm=eval('document.forms.'+frmNm))) {return false}
	for(i=0;i<fm.elements.length;i++) {
		fm.elements[i].disabled=s
	 }
 }

function setFocus(oId) {
	var obj; if(obj=document.getElementById(oId)) {obj.focus()}
 }
	
function isValidEmail(eml) {
	var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/
	return (filter.test(eml))
 }

function visibility(oId,bActn)
	{	//oID: objectID		bActn: 0=hide 1=show
	var c,o,h,s; c=IsN4?'visibility':'style.display'; h=IsN4?'hide':'none'; s=IsN4?'show':'block'
	if(o=document.getElementById(oId)){eval('o.'+c+'=(bActn==0)?h:s')}
	}

function checkAll(fn,cn,xn)
	{	// fn=FormObjectName	cn=ControlObjectName	xn=checkboxname(string)
		for (i=0;i<fn.elements.length;i++) {
			var e=fn.elements[i]
			if ( e.id==xn && e.type=='checkbox' && !e.disabled ) {
				e.checked = cn.checked }
		 }
	}
	
function unChk(fn,cn)
	{	// fn=FormObjectName	cn=ControlObjectName
	var e=fn.elements[cn];
	e.checked=false;
	}
	
function acctWin(lnk,url)
	{	//	lnk=the '<a' object
	//var accwin;
	//accwin = window.open("",'acct_win','resizable=yes');
	//accwin.close();
	var numb = (Math.round((Math.random()*99)+1))
	var accwin = window.open('','acct_win'+numb,'resizable=yes,scrollbars=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,width=670,height=650,top=50,left=150')
	accwin.focus(); lnk.href=url; lnk.target='acct_win'+numb; acctWinRevGo(url); accwin = null;return true
	}		
function acctWin1(lnk,url,width,height,top,left)
	{	//	lnk=the '<a' object
	//var accwin;
	//accwin = window.open("",'acct_win','resizable=yes');
	//accwin.close();
	var numb = (Math.round((Math.random()*99)+1))
	var accwin = window.open('','acct_win'+numb,'resizable=yes,scrollbars=yes,toolbar=no,location=no,directories=no,status=no,menubar=no,width='+width+',height='+height+',top='+top+',left='+left)
	accwin.focus(); lnk.href=url; lnk.target='acct_win'+numb; acctWinRevGo(url); accwin = null;return true
	}		

function acctWinRevGo(url) {
	iTsk=window.setInterval('acctWinRev(\''+escape(url)+'\')',100); return //1000ms=1sec
	}			 
function acctWinRev(url) {
	window.clearInterval(iTsk); url=unescape(url)
	for(i=0;i<document.links.length;i++)
		{ lnk=document.links[i]; if(lnk.href==url) {lnk.href=siteRoot+'z.asp'} }
	}


function doSort(sortMaxNum) {
	fm=document.forms.fSubmit; fm.action=''
	fe=fm.elements; fl=fe.length
	sOrd=''; bWrn=0; sWrn='Order must be number between 0 and '+sortMaxNum+' !'
	for (i=0;i<fl;i++) { e=fe[i]
		if (e.name=='sortOrder') {
			v=parseInt(e.value)
			if(v.toString()=='NaN'||v.toString()!=e.value.toString()){v=0;bWrn=1} else{if(v<0||v>sortMaxNum){v=0;bWrn=1}}
			if(bWrn==1){e.focus();e.style.backgroundColor='red';alert(sWrn);e.style.backgroundColor='';e.value='';return false} else{if(sOrd==''){sOrd=v.toString()} else{sOrd+=','+v} }
			}
	}
	fm.action1x.value='sort'; fm.submit(); return false
}

// ===================================================================
// Author: Matt Kruse <matt@mattkruse.com>
// WWW: http://www.mattkruse.com/
//
// NOTICE: You may use this code for any purpose, commercial or
// private, without any further permission from the author. You may
// remove this notice from your final code if you wish, however it is
// appreciated by the author if at least my web site address is kept.
//
// You may *NOT* re-distribute this code in any way except through its
// use. That means, you can include it in your product, or your web
// site, or any other form where the code is actually being used. You
// may not put the plain javascript up on your site for download or
// include it in your javascript libraries for download. 
// If you wish to share this code with others, please just point them
// to the URL instead.
// Please DO NOT link directly to my .js files from your site. Copy
// the files to your server and use them there. Thank you.
// ===================================================================

// HISTORY
// ------------------------------------------------------------------
// May 17, 2003: Fixed bug in parseDate() for dates <1970
// March 11, 2003: Added parseDate() function
// March 11, 2003: Added "NNN" formatting option. Doesn't match up
//                 perfectly with SimpleDateFormat formats, but 
//                 backwards-compatability was required.

// ------------------------------------------------------------------
// These functions use the same 'format' strings as the 
// java.text.SimpleDateFormat class, with minor exceptions.
// The format string consists of the following abbreviations:
// 
// Field        | Full Form          | Short Form
// -------------+--------------------+-----------------------
// Year         | yyyy (4 digits)    | yy (2 digits), y (2 or 4 digits)
// Month        | MMM (name or abbr.)| MM (2 digits), M (1 or 2 digits)
//              | NNN (abbr.)        |
// Day of Month | dd (2 digits)      | d (1 or 2 digits)
// Day of Week  | EE (name)          | E (abbr)
// Hour (1-12)  | hh (2 digits)      | h (1 or 2 digits)
// Hour (0-23)  | HH (2 digits)      | H (1 or 2 digits)
// Hour (0-11)  | KK (2 digits)      | K (1 or 2 digits)
// Hour (1-24)  | kk (2 digits)      | k (1 or 2 digits)
// Minute       | mm (2 digits)      | m (1 or 2 digits)
// Second       | ss (2 digits)      | s (1 or 2 digits)
// AM/PM        | a                  |
//
// NOTE THE DIFFERENCE BETWEEN MM and mm! Month=MM, not mm!
// Examples:
//  "MMM d, y" matches: January 01, 2000
//                      Dec 1, 1900
//                      Nov 20, 00
//  "M/d/yy"   matches: 01/20/00
//                      9/2/00
//  "MMM dd, yyyy hh:mm:ssa" matches: "January 01, 2000 12:30:45AM"
// ------------------------------------------------------------------

var MONTH_NAMES=new Array('January','February','March','April','May','June','July','August','September','October','November','December','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var DAY_NAMES=new Array('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sun','Mon','Tue','Wed','Thu','Fri','Sat');
function LZ(x) {return(x<0||x>9?"":"0")+x}

// ------------------------------------------------------------------
// isDate ( date_string, format_string )
// Returns true if date string matches format of format string and
// is a valid date. Else returns false.
// It is recommended that you trim whitespace around the value before
// passing it to this function, as whitespace is NOT ignored!
// ------------------------------------------------------------------
function isDate(val,format) {
	var date=getDateFromFormat(val,format);
	if (date==0) { return false; }
	return true;
	}

// -------------------------------------------------------------------
// compareDates(date1,date1format,date2,date2format)
//   Compare two date strings to see which is greater.
//   Returns:
//   1 if date1 is greater than date2
//   0 if date2 is greater than date1 of if they are the same
//  -1 if either of the dates is in an invalid format
// -------------------------------------------------------------------
function compareDates(date1,dateformat1,date2,dateformat2) {
	var d1=getDateFromFormat(date1,dateformat1);
	var d2=getDateFromFormat(date2,dateformat2);
	if (d1==0 || d2==0) {
		return -1;
		}
	else if (d1 > d2) {
		return 1;
		}
	return 0;
	}

// ------------------------------------------------------------------
// formatDate (date_object, format)
// Returns a date in the output format specified.
// The format string uses the same abbreviations as in getDateFromFormat()
// ------------------------------------------------------------------
function formatDate(date,format) {
	format=format+"";
	var result="";
	var i_format=0;
	var c="";
	var token="";
	var y=date.getYear()+"";
	var M=date.getMonth()+1;
	var d=date.getDate();
	var E=date.getDay();
	var H=date.getHours();
	var m=date.getMinutes();
	var s=date.getSeconds();
	var yyyy,yy,MMM,MM,dd,hh,h,mm,ss,ampm,HH,H,KK,K,kk,k;
	// Convert real date parts into formatted versions
	var value=new Object();
	if (y.length < 4) {y=""+(y-0+1900);}
	value["y"]=""+y;
	value["yyyy"]=y;
	value["yy"]=y.substring(2,4);
	value["M"]=M;
	value["MM"]=LZ(M);
	value["MMM"]=MONTH_NAMES[M-1];
	value["NNN"]=MONTH_NAMES[M+11];
	value["d"]=d;
	value["dd"]=LZ(d);
	value["E"]=DAY_NAMES[E+7];
	value["EE"]=DAY_NAMES[E];
	value["H"]=H;
	value["HH"]=LZ(H);
	if (H==0){value["h"]=12;}
	else if (H>12){value["h"]=H-12;}
	else {value["h"]=H;}
	value["hh"]=LZ(value["h"]);
	if (H>11){value["K"]=H-12;} else {value["K"]=H;}
	value["k"]=H+1;
	value["KK"]=LZ(value["K"]);
	value["kk"]=LZ(value["k"]);
	if (H > 11) { value["a"]="PM"; }
	else { value["a"]="AM"; }
	value["m"]=m;
	value["mm"]=LZ(m);
	value["s"]=s;
	value["ss"]=LZ(s);
	while (i_format < format.length) {
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		if (value[token] != null) { result=result + value[token]; }
		else { result=result + token; }
		}
	return result;
	}
	
// ------------------------------------------------------------------
// Utility functions for parsing in getDateFromFormat()
// ------------------------------------------------------------------
function _isInteger(val) {
	var digits="1234567890";
	for (var i=0; i < val.length; i++) {
		if (digits.indexOf(val.charAt(i))==-1) { return false; }
		}
	return true;
	}
function _getInt(str,i,minlength,maxlength) {
	for (var x=maxlength; x>=minlength; x--) {
		var token=str.substring(i,i+x);
		if (token.length < minlength) { return null; }
		if (_isInteger(token)) { return token; }
		}
	return null;
	}
	
// ------------------------------------------------------------------
// getDateFromFormat( date_string , format_string )
//
// This function takes a date string and a format string. It matches
// If the date string matches the format string, it returns the 
// getTime() of the date. If it does not match, it returns 0.
// ------------------------------------------------------------------
function getDateFromFormat(val,format) {
	val=val+"";
	format=format+"";
	var i_val=0;
	var i_format=0;
	var c="";
	var token="";
	var token2="";
	var x,y;
	var now=new Date();
	var year=now.getYear();
	var month=now.getMonth()+1;
	var date=1;
	var hh=now.getHours();
	var mm=now.getMinutes();
	var ss=now.getSeconds();
	var ampm="";
	
	while (i_format < format.length) {
		// Get next token from format string
		c=format.charAt(i_format);
		token="";
		while ((format.charAt(i_format)==c) && (i_format < format.length)) {
			token += format.charAt(i_format++);
			}
		// Extract contents of value based on format token
		if (token=="yyyy" || token=="yy" || token=="y") {
			if (token=="yyyy") { x=4;y=4; }
			if (token=="yy")   { x=2;y=2; }
			if (token=="y")    { x=2;y=4; }
			year=_getInt(val,i_val,x,y);
			if (year==null) { return 0; }
			i_val += year.length;
			if (year.length==2) {
				if (year > 70) { year=1900+(year-0); }
				else { year=2000+(year-0); }
				}
			}
		else if (token=="MMM"||token=="NNN"){
			month=0;
			for (var i=0; i<MONTH_NAMES.length; i++) {
				var month_name=MONTH_NAMES[i];
				if (val.substring(i_val,i_val+month_name.length).toLowerCase()==month_name.toLowerCase()) {
					if (token=="MMM"||(token=="NNN"&&i>11)) {
						month=i+1;
						if (month>12) { month -= 12; }
						i_val += month_name.length;
						break;
						}
					}
				}
			if ((month < 1)||(month>12)){return 0;}
			}
		else if (token=="EE"||token=="E"){
			for (var i=0; i<DAY_NAMES.length; i++) {
				var day_name=DAY_NAMES[i];
				if (val.substring(i_val,i_val+day_name.length).toLowerCase()==day_name.toLowerCase()) {
					i_val += day_name.length;
					break;
					}
				}
			}
		else if (token=="MM"||token=="M") {
			month=_getInt(val,i_val,token.length,2);
			if(month==null||(month<1)||(month>12)){return 0;}
			i_val+=month.length;}
		else if (token=="dd"||token=="d") {
			date=_getInt(val,i_val,token.length,2);
			if(date==null||(date<1)||(date>31)){return 0;}
			i_val+=date.length;}
		else if (token=="hh"||token=="h") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>12)){return 0;}
			i_val+=hh.length;}
		else if (token=="HH"||token=="H") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>23)){return 0;}
			i_val+=hh.length;}
		else if (token=="KK"||token=="K") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<0)||(hh>11)){return 0;}
			i_val+=hh.length;}
		else if (token=="kk"||token=="k") {
			hh=_getInt(val,i_val,token.length,2);
			if(hh==null||(hh<1)||(hh>24)){return 0;}
			i_val+=hh.length;hh--;}
		else if (token=="mm"||token=="m") {
			mm=_getInt(val,i_val,token.length,2);
			if(mm==null||(mm<0)||(mm>59)){return 0;}
			i_val+=mm.length;}
		else if (token=="ss"||token=="s") {
			ss=_getInt(val,i_val,token.length,2);
			if(ss==null||(ss<0)||(ss>59)){return 0;}
			i_val+=ss.length;}
		else if (token=="a") {
			if (val.substring(i_val,i_val+2).toLowerCase()=="am") {ampm="AM";}
			else if (val.substring(i_val,i_val+2).toLowerCase()=="pm") {ampm="PM";}
			else {return 0;}
			i_val+=2;}
		else {
			if (val.substring(i_val,i_val+token.length)!=token) {return 0;}
			else {i_val+=token.length;}
			}
		}
	// If there are any trailing characters left in the value, it doesn't match
	if (i_val != val.length) { return 0; }
	// Is date valid for month?
	if (month==2) {
		// Check for leap year
		if ( ( (year%4==0)&&(year%100 != 0) ) || (year%400==0) ) { // leap year
			if (date > 29){ return 0; }
			}
		else { if (date > 28) { return 0; } }
		}
	if ((month==4)||(month==6)||(month==9)||(month==11)) {
		if (date > 30) { return 0; }
		}
	// Correct hours value
	if (hh<12 && ampm=="PM") { hh=hh-0+12; }
	else if (hh>11 && ampm=="AM") { hh-=12; }
	var newdate=new Date(year,month-1,date,hh,mm,ss);
	return newdate.getTime();
	}

// ------------------------------------------------------------------
// parseDate( date_string [, prefer_euro_format] )
//
// This function takes a date string and tries to match it to a
// number of possible date formats to get the value. It will try to
// match against the following international formats, in this order:
// y-M-d   MMM d, y   MMM d,y   y-MMM-d   d-MMM-y  MMM d
// M/d/y   M-d-y      M.d.y     MMM-d     M/d      M-d
// d/M/y   d-M-y      d.M.y     d-MMM     d/M      d-M
// A second argument may be passed to instruct the method to search
// for formats like d/M/y (european format) before M/d/y (American).
// Returns a Date object or null if no patterns match.
// ------------------------------------------------------------------
function parseDate(val) {
	var preferEuro=(arguments.length==2)?arguments[1]:false;
	generalFormats=new Array('y-M-d','MMM d, y','MMM d,y','y-MMM-d','d-MMM-y','MMM d');
	monthFirst=new Array('M/d/y','M-d-y','M.d.y','MMM-d','M/d','M-d');
	dateFirst =new Array('d/M/y','d-M-y','d.M.y','d-MMM','d/M','d-M');
	var checkList=new Array('generalFormats',preferEuro?'dateFirst':'monthFirst',preferEuro?'monthFirst':'dateFirst');
	var d=null;
	for (var i=0; i<checkList.length; i++) {
		var l=window[checkList[i]];
		for (var j=0; j<l.length; j++) {
			d=getDateFromFormat(val,l[j]);
			if (d!=0) { return new Date(d); }
			}
		}
	return null;
	}


