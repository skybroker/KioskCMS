// var aDay = new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
  var aDay = new Array("星期一","星期二","星期三","星期四","星期五","星期六","星期天");
// var aMonth = new Array("January","February","March","April","May","June","July","August","September","October","November","December")
var aMonth = new Array("1","2","3","4","5","6","7","8","9","10","11","12")

 function DateFormat(xdate,x)
  {
  x = x.toLowerCase();
  return (((x == "d")  ? xdate.getDate() : ((x == "dd") ? ((xdate.getDate() <= 9) ? "0"+xdate.getDate() : xdate.getDate()) : ((x == "ddd") ? aDay[xdate.getDay()].substring(0,3)+". " : ((x == "dddd") ? aDay[xdate.getDay()]+" " : ((x == "m")  ? xdate.getMonth()+1 : ((x == "mm") ? (((xdate.getMonth()+1) <= 9) ? "0"+(xdate.getMonth()+1) : xdate.getMonth()+1) : ((x == "mmm") ? aMonth[xdate.getMonth()].substring(0,3) : ((x == "mmmm") ? aMonth[xdate.getMonth()] : ((x == "y" || x == "yy" || x == "yyy") ? xdate.getFullYear().toString().substring(2,4) : ((x == "yyyy") ? xdate.getFullYear().toString() : "")))))))))))
  }


 function Showdate(_date, _var1, _var2, _var3, _var4, _del)
  {
  // ----------------------------------------------------------------------------------------
  // Content: Date formatter where _month = month, _day = day of week spelled out, yyyy = four digit year. 
  // The script will only display a day spelled out if the parameter is ddd, or dddd. dd or d means do not display the day.
  //
  // Syntax: document.writeln(Showdate(date object, "ddd", "mmm", "dd", "yyyy", "-"));
  //  where date can be a field variable = rs("Datecreated"); constant= new Date(yyyy,m,dd); 
  //
  // Results in: Thursday, January-01-1970  NOTE that the numeric month in Javascript is from 0-11
  //
  // timerID=setInterval("if (_ie){++nd; if (nd > xStyle.length) nd=0; document.all.md.innerHTML=Showdate(null, "mmm", "ddd", "yy" , "-");}",3000)
  // timerOn=true
  // ----------------------------------------------------------------------------------------

  var today = (_date == null) ? new Date() : new Date(_date);
  _del =  ((_del == null) ? " " : _del);
  return ( DateFormat(today, _var1) + DateFormat(today, _var2) + ((_var2 != "")? _del : "") + DateFormat(today, _var3) + ((_var3 != "")? _del : "") + DateFormat(today, _var4))
  }
  //Showdate(new Date(), 'dddd', 'mmm', 'dd', 'yyyy', '-') Monday, April-13-2009  
  //Showdate(new Date(), 'dd', 'mm', 'dd', 'yyyy', '.') 1304.13.2009 
  //Showdate(new Date(2001,11,25), 'dddd', 'mmm', 'dd', 'yy', ' ') Tuesday, Dec 25 01 
  //
  // document.write(Showdate(new Date(), 'dddd', 'mmmm', 'dd', 'yyyy', ','))
<!--
function Flash(url,w,h,s){
	if (s==1){
	document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="'+w+'" height="'+h+'"> ');
	document.write('<param name="movie" value="' + url + '">');
	document.write('<param name="quality" value="high"> ');
	document.write('<param name="wmode" value="transparent"> ');
	document.write('<param name="menu" value="false"> ');
	document.write('<embed src="' + url + '" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="'+w+'" height="'+h+'" wmode="transparent"></embed> ');
	document.write('</object> ');
	}
	else{
	document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="'+w+'" height="'+h+'"> ');
	document.write('<param name="movie" value="' + url + '">');
	document.write('<param name="quality" value="high"> ');
	document.write('<param name="menu" value="false"> ');
	document.write('<embed src="' + url + '" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="'+w+'" height="'+h+'"></embed> ');
	document.write('</object> ');
	}
}
//-->
function checkformcn(theForm)
{
  if (theForm.keywords.value == "Search phrase...")
  {
    alert("Please input keywords!");
    theForm.keywords.focus();
    return (false);
  }
}

function checkformeb(theForm)
{
  if (theForm.cUser.value == "")
  {
    alert("请输入用户名!");
    theForm.cUser.focus();
    return (false);
  }
  if (theForm.cPass.value == "")
  {
    alert("请输入密码!");
    theForm.cPass.focus();
    return (false);
  }
}