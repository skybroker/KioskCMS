<!DOCTYPE html>
<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title>新闻杂志</title>
	<link rel="shortcut icon" href="./favicon.ico">
	<link rel="stylesheet" href="./css/themes/default/jquery.mobile-1.4.5.min.css">
	<link rel="stylesheet" href="./css/kiosk.css">
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
	<script src="./js/idle-timer.min.js"></script>
</head>
<body>
<div data-role="page" class="jqm-demos" ><!-- data-quicklinks="true" -->
	<div data-role="header" class="my-header-800 main_title">
	    <a href="/" class="ui-btn-left ui-btn ui-btn-inline ui-mini ui-corner-all ui-btn-icon-left ui-icon-back" id="btn_back">返回</a>
			<h3>新闻杂志</h3>
			<!--
	    <button class="ui-btn-right ui-btn ui-btn-b ui-btn-inline ui-mini ui-corner-all ui-btn-icon-right ui-icon-check">Save</button>
			-->
	</div>
	<div role="main" class="ui-content-800">
<ul data-role="listview" class="news-list" data-inset="true">
<%
ClassID	=	trim(request.QueryString("ID"))
if not IsNumeric(ClassID) and ClassID<>"" then
	response.Write("传递参数出错！")
	response.End()
end if

sqlstr = "select id,NewsTitle,PostTime,NewsSPic from NewsInfo where NewsLock=0 and ClassID="&ClassID&"  order by NewsOrder desc "
set rs = Server.CreateObject("adodb.recordset")
rs.open sqlstr,conn,1,1
if rs.eof then
	response.Write("<li>暂时没有相关信息!</li>")
else
	CurrentPage = Request("page")
	if Not IsNumeric(CurrentPage) Then
		CurrentPage = "1"
	end if
	CurrentPage=Cint(CurrentPage)
	rs.PageSize = 5
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
		NewsSPic	=	rs("NewsSPic")
%>
	<li>
		<a href="newsviewer.asp?id=<%=id%>" target="_self">
			<img src="<%=NewsSPic%>" />
			<h2><%=NewsTitle%></h2>
			<p><%=PostTime%></p>
		</a>
	</li>
<%
		i=i+1
		if i >= rs.PageSize Then exit do 
		rs.movenext
	loop
end if
%>

</ul>
<div class="nav_bar">
<%
linkPrev="#"
linkNext="#"
if CurrentPage>1 then
	linkPrev = "./newslist.asp?id="& ClassID &"&page="& (CurrentPage-1)
end if
if CurrentPage<rs.pagecount then
	linkNext = "./newslist.asp?id="& ClassID &"&page="& (CurrentPage+1)
end if
%>
	<a class="ui-btn ui-btn-inline ui-btn-b" href='./newslist.asp?id=<%=ClassID%>&page=1'>首页</a>
	<a class="ui-btn ui-btn-inline ui-btn-b" href='<%=linkPrev%>'>上一页</a>
	第&nbsp;<%=CurrentPage%>&nbsp;/&nbsp;<%=rs.pagecount%>&nbsp;页
	<a class="ui-btn ui-btn-inline ui-btn-b" href='<%=linkNext%>'>下一页</a>
	<a class="ui-btn ui-btn-inline ui-btn-b" href='./newslist.asp?id=<%=ClassID%>&page=<%=rs.pagecount%>'>尾页</a>

</div>
	</div><!-- main div -->

</div><!-- /page -->
<%
rs.close
set rs = nothing
%>
</body>
<script src="./js/idle-back-home.js"></script>
</html>
