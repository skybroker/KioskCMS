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
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
	<style>
		.ui-content {
			max-width: 800px;
			margin: auto;
		}
		body {
			background-position: left top;
			background-image: url(images/home_bg.jpg);
			background-repeat: no-repeat;
			background-size: 100%;
			background-attachment: fixed;
		}
		.jqm-demos {
			background-color: transparent !important;
		}
		li a {
			background-color: transparent !important;
			vertical-align: middle;
			border-left: none !important;
			border-right: none !important;
		}
		li a img {
			height: 80px;
		}
	</style>
</head>
<body>
<div data-role="page" class="jqm-demos" data-quicklinks="true">

	<div role="main" class="ui-content jqm-content">
<ul data-role="listview" data-inset="true">
<%
ClassID	=	trim(request.QueryString("ID"))
if not IsNumeric(ClassID) and ClassID<>"" then
	response.Write("传递参数出错！")
	response.End()
end if

sqlstr = "select id,NewsTitle,PostTime from NewsInfo where NewsLock=0 and ClassID="&ClassID&"  order by NewsOrder desc "
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
	<li>
		<a href="newsviewer.asp?id=<%=id%>" target="_self">
			<img src="./images/201203001.png" />
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
<center>

	<a href='<%=request.servervariables("URL")%>?page=1'>首页</a>
	<a href='<%=linkPrev%>'>上一页</a>
	第&nbsp;<%=CurrentPage%>&nbsp;/&nbsp;<%=rs.pagecount%>&nbsp;页
	<a href='<%=linkNext%>'>下一页</a>
	<a href='<%=request.servervariables("URL")%>?page=<%=rs.pagecount%>'>尾页</a>

</center>
	</div><!-- main div -->

</div><!-- /page -->
<%
rs.close
set rs = nothing
%>
</body>
</html>
