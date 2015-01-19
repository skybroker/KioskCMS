<!DOCTYPE html>
<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<%
ID	=	trim(request.QueryString("ID"))
if not IsNumeric(id) then
	response.Write("传递参数出错！")
	response.End()
end if
'------------------------------------------------------------------------------
conn.execute("update NewsInfo set NewsClick=NewsClick+1 where id="&id)
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql = "Select ClassID,NewsAuthor,NewsTitle,NewsBPic,NewsSPic,NewsContent,PostTime,NewsClick From NewsInfo Where NewsLock=0 and ID="&ID&""
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof Or Rs.Bof ) Then
	ClassID		=	rs("ClassID")
	NewsAuthor=	rs("NewsAuthor")
	NewsTitle	=	trim(rs("NewsTitle"))
	NewsContent	=	rs("NewsContent")
	PostTime	=	rs("PostTime")
	NewsClick	=	rs("NewsClick")
	NewsBPic	=	rs("NewsBPic")
	NewsSPic	=	rs("NewsSPic")
End If
Rs.Close
set rs = nothing
'-----------------------------------------------------------------------------

title	=	NewsTitle	'页面title
%>
<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title><%=title%></title>
	<link rel="shortcut icon" href="./favicon.ico">
	<link rel="stylesheet" href="./css/themes/default/jquery.mobile-1.4.5.min.css">
	<link rel="stylesheet" href="./css/kiosk.css">
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
	<script src="./js/idle-timer.min.js"></script>
	<script src="./js/newsviewer.js"></script>
</head>
<body>
  <div data-role="page">
    <div role="main" class="ui-content ui-content-1024">
      	<div class="title" align="center">
        	<h1><%=Title%></h1>
          作者：<%=NewsAuthor%>&nbsp;&nbsp;&nbsp;&nbsp;浏览次数：<%=NewsClick%>&nbsp;&nbsp;&nbsp;&nbsp;发布时间：<%=PostTime%>
        </div>
        <%if NewsBPic<>"" then%>
        <img src="<%=NewsBPic%>" border="0" class="img_content" align="middle"/>
        <%end if%>
        <%=NewsContent%>
     </div>
     <div data-role="footer" class="content-over-panel">
     	《---- 完 ----》
     </div>
     <div class="scroll-panel">
     	<div><a href="#" class="kiosk-element ui-btn" data-rel="back">返回</a></div>
     	<div><a href="#" class="kiosk-element scroll-up ui-btn">向上滚动</a></div>
     	<div><a href="#" class="kiosk-element scroll-down ui-btn">向下滚动</a></div>
     </div>
  </div>
</body>
<script src="./js/idle-back-home.js"></script>
</html>