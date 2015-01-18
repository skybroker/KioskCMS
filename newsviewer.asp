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
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
	<script src="./js/idle-timer.min.js"></script>
	<script src="./js/newsviewer.js"></script>
	<style>
		.ui-content {
			max-width: 1024px;
			margin: 0 auto 0 auto;
		}
		body {
			background-position: left top;
			background-image: url(images/home_bg.jpg);
			background-repeat: no-repeat;
			background-size: 100%;
			background-attachment: fixed;
		}
		#closing-panel {
			display: none;
			background-color: #eeeeee;
			padding: 2em;
		}
		.scroll-bar {
			position: fixed;
			left: 10px;
			top: 20%;
			width: 100px;
			height: 200px;
		}
		.scroll-bar a {
			width: 100%;
			height: 100px;
			line-height: 100px;
			background-color: transparent !important;
			font-size: 1.5em;
		}
    </style>
</head>
<body>
  <div data-role="page">
    <div role="main" class="ui-content">
      	<div class="title" align="center">
        	<h1><%=Title%></h1>
          作者：<%=NewsAuthor%>&nbsp;&nbsp;&nbsp;&nbsp;浏览次数：<%=NewsClick%>&nbsp;&nbsp;&nbsp;&nbsp;发布时间：<%=PostTime%>
        </div>
        <%if NewsBPic<>"" then%>
        <img src="<%=NewsBPic%>" border="0" class="img_content" align="middle"/>
        <%end if%>
        <%=NewsContent%>
        <div id="closing-panel">
        	用户无操作，<span id="count-down-seconds">10</span>秒后返回主页!
        </div>
     </div>
     <div class="scroll-bar">
     	<div><a href="#" class="kiosk-element ui-btn" data-rel="back">返回</a></div>
     	<div><a href="#" class="kiosk-element scroll-up ui-btn">向上滚动</a></div>
     	<div><a href="#" class="kiosk-element scroll-down ui-btn">向下滚动</a></div>
     </div>
  </div>
</body>
<script src="./js/idle-back-home.js"></script>
</html>