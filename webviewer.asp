<!DOCTYPE html>
<%
URL		= trim(request.QueryString("url"))
title	= trim(request.QueryString("title"))
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
	<script src="./js/webviewer.js"></script>
<style>
.ui-content {
	padding: 0 !important;
}
</style>
</head>
<body>
  <div data-role="page">
	<div role="main" class="ui-content">
		<iframe id="contentFrame" name="contentFrame" width="100%" src="<%=URL%>" frameborder="0"></iframe>
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