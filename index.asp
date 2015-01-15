﻿<!DOCTYPE html>
<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title>自助终端</title>
	<link rel="shortcut icon" href="./favicon.ico">
	<link rel="stylesheet" href="./css/themes/default/jquery.mobile-1.4.5.min.css">
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
	<style>
		.ui-content {
			max-width: 600px;
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
			font-size: 1.4em !important;
			background-color: yellow;
		}
		h1 {
			font-family: 微软雅黑, 宋体 !important;
			font-size: 3em !important;
			text-align: center;
		}
    </style>
</head>
<body>
<div data-role="page" class="jqm-demos" data-quicklinks="true">
	<div role="main" class="ui-content jqm-content">
		<h1>自助系统</h1>
		<ul data-role="listview" data-inset="true">
			<li><a href="#">排队叫号</a></li>
			<li><a href="#">气费查询</a></li>
			<li><a href="newslist.asp?id=1" target="_self">新闻杂志</a></li>
			<li><a href="#">产品介绍</a></li>
			<li><a href="#">热门链接</a></li>
		</ul>
	</div>
<!--
<p align="center"><a href="./admin/login.asp">进入后台</a></p>
-->
</div><!-- /page -->

</body>
</html>