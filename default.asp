<!DOCTYPE html>
<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title>自助终端</title>
	<link rel="shortcut icon" href="./favicon.ico">
	<link rel="stylesheet" href="./css/themes/default/jquery.mobile-1.4.5.min.css">
	<link rel="stylesheet" href="./css/kiosk.css">
	<script src="./js/jquery.min.js"></script>
	<script src="./js/jquery.mobile-1.4.5.min.js"></script>
</head>
<body>
<div data-role="page" class="jqm-demos" data-quicklinks="true">
	<div role="main" class="ui-content-600">
		<div class="title-space">&nbsp;</div>
		<div class="main_title">
			<!--
			<h1>自助系统</h1>
			-->
			<div class="ui-grid-a">
				<div class="ui-block-a"><a href="ivkexe:1"><h1>开户</h1></a></div>
				<div class="ui-block-b"><a href="ivkexe:2"><h1>缴费</h1></a></div>
			</div>
			<div id="wait_msg">
				<p>打印中...</p>
			</div>
		</div>
		<ul data-role="listview" class="main-function-list" data-icon="false">
			<!--
			<li><a href="ivkexe:1">开户业务</a></li>
			<li><a href="ivkexe:2">缴费业务</a></li>
			<li><a href="newslist.asp?id=1" target="_self"><h2>公司介绍</h2></a></li>
			<li><a href="newslist.asp?id=2" target="_self"><h2>业务流程</h2></a></li>
			<li><a href="#"><h2>政策法规</h2></a></li>
			-->
			<li><a href="webviewer.asp?http%3A%2F%2Fwww.qq.com&title=公司介绍" target="_self"><h2>公司介绍</h2></a></li>
			<li><a href="webviewer.asp?http%3A%2F%2Fwww.baidu.com&title=业务流程" target="_self"><h2>业务流程</h2></a></li>
			<li><a href="webviewer.asp?http%3A%2F%2Fnews.sogou.com&title=政策法规" target="_self"><h2>政策法规</h2></a></li>
		</ul>
	</div>
<!--
<p align="center"><a href="./admin/login.asp">进入后台</a></p>
-->
</div><!-- /page -->

</body>
<script src="./js/myjQueryEx.js"></script>
<script>
$(function(){
	$(".main_title a").click(function(){
		$('#wait_msg').center().show();
		setTimeout(function(){
			$('#wait_msg').hide();
		}, 2000);
	});
});
</script>
</html>