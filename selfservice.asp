<!DOCTYPE html>
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
		padding: 0 !important;
		overflow-y: none;
	}
	</style>
</head>
<body>
<div data-role="page" class="" data-quicklinks="true">
	<div role="main" class="ui-content">
<iframe src="./newslist.asp?id=<%=Request.QueryString("ID")%>" width="100%" frameborder="0">
</iframe>
	</div><!-- /main -->
</div><!-- /page -->
</body>
<script>
	$(".ui-content").height($(document).height());
	$("iframe").height($(document).height());
</script>
</html>
