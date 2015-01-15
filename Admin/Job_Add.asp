<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config2.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#Include File="../Editor/fckeditor.asp"-->
<%
Call ISPopedom(UserName,"News_Content")
Action=ReplaceBadChar(Trim(Request("Action")))
ClassID=ReplaceBadChar(Trim(Request("ClassID")))
If ClassID="" Then
	ClassID="0"
End If
If Action="Save" Then
	NewsTitle=ReplaceBadChar(Trim(Request("NewsTitle")))
	ClassID=ReplaceBadChar(Trim(Request("ClassID")))
	NewsSPic=Trim(Request("NewsSPic"))
	NewsBPic=Trim(Request("NewsBPic"))
	NewsAuthor=ReplaceBadChar(Trim(Request("NewsAuthor")))
	Keywords=ReplaceBadChar(Trim(Request("Keywords")))
	NewsContent=Trim(Request("NewsContent"))
	NewsLock=ReplaceBadChar(Trim(Request("NewsLock")))	
	NewsVisit=ReplaceBadChar(Trim(Request("NewsVisit")))
	PostTime=Trim(Request("PostTime"))
	If PostTime="" Or IsDate(PostTime)=false Then
		PostTime=Now()
	End If
	
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select NewsOrder From NewsInfo Order By NewsOrder Desc"
	Rs.Open Sql,Conn,1,1
	If Not (Rs.Eof Or Rs.Bof) Then
		NewsOrder=Cstr(Trim(Rs("NewsOrder")))+1
	Else
		NewsOrder=1
	End If
	Rs.Close
	Set Rs=Nothing
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From NewsInfo"
	Rs.Open Sql,Conn,1,3
	Rs.AddNew
	Rs("NewsTitle")=NewsTitle
	Rs("ClassID")=ClassID
	Rs("NewsSPic")=NewsSPic
	Rs("NewsBPic")=NewsBPic
	Rs("NewsAuthor")=NewsAuthor
	Rs("Keywords")=Keywords
	Rs("NewsContent")=NewsContent
	Rs("NewsClick")=1
	Rs("NewsLock")=NewsLock
	Rs("NewsOrder")=NewsOrder
	Rs("NewsVisit")=NewsVisit
	Rs("NewsIndex")=0
	Rs("PostTime")=PostTime
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u4fe1\u606f\u5185\u5bb9\u6dfb\u52a0\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='News_ContentAdd.asp?ClassID="&ClassID&"';</script>")
	Response.End()
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<link href="Style/Main.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="Common/Jquery.js"></script>
<script language="javascript" type="text/javascript" src="Common/Common.js"></script>
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：<a href="Job_List.asp">职位管理</a> >> 添加职位</td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">信息内容你可以设置会员阅读权限，仅网站注册会员方能阅读该信息。<br>
注意：不想对外发布的信息你可以设置成锁定状态。</td>
</tr>
</table></td>
</tr>
<tr>
<td colspan="2" valign="top">
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#NewsTitle").val()=="")
	{
		alert("\u4fe1\u606f\u6807\u9898\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#NewsTitle").focus();
		return false;
	}
	return true;	
}
</script>
<form id="form1" name="form1" method="post" action="?Action=Save" onSubmit="return CheckForm();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
<th colspan="3">添加广告</th>
</tr>
<tr>
<td class="Right" align="right" width="25%">广告标题：</td>
<td width="75%" colspan="2"><input type="text" id="NewsTitle" name="NewsTitle" value="" class="Input300px"/></td>
</tr>
<tr>
<td class="Right" align="right">广告说明：</td>
<td colspan="2"><input type="text" id="NewsSPic" name="NewsSPic" value="" class="Input300px"/></td>
</tr>
<tr>
<td class="Right" width="25%" align="right">图片内容：</td>
<td width="38%"><input type="text" id="NavPicture" name="NavPicture" class="Input300px"></td>
<td width="37%"><input type="button" value="浏览图片" class="Button" onclick="OpenImageBrowser('NavPicture');"/></td>
</tr>
<tr>
<td class="Right" align="right">链接地址：</td>
<td colspan="2"><input type="text" id="NewsSPic" name="NewsSPic" value="" class="Input300px"/></td>
</tr>
<tr>
<td class="Right" align="right">打开方式：</td>
<td colspan="2">
<select id="NavTarget" name="NavTarget" style="width:120px">
<option value="">原窗口</option>
<option value="_blank">新窗口</option>
</select>
</td>
</tr>
<tr>
<td class="Right" align="right">发布状态：</td>
<td colspan="2"><input type="radio" id="NavLock" name="NavLock" value="0" checked="checked"/>已发布<input type="radio" id="NavLock" name="NavLock" value="1"/>未发布</td>
</tr>
<tr>
<td class="Right" align="right">&nbsp;</td>
<td colspan="2"><input type="submit" value="保 存" class="Button"> <input type="button" value="返 回" class="Button" onClick="window.location.href='AD_Banner.asp'"></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>