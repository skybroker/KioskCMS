﻿<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#Include File="../Editor/fckeditor.asp"-->
<%
Call ISPopedom(UserName,"Nav_Picture")
ID=ReplaceBadChar(Trim(Request("ID")))
Page=ReplaceBadChar(Trim(Request("Page")))
ParentID=ReplaceBadChar(Trim(Request("ParentID")))
Action=ReplaceBadChar(Trim(Request("Action")))
If Action="Save" Then
	NavTitle=ReplaceBadChar(Trim(Request("NavTitle")))
	NavRemark=ReplaceBadChar(Trim(Request("NavRemark")))
	NavPicture=Trim(Request("NavPicture"))
	NavContent=Trim(Request("NavContent"))
	NavLock=ReplaceBadChar(Trim(Request("NavLock")))	
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From SitePicture Where ID="&ID&""
	Rs.Open Sql,Conn,1,3
	Rs("NavTitle")=NavTitle
	Rs("NavRemark")=NavRemark
	Rs("NavPicture")=NavPicture
	Rs("NavContent")=NavContent
	Rs("NavLock")=NavLock
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u56fe\u7247\u5185\u5bb9\u4fee\u6539\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='Nav_Picture.asp?Page="&Page&"&ID="&ParentID&"';</script>")
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
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：<a href="Nav_Picture.asp">友情链接维护</a> >> 编辑友情链接</td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下所有项目均不能为空，请准确真实的填写相关信息。<br>
注意：友情链接名称和图片必填。可根据实际情况添加说明内容。</td>
</tr>
</table></td>
</tr>
<tr>
<td colspan="2" valign="top">
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#NavTitle").val()=="")
	{
		alert("\u56fe\u7247\u5185\u5bb9\u540d\u79f0\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#NavTitle").focus();
		return false;
	}
	if ($("#NavRemark").val()=="")
	{
		alert("\u56fe\u7247\u5185\u5bb9\u8bf4\u660e\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#NavRemark").focus();
		return false;
	}
	if ($("#NavPicture").val()=="")
	{
		alert("\u56fe\u7247\u5185\u5bb9\u56fe\u7247\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#NavPicture").focus();
		return false;
	}
}
</script>
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From SitePicture Where ID="&ID&""
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof Or Rs.Bof) Then
	NavTitle=Rs("NavTitle")
	NavRemark=Rs("NavRemark")
	NavPicture=Rs("NavPicture")
	NavContent=Rs("NavContent")
	NavLock=Rs("NavLock")
End If
Rs.Close
Set Rs=Nothing
%>
<form id="form1" name="form1" method="post" action="?Action=Save" onSubmit="return CheckForm();">
<input type="hidden" id="ID" name="ID" value="<%=ID%>"/>
<input type="hidden" id="Page" name="Page" value="<%=Page%>"/>
<input type="hidden" id="ParentID" name="ParentID" value="<%=ParentID%>"/>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
<th colspan="3">编辑图片内容</th>
</tr>
<tr>
<td class="Right" width="25%" align="right">友情链接名称：</td>
<td width="75%" colspan="2"><input type="text" id="NavTitle" name="NavTitle" value="<%=NavTitle%>" class="Input300px"/></td>
</tr>
<tr>
<td class="Right" width="25%" align="right">友情链接地址：</td>
<td width="75%" colspan="2"><input type="text" id="NavRemark" name="NavRemark" value="<%=NavRemark%>" class="Input300px"/> 
* 示例： www.zhujianqiang.com</td>
</tr>
<tr>
<td class="Right" width="25%" align="right">友情链接图片：</td>
<td width="38%"><input type="text" id="NavPicture" name="NavPicture" value="<%=NavPicture%>" class="Input300px"></td>
<td width="37%"><input type="button" value="浏览图片" class="Button" onclick="OpenImageBrowser('NavPicture');"/></td>
</tr>
<tr style="display:none">
<td class="Right" width="25%" align="right" valign="top">图片详细内容：</td>
<td width="75%" colspan="2"><%=Editor("NavContent",NavContent)%><br /><span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><script src="AutoSave.asp?Action=AutoSave&FrameName=NavContent"></script>
</td>
</tr>
<tr style="display:none">
<td class="Right" width="25%" align="right">导航条状态：</td>
<td width="75%" colspan="2"><input type="radio" id="NavLock" name="NavLock" value="0"<%If NavLock="0" Then Response.Write(" checked=""checked""")%>/>已发布<input type="radio" id="NavLock" name="NavLock" value="1"<%If NavLock="1" Then Response.Write(" checked=""checked""")%>/>未发布</td>
</tr>
<tr>
<td class="Right" width="25%" align="right">&nbsp;</td>
<td width="75%" colspan="2"><input type="submit" value="保 存" class="Button"> <input type="button" value="返 回" class="Button" onClick="window.location.href='Nav_Picture.asp?ID=<%=ParentID%>&Page=<%=Page%>'"></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>