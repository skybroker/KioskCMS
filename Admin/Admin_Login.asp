﻿<!--#include File="../Config/Config2.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_Md5.asp"-->
<%
If Request("Action")="Login" Then

UserName=ReplaceBadChar(Trim(Request("UserName")))
UserPass=ReplaceBadChar(Trim(request("UserPass")))

If UserName="" Or UserPass="" Then
Response.Write("<Script>alert('\u767b\u5f55\u5e10\u53f7\u53ca\u767b\u5f55\u5bc6\u7801\u4e0d\u80fd\u4e3a\u7a7a');history.go(-1);</Script>")
Response.End()
End If

UserPass=MD5(UserPass,32)
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From Admin Where UserName='"&UserName&"' And UserPass='"&UserPass&"'"
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof And Rs.Bof) Then
Response.Cookies("CNVP_CMS2")("UserName") = UserName
Response.Write("<script>top.window.location.href='Admin_Index.asp';</script>")
Else
Response.Write("<Script>alert('\u767b\u5f55\u5931\u8d25\uff0c\u767b\u5f55\u5e10\u53f7\u6216\u8005\u5bc6\u7801\u9519\u8bef');history.go(-1);</Script>")
Response.End()
End If
Rs.Close
Set Rs=Nothing
Conn.Close
Set Conn=Nothing

End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<link href="Style/Login.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="../Common/Flash.js"></script>
<script language="javascript" type="text/javascript" src="../Common/Jquery.js"></script>
</head>
<body scroll="no">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td align="center">
<table width="750" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="50" colspan="2" align="left">
<table width="750" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="500" align="left"><img src="Images/expo.jpg" border="0" /></td>
<td width="250" style="color:#669acc;font-size:14px; text-align:right">Expo CMS 1.0</td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan="2">
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#UserName").val()=="")
	{
		alert("\u7528\u6237\u540d\u4e0d\u80fd\u4e3a\u7a7a\uff0c\u8bf7\u8f93\u5165\u540e\u91cd\u65b0\u767b\u5f55\u7cfb\u7edf\u3002");
		$("#UserName").focus();
		return false;
	}
	if ($("#UserPass").val()=="")
	{
		alert("\u5bc6\u7801\u4e0d\u80fd\u4e3a\u7a7a\uff0c\u8bf7\u8f93\u5165\u540e\u91cd\u65b0\u767b\u5f55\u7cfb\u7edf\u3002");
		$("#UserPass").focus();
		return false;
	}
	return true;
}
</script>
<table width="750" height="300" border="0" cellspacing="0" cellpadding="0" style="border:solid 1px #FFFFFF;background:url(Images/da_bg.jpg) no-repeat">
<tr>
<td width="250" valign="top" align="center">
<form id="form1" name="form1" action="Admin_Login.asp?Action=Login" method="post" onSubmit="return CheckForm();">
<table width="210" border="0" cellpadding="0" cellspacing="0">
<tr>
<td height="50" colspan="2">&nbsp;</td>
</tr>
<tr>
<td height="30" colspan="2" align="left" style="font-size:14px;color:#FFFFFF;"><b>用户登录</b></td>
</tr>
<tr>
<td height="15" colspan="2"></td>
</tr>
<tr>
<td height="1" align="left" style="background:#666" colspan="2"></td>
</tr>
<tr>
<td height="15" colspan="2"></td>
</tr>
<tr>
<td width="60" height="40" style="color:#FFFFFF;font-size:14px;" align="left"><b>用户名</b></td>
<td width="140"><input type="text" id="UserName" name="UserName" style="width:150px;height:20px" maxlength="10"/></td>
</tr>
<tr>
<td height="40" style="color:#FFFFFF;font-size:14px;" align="left"><b>密&nbsp;&nbsp;&nbsp;码</b></td>
<td><input type="password" id="UserPass" name="UserPass" style="width:150px;height:20px;" maxlength="10"/></td>
</tr>
<tr>
<td height="15" colspan="2"></td>
</tr>
<tr>
<td height="30" style="color:#FFFFFF;font-size:14px;" align="left">&nbsp;</td>
<td align="right"><input type="submit" value="登  录" style="background:url(Images/Login_Btn.jpg) no-repeat;width:91px;height:24px;border:0px;color:#FFFFFF;font-weight:bolder;" /></td>
</tr>
</table>
</form>
</td>
<td width="500">&nbsp;</td>
</tr>
</table>
</td>
</tr>
<tr>
<td width="550" height="35" align="left" style="color:#669acc">CopyRight  &copy; 2005-<%=Year(Now())%> 网站信息管理系统 All Rights Reserved.</td>
<td width="300" align="right" style="color:#669acc">客服热线：010-68335476</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>