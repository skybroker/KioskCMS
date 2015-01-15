<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config2.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<link href="Style/Main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td width="60%" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td style="font-size:14px;font-weight:bolder;color:#224d7e;height:35px">欢迎登录网站管理系统！</td>
</tr>
<!--<tr>
<td style="color:#808080">Expo CMS企业网站管理系统已全面升级，您现在使用的是<span style="color:#F90;">CMS 2.0正式版</span>，如果您有何疑问请致电Expo CMS客服部进行咨询。
</td>
</tr>
--><tr>
<td height="35">
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="20%">空间使用情况：</td>
<td width="80%">
<div style="float:left;width:300px;height:15px;background:url(Images/Loading_Bg.gif);border:1px solid #426a98;">
<div style="margin:0px;padding:0px;">
<div style="position:absolute;text-align:center;height:13px;line-height:13px;width:300px;color:#FFFFFF; margin:0px;padding:0px;font-size:9px;">60MB/100MB(60%)</div>
<div style="background:url(images/loading_bg2.gif);width:60%;height:13px;line-height:13px;"></div>
</div>
</div>
</td>
</tr>
</table>-->
</td>
</tr>
<tr>
<td height="250" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="25">
<span style="height:25px;width:80px; line-height:25px; text-align:center;background:#e2e7eb; font-weight:bolder">主机参数</span></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="235" bgcolor="#e2e7eb" style="padding:5px;">
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border:solid 1px #cdd8df;height:100%; background:#f1f4f7">
<tr>
<td align="center">
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%">服务器名：</td>
    <td width="25%"><%=Request.ServerVariables("SERVER_NAME")%></td>
    <td width="20%">服务器端口：</td>
    <td width="30%"><%=Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
  <tr>
    <td>服务器IP：</td>
    <td><%=Request.ServerVariables("LOCAL_ADDR")%></td>
    <td>服务器时间：</td>
    <td><%=now%></td>
  </tr>
  <tr>
    <td>IIS版本：</td>
    <td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td>脚本超时时间：</td>
    <td><%=Server.ScriptTimeout%> 秒</td>
  </tr>
  <tr>
    <td>Application变量：</td>
    <td><%Response.Write(Application.Contents.Count & "个")%></td>
    <td>Session变量：</td>
    <td><%Response.Write(Session.Contents.Count&"个 ")%></td>
  </tr>
  <tr>
    <td>所有服务器参数：</td>
    <td><%Response.Write(Request.ServerVariables.Count&"个")%></td>
    <td>服务器环境变量：</td>
    <td><%
On Error Resume Next
dim WshShell,WshSysEnv
Set WshShell = server.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment
if err then
Response.Write("服务器不支持WScript.Shell组件")
err.clear
else
Response.Write(WshSysEnv.count &"个 ")
if WshSysEnv.count>0 then Response.Write("[<a href=""?action=showwsh"">遍历环境变量</a>]") 
end if
%></td>
  </tr>
  <tr>
    <td>本文件实际路径：</td>
    <td colspan="3"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
    </tr>
</table></td>
</tr>
</table>
</td>
</tr>
</table></td>
</tr>
</table>
</td>
<td width="10" valign="top" align="right">
<td width="38%" valign="top">
<table width="300" border="0" cellspacing="0" cellpadding="0" style="border:solid 1px #c5cdd4;">
<tr style="background:url(Images/Info-List-Bg.gif);font-weight:bolder;height:32px; line-height:32px">
<td>&nbsp;&nbsp;最新案例</td>
<td align="right"><a href="#" style="color:#000000; font-weight:normal" target="_blank"x>[更多]</a>&nbsp;&nbsp;</td>
</tr>
<tr>
<td colspan="2" valign="top" style="background:#f4f6f9;height:336px">&nbsp;</td>
</tr>
</table>
</td>
</tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="20">&nbsp;</td>
</tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="background:#cdd8df;height:1px"></td>
</tr>
</table>

</body>
</html>