<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<%
Call ISPopedom(UserName,"Sys_SiteInfo")
If Request("Action")="Save" Then
SiteName=ReplaceBadChar(Trim(Request("SiteName")))
EnSiteName=ReplaceBadChar(Trim(Request("EnSiteName")))
SiyeKeys=ReplaceBadChar(Trim(Request("SiyeKeys")))
EnSiyeKeys=ReplaceBadChar(Trim(Request("EnSiyeKeys")))
SiteDes=ReplaceBadChar(Trim(Request("SiteDes")))
EnSiteDes=ReplaceBadChar(Trim(Request("EnSiteDes")))
SiteICP=ReplaceBadChar(Trim(Request("SiteICP")))
EnSiteICP=ReplaceBadChar(Trim(Request("EnSiteICP")))
SiteCopy=ReplaceBadChar(Trim(Request("SiteCopy")))
EnSiteCopy=ReplaceBadChar(Trim(Request("EnSiteCopy")))
SiteAuthor=ReplaceBadChar(Trim(Request("SiteAuthor")))
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From SiteInfo"
Rs.Open Sql,Conn,1,3
Rs("SiteName")=SiteName
Rs("EnSiteName")=EnSiteName
Rs("SiyeKeys")=SiyeKeys
Rs("EnSiyeKeys")=EnSiyeKeys
Rs("SiteDes")=SiteDes
Rs("EnSiteDes")=EnSiteDes
Rs("SiteICP")=SiteICP
Rs("EnSiteICP")=EnSiteICP
Rs("SiteCopy")=SiteCopy
Rs("EnSiteCopy")=EnSiteCopy
Rs("SiteAuthor")=SiteAuthor
Rs.Update
Rs.Close
Set Rs=Nothing
Conn.Close
Set Conn=Nothing
Response.Write("<script>alert('\u7ad9\u70b9\u57fa\u672c\u4fe1\u606f\u4fee\u6539\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='Sys_SiteInfo.asp';</script>")
Response.End()
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<link href="Style/Main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px">当前位置：站点基本信息维护</td>
</tr>
<tr>
<td height="80">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">在这里你可以设置网站的基本信息，保存之后立即生效。<br>
注意：请谨慎操作，保存操作之后所有数据均不可恢复。</td>
</tr>
</table></td>
</tr>
<tr>
<td valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<form id="form1" name="form1" method="post" action="?action=Save">
<tr>
<th colspan="3">站点基本信息维护</th>
</tr>
<tr style="display:none">
<td class="Right" align="right">公司标志图片(200*30px)：</td>
<td width="33%"><img src="<%=SiteLogo%>" width="200" height="30"></td>
<td width="48%"><input type="button" value="更换图片" class="Button" onClick="window.location.href='Sys_ChangeLogo.asp'"></td>
</tr>
<tr>
<td width="19%" class="Right" align="right">站点名称(SiteName)：</td>
<td width="33%"><input type="text" id="SiteName" name="SiteName" value="<%=SiteName%>" class="Input300px"></td>
<td><input type="text" id="EnSiteName" name="EnSiteName" value="<%=EnSiteName%>" class="Input300px">
(英文版)</td>
</tr>
<tr>
<td class="Right" align="right">站点关键字(SiyeKeys)：</td>
<td><input type="text" id="SiyeKeys" name="SiyeKeys" value="<%=SiyeKeys%>" class="Input300px"></td>
<td><input type="text" id="EnSiyeKeys" name="EnSiyeKeys" value="<%=EnSiyeKeys%>" class="Input300px"> (英文版)</td>
</tr>
<tr>
<td class="Right" align="right">站点描述(SiteDes)：</td>
<td><input type="text" id="SiteDes" name="SiteDes" value="<%=SiteDes%>" class="Input300px"></td>
<td><input type="text" id="EnSiteDes" name="EnSiteDes" value="<%=EnSiteDes%>" class="Input300px"> (英文版)</td>
</tr>
<tr>
<td class="Right" align="right">站点描述(SiteCopy)：</td>
<td><input type="text" id="SiteCopy" name="SiteCopy" value="<%=SiteCopy%>" class="Input300px"></td>
<td><input type="text" id="EnSiteCopy" name="EnSiteCopy" value="<%=EnSiteCopy%>" class="Input300px"> (英文版)</td>
</tr>
<tr>
<td class="Right" align="right">站点描述(SiteAuthor)：</td>
<td><input type="text" id="SiteAuthor" name="SiteAuthor" value="<%=SiteAuthor%>" class="Input300px"></td>
<td>&nbsp;</td>
</tr>
<tr>
<td class="Right" align="right">ICP备案编号(SiteICP)：</td>
<td><input type="text" id="SiteICP" name="SiteICP" value="<%=SiteICP%>" class="Input300px"></td>
<td><input type="text" id="EnSiteICP" name="EnSiteICP" value="<%=EnSiteICP%>" class="Input300px"> (英文版)</td>
</tr>
<tr>
<td class="Right" align="right">&nbsp;</td>
<td colspan="2"><input type="submit" value="保 存" class="Button"> <input type="button" value="关闭窗口" class="Button" onClick="top.DeleteTabTitle('Sys_SiteInfo')"></td>
</tr>
</form>
</table>
</td>
</tr>
</table>
</body>
</html>