﻿<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#Include File="../Editor/fckeditor.asp"-->
<%
Call ISPopedom(UserName,"Pro_ContentAdd")
Action=ReplaceBadChar(Trim(Request("Action")))
ID=ReplaceBadChar(Trim(Request("ID")))
Page=ReplaceBadChar(Trim(Request("Page")))
ClassID=ReplaceBadChar(Trim(Request("ClassID")))
FileName=ReplaceBadChar(Trim(Request("FileName")))
ShowType=ReplaceBadChar(Trim(Request("ShowType")))
If Action="Save" Then
	ShopName=Trim(Request("ShopName"))
	ClassID=ReplaceBadChar(Trim(Request("ClassID")))
	ShopSPic=Trim(Request("ShopSPic"))
	ShopBPic=Trim(Request("ShopBPic"))
	PlaceofOrigin=ReplaceBadChar(Trim(Request("PlaceofOrigin")))
	ModelNumber=ReplaceBadChar(Trim(Request("ModelNumber")))
	MinimumOrderQuantity=ReplaceBadChar(Trim(Request("MinimumOrderQuantity")))
	PackagingDetails=ReplaceBadChar(Trim(Request("PackagingDetails")))
	DeliveryTime=ReplaceBadChar(Trim(Request("DeliveryTime")))
	PaymentTerms=ReplaceBadChar(Trim(Request("PaymentTerms")))
	SupplyAbility=ReplaceBadChar(Trim(Request("SupplyAbility")))
	ShopContent=Trim(Request("ShopContent"))
	ShopEnContent=Trim(Request("ShopEnContent"))
	ShopLock=ReplaceBadChar(Trim(Request("ShopLock")))	
	ShopVisit=ReplaceBadChar(Trim(Request("ShopVisit")))

	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From ShopInfo Where ID="&ID&""
	Rs.Open Sql,Conn,1,3
	Rs("ShopName")=ShopName
	Rs("ClassID")=ClassID
	Rs("ShopSPic")=ShopSPic
	Rs("ShopBPic")=ShopBPic
	Rs("PlaceofOrigin")=PlaceofOrigin
	Rs("ModelNumber")=ModelNumber
	Rs("MinimumOrderQuantity")=MinimumOrderQuantity
	Rs("PackagingDetails")=PackagingDetails
	Rs("DeliveryTime")=DeliveryTime
	Rs("PaymentTerms")=PaymentTerms
	Rs("SupplyAbility")=SupplyAbility
	Rs("ShopContent")=ShopContent
	Rs("ShopEnContent")=ShopEnContent
	Rs("ShopLock")=ShopLock
	Rs("ShopVisit")=ShopVisit
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u4ea7\u54c1\u4fe1\u606f\u7f16\u8f91\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='"&FileName&".asp?ClassID="&ClassID&"&Page="&Page&"&ShowType="&ShowType&"&ShopName="&Request("Keyword")&"';</script>")
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
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：产品管理 >> 编辑产品</td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下所有项目均不能为空，请准确真实的填写相关信息。<br>
注意：默认产品状态为产品展示(PRODUCTS)里，若要改变状态为现货供应(STORAGE)则请在其后选中。<br />产品最佳尺寸： 小图最佳尺寸（155px*155px）  大图（500px*500px）
</td>
</tr>
</table></td>
</tr>
<tr>
<td colspan="2" valign="top">
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From ShopInfo Where ID="&ID&""
Rs.Open Sql,Conn,1,1
If Not (Rs.Bof Or Rs.Eof) Then
	ClassID=Rs("ClassID")
	ShopName=Rs("ShopName")
	ShopSPic=Rs("ShopSPic")
	ShopBPic=Rs("ShopBPic")
	PlaceofOrigin=Rs("PlaceofOrigin")
	ModelNumber=Rs("ModelNumber")
	MinimumOrderQuantity=Rs("MinimumOrderQuantity")
	PackagingDetails=Rs("PackagingDetails")
	DeliveryTime=Rs("DeliveryTime")
	PaymentTerms=Rs("PaymentTerms")
	SupplyAbility=Rs("SupplyAbility")
	ShopContent=Rs("ShopContent")
	ShopEnContent=Rs("ShopEnContent")
	ShopLock=Rs("ShopLock")
	ShopVisit=Rs("ShopVisit")
End If
Rs.Close
Set Rs=Nothing
%>
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#ShopName").val()=="")
	{
		alert("\u4ea7\u54c1\u578b\u53f7\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#ShopName").focus();
		return false;
	}
	if ($("#ClassID").val()=="0")
	{
		alert("\u8bf7\u9009\u62e9\u7c7b\u522b");
		$("#ClassID").focus();
		return false;
	}
	return true;
}
$(document).ready(function(){
	$("#ClassID").val("<%=ClassID%>");
});
</script>
<form id="form1" name="form1" method="post" action="?Action=Save" onSubmit="return CheckForm();">
<input type="hidden" id="ID" name="ID" value="<%=ID%>"/>
<input type="hidden" id="Page" name="Page" value="<%=Page%>"/>
<input type="hidden" id="FileName" name="FileName" value="<%=FileName%>"/>
<input type="hidden" id="ShowType" name="ShowType" value="<%=ShowType%>"/>
<input type="hidden" id="Keyword" name="Keyword" value="<%=Request("ShopName")%>"/>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
<th colspan="4">编辑产品</th>
</tr>
<tr>
<td class="Right" align="right" width="20%">产品名称：</td>
<td class="Right" width="32%"><input type="text" id="ShopName" name="ShopName" value="<%=ShopName%>" class="Input200px"/>*</td>
<td class="Right" width="13%" align="right">产品类别：</td>
<td width="35%"><select id="ClassID" name="ClassID" style="width:300px;">
<option value="0">|--请选择栏目</option>
<%=GetSelect("ShopClass",0)%>
</select></td>
</tr>
<tr>
<td class="Right" align="right">产品小图：</td>
<td class="Right"><input type="text" id="ShopSPic" name="ShopSPic" value="<%=ShopSPic%>" class="Input200px"/> <a href="javascript:OpenImageBrowser('ShopSPic');">浏览...</a></td>
<td class="Right" align="right">产品大图：</td>
<td><input type="text" id="ShopBPic" name="ShopBPic" value="<%=ShopBPic%>" class="Input200px"/> <a href="javascript:OpenImageBrowser('ShopBPic');">浏览...</a></td>
</tr>
<tr>
<td class="Right" align="right">大小：</td>
<td class="Right"><input type="text" id="PlaceofOrigin" name="PlaceofOrigin" value="<%=PlaceofOrigin%>" class="Input200px"/></td>
<td class="Right" align="right">Size：</td>
<td><input type="text" id="ModelNumber" name="ModelNumber" value="<%=ModelNumber%>" class="Input200px"/></td>
</tr>
<tr style="display:none">
<td class="Right" align="right">简短概述：</td>
<td class="Right"><textarea name="MinimumOrderQuantity" cols="45" rows="6" id="MinimumOrderQuantity"><%=MinimumOrderQuantity%></textarea></td>
<td class="Right" align="right">简短概述（英文）：</td>
<td><textarea name="PackagingDetails" cols="45" rows="6" id="PackagingDetails"><%=PackagingDetails%></textarea></td>
</tr>
<tr style="display:none">
<td class="Right" align="right">Delivery Time：</td>
<td class="Right"><input type="text" id="DeliveryTime" name="DeliveryTime" value="<%=DeliveryTime%>" class="Input200px"/></td>
<td class="Right" align="right">Payment Terms：</td>
<td><input type="text" id="PaymentTerms" name="PaymentTerms" value="<%=PaymentTerms%>" class="Input200px"/></td>
</tr>
<tr style="display:none">
  <td class="Right" align="right" valign="top">Supply Ability：</td>
  <td colspan="3"><input type="text" id="SupplyAbility" name="SupplyAbility" value="<%=SupplyAbility%>" class="Input200px"/></td>
</tr>
<tr>
<td class="Right" align="right" valign="top">产品描述：</td>
<td colspan="3"><%=Editor("ShopContent",ShopContent)%><br /><span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><script src="AutoSave.asp?Action=AutoSave&FrameName=ShopContent"></script>
</td>
</tr>
<tr>
  <td class="Right" align="right" valign="top">产品描述(英文)：</td>
  <td colspan="3"><%=Editor("ShopEnContent",ShopEnContent)%><br /><span id="timemsg"></span><span id="msg2"></span><span id="msg"></span><script src="AutoSave.asp?Action=AutoSave&FrameName=ShopEnContent"></script></td>
</tr>
<tr style="display:none;">
<td class="Right" align="right">是否最新产品：</td>
<td class="Right"><input type="radio" id="ShopVisit" name="ShopVisit" value="0"<%If ShopVisit="0" Then Response.Write(" checked=""checked""")%>/>否<input type="radio" id="ShopVisit" name="ShopVisit" value="1"<%If ShopVisit="1" Then Response.Write(" checked=""checked""")%>/>是</td>
<td class="Right" align="right">产品状态：</td>
<td><input type="radio" id="ShopLock" name="ShopLock" value="0" checked="checked"<%If ShopLock="0" Then Response.Write(" checked=""checked""")%>/>产品目录(PRODUCTS)<input type="radio" id="ShopLock" name="ShopLock" value="1"<%If ShopLock="1" Then Response.Write(" checked=""checked""")%>/>现货供应(STORAGE)</td>
</tr>
<tr>
<td class="Right" align="right">&nbsp;</td>
<td colspan="3"><input type="submit" value="保 存" class="Button"> <input type="button" value="返回" class="Button" onClick="history.back();"></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>