﻿<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->
<%
Call ISPopedom(UserName,"Pro_ContentListOut")
Action=ReplaceBadChar(Trim(Request("Action")))
ShopName=ReplaceBadChar(Trim(Request("ShopName")))
ClassID=ReplaceBadChar(Trim(Request("ClassID")))
If ClassID="" Then ClassID=0
ShowType=ReplaceBadChar(Trim(Request("ShowType")))
If ShowType="" Then ShowType=0
ID=ReplaceBadChar(Trim(Request("ID")))

Select Case Action
Case "Up"
	MaxOrder=Conn.Execute("Select Max(ShopOrder) From ShopInfo")(0)
	ShopOrder=ReplaceBadChar(Trim(Request("ShopOrder")))
	If Cstr(ShopOrder)=Cstr(MaxOrder) Then
		Response.Write("<script>alert('\u8be5\u4fe1\u606f\u5df2\u7ecf\u5904\u4e8e\u6700\u5934\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update ShopInfo Set ShopOrder="&ShopOrder&" Where ShopOrder="&ShopOrder+1&"")
		Conn.Execute("Update ShopInfo Set ShopOrder="&ShopOrder+1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "Down"
	MinOrder=Conn.Execute("Select Min(ShopOrder) From ShopInfo")(0)
	If MinOrder="" Then MinOrder="1"
	ShopOrder=ReplaceBadChar(Trim(Request("ShopOrder")))
	If Cstr(ShopOrder)=Cstr(MinOrder) Then
		Response.Write("<script>alert('\u8be5\u4fe1\u606f\u5df2\u7ecf\u5904\u4e8e\u6700\u5e95\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update ShopInfo Set ShopOrder="&ShopOrder&" Where ShopOrder="&ShopOrder-1&"")
		Conn.Execute("Update ShopInfo Set ShopOrder="&ShopOrder-1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "UnIndex"
	Conn.Execute("Update ShopInfo Set ShopIndex=1 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "InIndex"
	Conn.Execute("Update ShopInfo Set ShopIndex=0 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "UnLock"
	Page=ReplaceBadChar(Trim(Request("Page")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Update ShopInfo Set ShopLock=0 Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u4ea7\u54c1\u4e0a\u67b6\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&ClassID="&ClassID&"&ShowType="&ShowType&"';</script>")
	Response.End()
Case "Delete"
	Page=ReplaceBadChar(Trim(Request("Page")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Delete From ShopInfo Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u5220\u9664\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&ClassID="&ClassID&"&ShowType="&ShowType&"';</script>")
	Response.End()
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<link href="Style/Main.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="Common/Jquery.js"></script>
<script language="javascript" type="text/javascript" src="Common/Common.js"></script>
<script language="javascript" type="text/javascript">
$(document).ready(function(){
	$("#ClassID").val("<%=ClassID%>");
	$("#ClassID").change(function(){
		window.location.href='?ClassID='+$("#ClassID").val()+'&ShowType=<%=ShowType%>&ShopName=<%=Server.URLEncode(ShopName)%>';
	});
});
</script>
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:60%">当前位置：<a href="Pro_ContentListOut.asp">下架产品维护</a>
<%
If ShopName<>"" Then
	StrSql=" And ShopName Like '%"&ShopName&"%' "
	Response.Write(">> 按产品名称查找，关键字为"&ShopName&"")
End If
%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a href="Pro_ContentListOut.asp">页面刷新</a></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:40%; text-align:right">
<select id="ClassID" name="ClassID" style="width:300px;">
<option value="0">|--所有下架产品</option>
<%=GetSelect("ShopClass",0)%>
</select>&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
  <td width="75%" valign="top">以下为所有下架产品的信息列表，如果想对外发布产品您可以进行上架操作。<br>
    注意：您可以切换显示布局，目前共有列表、图片两个布局模式。</td>
  <td width="15%" valign="bottom" style="text-align:right">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td>选择布局：</td>
  <td><a href="<%=GetURL("showtype")%>0">
  <%
If showtype=0 Then
PageSize=10
Response.Write("<img src=""Images/List_02.gif"" width=""18"" height=""15"" border=""0""/>")
Else
Response.Write("<img src=""Images/List_01.gif"" width=""18"" height=""15"" border=""0""/>")
End If
%>
  </a></td>
  <td><a href="<%=GetURL("showtype")%>1">
  <%
If showtype=1 Then
PageSize=60
Response.Write("<img src=""Images/Image_02.gif"" width=""18"" height=""15"" border=""0""/>")
Else
Response.Write("<img src=""Images/Image_01.gif"" width=""18"" height=""15"" border=""0""/>")
End If
%>
  </a></td>
  </tr>
  </table>
  </td>
  </tr>
</table>
</td>
</tr>
<tr>
<td colspan="2" valign="top">
<form id="form1" name="form1" method="post">
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From ShopInfo Where ShopLock=1 "&StrSql&" And ClassID In ("&ClassID&GetAllChild("ShopClass",ClassID)&") Order By ShopOrder Desc"
Rs.Open Sql,Conn,1,1
Dim Page
Page=Request("Page")                    
Rs.PageSize = PageSize
Total=Rs.RecordCount
PGNum=Rs.PageCount
If Page="" Or clng(Page)<1 Then Page=1
If Clng(Page) > PGNum Then Page=PGNum
If PGNum>0 Then Rs.AbsolutePage=Page                         
i=0
%>
<%
Select Case ShowType
Case 0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form" id="GridView1">
<tr>
<th width="5%" class="Right"><input type="checkbox" name="chkSelectAll" onClick="doCheckAll(this)"></th>
<th width="30%" class="Right">产品名称<a href="#" onClick="top.CreateNewTab('Pro_ContentAdd.asp?ClassID=<%=ClassID%>','Pro_ContentAdd','添加产品')">[+]</a></th>
<th width="25%" class="Right">产品类别</th>
<th width="10%" class="Right">产品排序</th>
<th width="10%" class="Right">首页推荐</th>
<th width="10%">管理操作</th>
</tr>
<%Do While Not Rs.Eof And i<Rs.PageSize%>
<tr>
<td class="Right"><input type="checkbox" name="ID" value="<%=Rs("ID")%>"></td>
<td class="Right"><a href="Pro_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&FileName=Pro_ContentListOut&ShowType=<%=ShowType%>&ShopName=<%=Server.URLEncode(ShopName)%>"><%=Rs("ShopName")%></a></td>
<td class="Right"><%=GetSubNavName("ShopClass",Rs("ClassID"))%>&nbsp;</td>
<td class="Right"><a href="?Action=Up&ID=<%=Rs("ID")%>&ShopOrder=<%=Rs("ShopOrder")%>">上移</a>┊<a href="?Action=Down&ID=<%=Rs("ID")%>&ShopOrder=<%=Rs("ShopOrder")%>">下移</a></td>
<td class="Right">
<%
If Rs("ShopIndex")="0" Then
Response.Write("<a href=""?Action=UnIndex&ID="&Rs("ID")&""">首页推荐</a>")
Else
Response.Write("<a href=""?Action=InIndex&ID="&Rs("ID")&""" style=""color:red"">取消推荐</a>")
End If
%></td>
<td><a href="Pro_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&FileName=Pro_ContentListOut&ShowType=<%=ShowType%>&ShopName=<%=Server.URLEncode(ShopName)%>">编辑</a>┊<a href="?Action=Delete&ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>" onClick="if(!confirm('\u786e\u8ba4\u8981\u5c06\u8be5\u5546\u54c1\u4fe1\u606f\u5220\u9664\u5417\uff1f')) return false;">删除</a></td>
</tr>
<%
i=i+1
Rs.MoveNext
Loop
%>
<tr>
<th colspan="2" style="font-weight:normal">操作：<a href="javascript:UnLock();" style="font-weight:normal">上架</a>&nbsp;┊&nbsp;<a href="javascript:Delete();" style="font-weight:normal">删除</a>&nbsp;┊&nbsp;<a href="javascript:ChangeParent();" style="font-weight:normal">转移</a>&nbsp;┊&nbsp;<a href="Pro_ContentListOutFind.asp?ClassID=<%=ClassID%>&ShowType=<%=ShowType%>" style="font-weight:normal">查找</a></th>
<th colspan="3" style="font-weight:normal;text-align:right">共<%=Rs.PageCount%>页&nbsp;第<%=Page%>页&nbsp;<%=PageSize%>条/页&nbsp;共<%=Total%>条&nbsp;
<%if Page=1 then%>
首 页&nbsp;上一页&nbsp;
<%Else%>
<a href="<%=GetUrl("page")%>1">首 页</a>&nbsp;<a href="<%=GetUrl("page")%><%=Page-1%>">上一页</a>&nbsp;
<%End If%>
<%If Rs.PageCount-Page<1 Then%>下一页&nbsp;尾 页&nbsp;
<%Else%><a href="<%=GetUrl("page")%><%=Page+1%>">下一页</a>&nbsp;<a href="<%=GetUrl("page")%><%=Rs.PageCount%>">尾 页</a>&nbsp;
<%End If%>
</th>
<th width="10%">
<select style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体;width:90%;" onChange="location=this.options[this.selectedIndex].value" name="Menu_1"> 
<%For Pagei=1 To Rs.PageCount%>
<%if Cint(Pagei)=Cint(Page) Then%>
<option value="<%=GetUrl("page")%><%=Pagei%>" selected="selected">第<%=Pagei%>页</option>
<%Else%>
<option value="<%=GetUrl("page")%><%=Pagei%>">第<%=Pagei%>页</option>
<%End If%>
<%Next%>
</select>
</th>
</tr>
</table>
<%Case 1%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
<td colspan="6" style="padding:5px 0px">
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="GridView1">
<tr>
<%Do While Not Rs.Eof And i<Rs.PageSize%>
<td style="border:0px;">
<a href="Pro_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&FileName=Pro_ContentListIn&ShowType=<%=ShowType%>"><img src="<%=Rs("ShopSPic")%>" width="120" height="120" border="0"/></a><br />
<input type="checkbox" name="ID" value="<%=Rs("ID")%>"><a href="Pro_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>&FileName=Pro_ContentListOut&ShowType=<%=ShowType%>&ShopName=<%=Server.URLEncode(ShopName)%>"><%=Rs("ShopName")%></a>
</td>
<%
i=i+1
If I Mod 6=0 Then
Response.Write("</tr><tr>")
End If
Rs.MoveNext
Loop
%>
</tr>
</table>
</td>
</tr>
<tr>
<th colspan="2" style="font-weight:normal">操作：<a href="javascript:UnLock();" style="font-weight:normal">上架</a>&nbsp;┊&nbsp;<a href="javascript:Delete();" style="font-weight:normal">删除</a>&nbsp;┊&nbsp;<a href="javascript:ChangeParent();" style="font-weight:normal">转移</a>&nbsp;┊&nbsp;<a href="Pro_ContentListOutFind.asp?ClassID=<%=ClassID%>&ShowType=<%=ShowType%>" style="font-weight:normal">查找</a></th>
<th colspan="3" style="font-weight:normal;text-align:right">共<%=Rs.PageCount%>页&nbsp;第<%=Page%>页&nbsp;<%=PageSize%>条/页&nbsp;共<%=Total%>条&nbsp;
<%if Page=1 then%>
首 页&nbsp;上一页&nbsp;
<%Else%>
<a href="<%=GetUrl("page")%>1">首 页</a>&nbsp;<a href="<%=GetUrl("page")%><%=Page-1%>">上一页</a>&nbsp;
<%End If%>
<%If Rs.PageCount-Page<1 Then%>下一页&nbsp;尾 页&nbsp;
<%Else%><a href="<%=GetUrl("page")%><%=Page+1%>">下一页</a>&nbsp;<a href="<%=GetUrl("page")%><%=Rs.PageCount%>">尾 页</a>&nbsp;
<%End If%>
</th>
<th width="10%">
<select style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体;width:90%;" onChange="location=this.options[this.selectedIndex].value" name="Menu_1"> 
<%For Pagei=1 To Rs.PageCount%>
<%if Cint(Pagei)=Cint(Page) Then%>
<option value="<%=GetUrl("page")%><%=Pagei%>" selected="selected">第<%=Pagei%>页</option>
<%Else%>
<option value="<%=GetUrl("page")%><%=Pagei%>">第<%=Pagei%>页</option>
<%End If%>
<%Next%>
</select>
</th>
</tr>
</table>
<%End Select%>
</form>
</td>
</tr>
</table>
<script language="javascript" type="text/javascript">
function UnLock()
{
	var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
	if (confirm('\u786e\u5b9a\u8981\u5c06\u9009\u4e2d\u7684\u5546\u54c1\u8fdb\u884c\u4e0a\u67b6\u64cd\u4f5c\u5417\uff1f')) {
        window.location.href = '?Action=UnLock&ID='+l+'&Page=<%=Page%>&ClassID=<%=ClassID%>';
    }
}
function ChangeParent()
{
	var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u5b9a\u8981\u66f4\u6539\u9009\u4e2d\u5546\u54c1\u7684\u6240\u5c5e\u7236\u7c7b\u522b\u5417\uff1f')) {
        window.location.href = 'Pro_ContentParent.asp?FileName=Pro_ContentListOut&ID='+l;
    }
}
function Delete() {
    var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u8ba4\u8981\u5c06\u9009\u4e2d\u7684\u5546\u54c1\u4fe1\u606f\u5220\u9664\u5417\uff1f')) {
        window.location.href = '?Action=Delete&ID='+l+'&Page=<%=Page%>&ClassID=<%=ClassID%>';
    }
}
</script>
</body>
</html>