﻿<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->
<%
Call ISPopedom(UserName,"News_List")
Action=ReplaceBadChar(Trim(Request("Action")))
NewsTitle=ReplaceBadChar(Trim(Request("NewsTitle")))
ClassID=ReplaceBadChar(Trim(Request("ClassID")))
If ClassID="" Then
	ClassID=0
End If
ID=ReplaceBadChar(Trim(Request("ID")))
Select Case Action
Case "Up"
	MaxOrder=Conn.Execute("Select Max(NewsOrder) From NewsInfo")(0)
	NewsOrder=ReplaceBadChar(Trim(Request("NewsOrder")))
	If Cstr(NewsOrder)=Cstr(MaxOrder) Then
		Response.Write("<script>alert('\u8be5\u4fe1\u606f\u5df2\u7ecf\u5904\u4e8e\u6700\u5934\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update NewsInfo Set NewsOrder="&NewsOrder&" Where NewsOrder="&NewsOrder+1&"")
		Conn.Execute("Update NewsInfo Set NewsOrder="&NewsOrder+1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "Down"
	MinOrder=Conn.Execute("Select Min(NewsOrder) From NewsInfo")(0)
	If MinOrder="" Then MinOrder="1"
	NewsOrder=ReplaceBadChar(Trim(Request("NewsOrder")))
	If Cstr(NewsOrder)=Cstr(MinOrder) Then
		Response.Write("<script>alert('\u8be5\u4fe1\u606f\u5df2\u7ecf\u5904\u4e8e\u6700\u5e95\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update NewsInfo Set NewsOrder="&NewsOrder&" Where NewsOrder="&NewsOrder-1&"")
		Conn.Execute("Update NewsInfo Set NewsOrder="&NewsOrder-1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "UnIndex"
	Conn.Execute("Update NewsInfo Set NewsIndex=1 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "InIndex"
	Conn.Execute("Update NewsInfo Set NewsIndex=0 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "IsLock"
	Page=ReplaceBadChar(Trim(Request("Page")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Update NewsInfo Set NewsLock=1 Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u5546\u54c1\u4e0b\u67b6\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&ClassID="&ClassID&"';</script>")
	Response.End()
Case "Delete"
	Page=ReplaceBadChar(Trim(Request("Page")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Delete From NewsInfo Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u5220\u9664\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&ClassID="&ClassID&"';</script>")
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
		window.location.href='?ClassID='+$("#ClassID").val()+'&NewsTitle=<%=Server.URLEncode(NewsTitle)%>';
	});
});
</script>
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：<a href="News_List.asp">信息内容维护</a>
<%
If NewsTitle<>"" Then
	StrSql=" And NewsTitle Like '%"&NewsTitle&"%' "
	Response.Write(">> 按商品名称查找，关键字为"&NewsTitle&"")
End If
%></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">
<select id="ClassID" name="ClassID" style="width:150px;">
<option value="0">|--所有信息列表</option>
<%=GetSelect("NewsClass",0)%>
</select>&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下为所有信息的列表，对于需要优先显示的信息你可以使用排序功能。<br>
注意：您可以进行添加、修改、删除等操作，保存之后立即生效。</td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan="2" valign="top">
<form id="form1" name="form1" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form" id="GridView1">
<tr>
<th width="5%" class="Right"><input type="checkbox" name="chkSelectAll" onClick="doCheckAll(this)"></th>
<th width="35%" class="Right">信息标题<a href="#" onClick="top.CreateNewTab('News_ContentAdd.asp?ClassID=<%=ClassID%>','News_ContentAdd','添加信息')">[+]</a></th>
<th width="15%" class="Right">信息类别</th>
<th width="15%" class="Right">信息排序</th>
<th width="15%" class="Right">首页推荐</th>
<th width="15%">管理操作</th>
</tr>
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select id,NewsTitle,NewsOrder,ClassID,NewsIndex From NewsInfo Where 1=1 "&StrSql&" And ClassID In ("&ClassID&GetAllChild("NewsClass",ClassID)&") Order By NewsOrder Desc"
Rs.Open Sql,Conn,1,1
Dim Page
Page=Request("Page")                            
PageSize = 10                                    
Rs.PageSize = PageSize               
Total=Rs.RecordCount               
PGNum=Rs.PageCount               
If Page="" Or clng(Page)<1 Then Page=1               
If Clng(Page) > PGNum Then Page=PGNum               
If PGNum>0 Then Rs.AbsolutePage=Page                         
i=0
Do While Not Rs.Eof And i<Rs.PageSize
%>
<tr>
<td class="Right"><input type="checkbox" name="ID" value="<%=Rs("ID")%>"></td>
<td class="Right"><a href="News_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>"><%=Rs("NewsTitle")%></a></td>
<td class="Right"><%=GetSubNavName("NewsClass",Rs("ClassID"))%>&nbsp;</td>
<td class="Right"><a href="?Action=Up&ID=<%=Rs("ID")%>&NewsOrder=<%=Rs("NewsOrder")%>">上移</a>┊<a href="?Action=Down&ID=<%=Rs("ID")%>&NewsOrder=<%=Rs("NewsOrder")%>">下移</a></td>
<td class="Right">
<%
If Rs("NewsIndex")="0" Then
Response.Write("<a href=""?Action=UnIndex&ID="&Rs("ID")&""">首页推荐</a>")
Else
Response.Write("<a href=""?Action=InIndex&ID="&Rs("ID")&""" style=""color:red"">取消推荐</a>")
End If
%></td>
<td><a href="News_ContentEdit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>">编辑</a>┊<a href="?Action=Delete&ID=<%=Rs("ID")%>&Page=<%=Page%>&ClassID=<%=ClassID%>" onClick="if(!confirm('\u786e\u8ba4\u8981\u5c06\u8be5\u4fe1\u606f\u5220\u9664\u5417\uff1f')) return false;">删除</a></td>
</tr>
<%
i=i+1
Rs.MoveNext
Loop
%>
<tr>
<th colspan="2" style="font-weight:normal">操作：<a href="javascript:Delete();" style="font-weight:normal">删除</a>&nbsp;┊&nbsp;<a href="javascript:ChangeParent();" style="font-weight:normal">转移</a>&nbsp;┊&nbsp;<a href="News_ContentListFind.asp?ClassID=<%=ClassID%>" style="font-weight:normal">查找</a></th>
<th colspan="3" style="font-weight:normal;text-align:right">共<%=Rs.PageCount%>页&nbsp;第<%=Page%>页&nbsp;<%=PageSize%>条/页&nbsp;共<%=Total%>条&nbsp;
<%if Page=1 then%>
首 页&nbsp;上一页&nbsp;
<%Else%>
<a href="?Page=1&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>">首 页</a>&nbsp;<a href="?Page=<%=Page-1%>&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>">上一页</a>&nbsp;
<%End If%>
<%If Rs.PageCount-Page<1 Then%>下一页&nbsp;尾 页&nbsp;
<%Else%><a href="?Page=<%=Page+1%>&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>">下一页</a>&nbsp;<a href="?Page=<%=Rs.PageCount%>&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>">尾 页</a>&nbsp;
<%End If%>
</th>
<th>
<select style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体;width:90%;" onChange="location=this.options[this.selectedIndex].value" name="Menu_1"> 
<%For Pagei=1 To Rs.PageCount%>
<%if Cint(Pagei)=Cint(Page) Then%>
<option value="?Page=<%=Pagei%>&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>" selected="selected">第<%=Pagei%>页</option>
<%Else%>
<option value="?Page=<%=Pagei%>&ClassID=<%=ClassID%>&NewsTitle=<%=Server.URLEncode(NewsTitle)%>">第<%=Pagei%>页</option>
<%End If%>
<%Next%>
</select>
</th>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script language="javascript" type="text/javascript">
function IsLock()
{
	var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
	if (confirm('\u786e\u5b9a\u8981\u5c06\u9009\u4e2d\u7684\u5546\u54c1\u8fdb\u884c\u4e0b\u67b6\u64cd\u4f5c\u5417\uff1f')) {
        window.location.href = '?Action=IsLock&ID='+l+'&Page=<%=Page%>&ClassID=<%=ClassID%>';
    }
}
function ChangeParent()
{
	var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u5b9a\u8981\u66f4\u6539\u9009\u4e2d\u4fe1\u606f\u7684\u6240\u5c5e\u7236\u7c7b\u522b\u5417\uff1f')) {
        window.location.href = 'News_ContentParent.asp?ID='+l;
    }
}
function Delete() {
    var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u8ba4\u8981\u5c06\u9009\u4e2d\u7684\u4fe1\u606f\u5220\u9664\u5417\uff1f')) {
        window.location.href = '?Action=Delete&ID='+l+'&Page=<%=Page%>&ClassID=<%=ClassID%>';
    }
}
</script>
</body>
</html>