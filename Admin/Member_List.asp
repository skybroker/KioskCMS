<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->
<%
Call ISPopedom(UserName,"Member_List")
Action=ReplaceBadChar(Trim(Request("Action")))
UserName=ReplaceBadChar(Trim(Request("UserName")))
ID=ReplaceBadChar(Trim(Request("ID")))
Page=ReplaceBadChar(Trim(Request("Page")))
UserLock=ReplaceBadChar(Trim(Request("UserLock")))
Select Case Action
Case "Reset"
	If IsNumeric(ID)=True Then
		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Sql="Select * From SiteUsers Where ID="&ID&""
		Rs.Open Sql,Conn,1,3
		Rs("UserPass")=Md5("123456",32)
		Rs.Update
		Rs.Close
		Set Rs=Nothing
	End If
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u606d\u559c\uff0c\u6210\u529f\u5c06\u9009\u4e2d\u7684\u4f1a\u5458\u8d26\u53f7\u5bc6\u7801\u91cd\u7f6e\u6210\u0031\u0032\u0033\u0034\u0035\u0036\u3002');window.location.href='Member_List.asp?Page="&Page&"';</script>")
	Response.End()
Case "IsLock"
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Update SiteUsers Set UserLock=1 Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u9501\u5b9a\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&UserLock="&UserLock&"';</script>")
	Response.End()
Case "UnLock"
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Update SiteUsers Set UserLock=0 Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Redirect("?Page="&Page&"&UserLock="&UserLock&"")
	Response.End()
Case "Delete"
	Page=ReplaceBadChar(Trim(Request("Page")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Delete From SiteUsers Where ID="&AryID(i)&"")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u5220\u9664\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&UserLock="&UserLock&"';</script>")
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
	$("#ClassID").val("<%=UserLock%>");
	$("#ClassID").change(function(){
		window.location.href='?UserLock='+$("#ClassID").val()+'&UserName=<%=Server.URLEncode(UserName)%>';
	});
});
</script>
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：<a href="Member_List.asp">会员列表</a></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">
<select id="ClassID" name="ClassID" style="width:150px;">
<option value="">|--所有会员列表</option>
<option value="0">|--有效会员列表</option>
<option value="1">|--锁定会员列表</option>
</select>&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下为所有会员的列表，您可以进行添加、修改、删除等操作。<br />
注意：如果不想让该会员登录，请使用锁定功能。</td>
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
<th width="15%" class="Right">登录账号<a href="#" onclick="top.CreateNewTab('Member_Add.asp','Member_Add','添加会员')">[+]</a></th>
<th width="15%" class="Right">真实姓名</th>
<th width="30%" class="Right">公司名称</th>
<th width="20%" class="Right">联系电话</th>
<th width="15%">管理操作</th>
</tr>
<%
If UserLock<>"" Then
	StrSql=" And UserLock="&UserLock&" "
End If
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From SiteUsers Where 1=1 "&StrSql&" Order By ID Desc"
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
<td class="Right"><a href="Member_Edit.asp?ID=<%=Rs("ID")%>&Page=<%=Page%>"><%=Rs("UserName")%></a></td>
<td class="Right"><%=Rs("TrueName")%>&nbsp;</td>
<td class="Right"><%=Rs("CompanyName")%>&nbsp;</td>
<td class="Right"><%=Rs("TelPhone")%>&nbsp;</td>
<td><a href="?Action=Reset&ID=<%=Rs("ID")%>&Page=<%=Page%>" onclick="if(!confirm('\u786e\u8ba4\u8981\u5bf9\u8be5\u7ba1\u7406\u8d26\u53f7\u8fdb\u884c\u5bc6\u7801\u91cd\u7f6e\u7684\u64cd\u4f5c\u5417\uff1f\u9ed8\u8ba4\u5bc6\u7801\uff1a\u0031\u0032\u0033\u0034\u0035\u0036')) return false;">重置</a>┊<%If Rs("UserLock")="0" Then
Response.Write("<a href=""?Action=IsLock&ID="&Rs("ID")&"&Page="&Page&"&UserLock="&UserLock&""">锁定</a>")
Else
Response.Write("<a href=""?Action=UnLock&ID="&Rs("ID")&"&Page="&Page&"&UserLock="&UserLock&""">解锁</a>")
End If
%>
</td>
</tr>
<%
i=i+1
Rs.MoveNext
Loop
%>
<tr>
<th colspan="2" style="font-weight:normal">操作：<a href="javascript:IsLock();" style="font-weight:normal">锁定</a>&nbsp;┊&nbsp;<a href="javascript:Delete();" style="font-weight:normal">删除</a></th>
<th colspan="3" style="font-weight:normal;text-align:right">共<%=Rs.PageCount%>页&nbsp;第<%=Page%>页&nbsp;<%=PageSize%>条/页&nbsp;共<%=Total%>条&nbsp;
<%if Page=1 then%>
首 页&nbsp;上一页&nbsp;
<%Else%>
<a href="?Page=1&UserLock=<%=UserLock%>">首 页</a>&nbsp;<a href="?Page=<%=Page-1%>&UserLock=<%=UserLock%>">上一页</a>&nbsp;
<%End If%>
<%If Rs.PageCount-Page<1 Then%>下一页&nbsp;尾 页&nbsp;
<%Else%><a href="?Page=<%=Page+1%>&UserLock=<%=UserLock%>">下一页</a>&nbsp;<a href="?Page=<%=Rs.PageCount%>&UserLock=<%=UserLock%>">尾 页</a>&nbsp;
<%End If%>
</th>
<th>
<select style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体;width:90%;" onChange="location=this.options[this.selectedIndex].value" name="Menu_1"> 
<%For Pagei=1 To Rs.PageCount%>
<%if Cint(Pagei)=Cint(Page) Then%>
<option value="?Page=<%=Pagei%>&UserLock=<%=UserLock%>" selected="selected">第<%=Pagei%>页</option>
<%Else%>
<option value="?Page=<%=Pagei%>&UserLock=<%=UserLock%>">第<%=Pagei%>页</option>
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
	if (confirm('\u786e\u5b9a\u8981\u5c06\u9009\u4e2d\u7684\u4f1a\u5458\u8fdb\u884c\u9501\u5b9a\u64cd\u4f5c\u5417\uff1f')) {
        window.location.href = '?Action=IsLock&ID='+l+'&Page=<%=Page%>&UserLock=<%=UserLock%>';
    }
}
function Delete() {
    var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u8ba4\u8981\u5c06\u9009\u4e2d\u7684\u4f1a\u5458\u5220\u9664\u5417\uff1f')) {
        window.location.href = '?Action=Delete&ID='+l+'&Page=<%=Page%>&UserLock=<%=UserLock%>';
    }
}
</script>
</body>
</html>