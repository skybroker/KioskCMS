<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->
<%
Call ISPopedom(UserName,"Job_Peroson")
Action=ReplaceBadChar(Trim(Request("Action")))
NavParent=ReplaceBadChar(Trim(Request("NavParent")))
ID=ReplaceBadChar(Trim(Request("ID")))
If ID="" Then
	ID="0"
End If
Select Case Action
Case "Up"
	MinOrder=Conn.Execute("Select Min(NavOrder) From JobPersonInfo Where NavParent="&NavParent&"")(0)
	If MinOrder="" Then MinOrder="1"
	NavOrder=ReplaceBadChar(Trim(Request("NavOrder")))
	If Cstr(NavOrder)=Cstr(MinOrder) Then
		Response.Write("<script>alert('\u8be5\u7c7b\u522b\u5df2\u7ecf\u5904\u4e8e\u6700\u5934\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update JobPersonInfo Set NavOrder="&NavOrder&" Where NavOrder="&NavOrder-1&" And Navparent="&Navparent&"")
		Conn.Execute("Update JobPersonInfo Set NavOrder="&NavOrder-1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "Down"
	MaxOrder=Conn.Execute("Select Max(NavOrder) From JobPersonInfo Where NavParent="&NavParent&"")(0)
	NavOrder=ReplaceBadChar(Trim(Request("NavOrder")))
	If Cstr(NavOrder)=Cstr(MaxOrder) Then
		Response.Write("<script>alert('\u8be5\u7c7b\u522b\u5df2\u7ecf\u5904\u4e8e\u6700\u5e95\u90e8\u4e86\uff0c\u65e0\u6cd5\u8fdb\u884c\u987a\u5e8f\u7684\u8c03\u6574\u3002');history.back();</script>")
	Else
		Conn.Execute("Update JobPersonInfo Set NavOrder="&NavOrder&" Where NavOrder="&NavOrder+1&" And NavParent="&Navparent&"")
		Conn.Execute("Update JobPersonInfo Set NavOrder="&NavOrder+1&" Where ID="&ID&"")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End If
	Conn.Close
	Set Conn=Nothing
	Response.End()
Case "UnPublic"
	Conn.Execute("Update JobPersonInfo Set NavLock=1 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "InPublic"
	Conn.Execute("Update JobPersonInfo Set NavLock=0 Where ID="&ID&"")
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Response.End()
Case "Delete"
	Page=ReplaceBadChar(Trim(Request("Page")))
	ParentID=ReplaceBadChar(Trim(Request("ParentID")))
	AryID = Split(ID,",")
	For i = LBound(AryID) To UBound(AryID)
		If IsNumeric(AryID(i))=True Then
			Conn.Execute("Delete From JobPersonInfo Where ID In ("&AryID(i)&GetAllChild("ShopClass",AryID(i))&")")
		End If
	Next
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u5220\u9664\u64cd\u4f5c\u6210\u529f\uff0c\u786e\u5b9a\u540e\u8fd4\u56de\u7c7b\u522b\u5217\u8868\u9875\u9762\u3002');window.location.href='?Page="&Page&"&ID="&ParentID&"';</script>")
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
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：<a href="Job_Person.asp">人才列表->应聘人员详细信息</a></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下为所有信息分类的信息列表。<br>
注意：您可以进行添加、修改、删除等操作，保存之后立即生效。</td>
</tr>
</table>
</td>
</tr>
<tr>
<td colspan="2" valign="top">
<% 
ida=GetSafeInt1(request("id"))
page=GetSafeInt(request("page"))
set rs=conn.execute("select * from JobPersonInfo where id="&ida)
if rs.eof and rs.bof then
	rs.close:set rs=nothing
	AlertMsg "找不到对应记录!","2"
	Response.end()
else
	conn.execute("update JobPersonInfo set cZt='1' where id="&ida)
	Person=rs("Person")
%>
<form id="form1" name="form1" method="post">
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="Form" id="GridView1">
      <tr>
        <td height="23" colspan="4" class="topbar">&nbsp;应聘人员详细信息：</td>
        </tr>
      <tr>
        <td width="15%" height="25" align="right" class="tab1_td">姓 名：</td>
        <td width="27%" height="25" class="tab1_td"><%= rs("xinmin") %></td>
        <td width="16%" height="25" align="right" class="tab1_td">姓 别：</td>
        <td width="42%" height="25" class="tab1_td"><%= rs("xinbie") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">应聘职位：</td>
        <td height="25" class="tab1_td"><%= rs("gangwei") %></td>
        <td height="25" align="right" class="tab1_td">出生年月：</td>
        <td height="25" class="tab1_td"><%= rs("birthday") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">民族：</td>
        <td height="25" class="tab1_td"><%= rs("minzu") %></td>
        <td height="25" align="right" class="tab1_td">婚姻情况：</td>
        <td height="25" class="tab1_td"><%= rs("hunyin") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">身高：</td>
        <td height="25" class="tab1_td"><%= rs("shengao") %></td>
        <td height="25" align="right" class="tab1_td">政治面貌：</td>
        <td height="25" class="tab1_td"><%= rs("polital") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">出生日期：</td>
        <td height="25" class="tab1_td"><%= rs("birthday") %></td>
        <td height="25" align="right" class="tab1_td">文化程度：</td>
        <td height="25" class="tab1_td"><%= rs("wenhua") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">籍贯：</td>
        <td height="25" class="tab1_td"><%= rs("jiguan") %></td>
        <td height="25" align="right" class="tab1_td">毕业时间：</td>
        <td height="25" class="tab1_td"><%= rs("biyetime") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">所学专业：</td>
        <td height="25" class="tab1_td"><%= rs("zhuanye") %></td>
        <td height="25" align="right" class="tab1_td">毕业院校：</td>
        <td height="25" class="tab1_td"><%= rs("yuanxiao") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">联系电话：</td>
        <td height="25" class="tab1_td"><font color="#AA3E00"><%= rs("dianhua") %></font></td>
        <td height="25" align="right" class="tab1_td">电子邮件：</td>
        <td height="25" class="tab1_td"><a href="mailto:<%= rs("mail") %>"><%= rs("mail") %></a></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">联系地址：</td>
        <td height="25" class="tab1_td"><%= rs("dizhi") %></td>
        <td height="25" align="right" class="tab1_td">邮政编码：</td>
        <td height="25" class="tab1_td"><%= rs("Post") %></td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">希望待遇：</td>
        <td height="25" class="tab1_td"><%= rs("Money") %></td>
        <td height="25" align="right" class="tab1_td">&nbsp;</td>
        <td height="25" class="tab1_td">&nbsp;</td>
      </tr>
      <tr>
        <td height="25" align="right" class="tab1_td">个人简历：</td>
        <td height="25" colspan="3" class="tab1_td"><%= Ubb(Person) %></td>
      </tr>
      <tr>
        <td height="35" align="right" class="tab1_td">&nbsp;</td>
        <td height="35" colspan="3" class="tab1_td"><input type="button" value="返 回" class="Button" onClick="window.location.href='Job_Person.asp'"></td>
      </tr>
    </table>
</form>
    <% 
end if
rs.close:set rs=nothing
%>
</td>
</tr>
</table>
<script language="javascript" type="text/javascript">
function ChangeParent()
{
	var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u5b9a\u8981\u66f4\u6539\u9009\u4e2d\u7684\u7c7b\u522b\u7684\u6240\u5c5e\u7236\u7c7b\u522b\u5417\uff1f')) {
        window.location.href = 'News_ClassParent.asp?ID='+l;
    }
}
function Delete() {
    var l = GetAllChecked();
    if (l == "") {
        alert("\u4f60\u8fd8\u6ca1\u6709\u9009\u62e9\u8981\u64cd\u4f5c\u7684\u8bb0\u5f55\uff01");
        return;
    }
    if (confirm('\u786e\u8ba4\u8981\u5c06\u9009\u4e2d\u7684\u7c7b\u522b\u5220\u9664\u5417\uff1f')) {
        window.location.href = '?Action=Delete&ID='+l+'&Page=<%=Page%>&ParentID=<%=ID%>';
    }
}
</script>
</body>
</html>