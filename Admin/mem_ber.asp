<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config2.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->

<%
Call ISPopedom(UserName,"mem_ber")
%>
<%
if request("act")="edit" then
	id=GetSafeInt1(request("id"))
	set rs=conn.execute("select cMuser,cMpass,cAbout from t_Member where id="&id)
	if rs.eof and rs.bof then
		rs.close:set rs=nothing
		AlertMsg "没有对应记录！","2"
		response.End()
	end if
	cMuser1=rs("cMuser")
	cMpass1=rs("cMpass")
	cAbout1=rs("cAbout")
	rs.close:set rs=nothing
	tjstr="?act1=saveedit&id="&id
else
	tjstr="?act1=saveadd"
end if

if request("act1")<>"" then
	cMuser=GetSafeStr(request.form("cMuser"))
	cMpass=GetSafeStr(request.form("cMpass"))
	cAbout=GetSafeStr(request.form("cAbout"))
end if
if request("act1")="saveadd" then
	if ctuser(cMuser)>0 then
		AlertMsg "此会员名已存在,请填写其它名称!","2"
		response.End()
	end if
	conn.execute("insert into t_Member (cMuser,cMpass,cAbout) values ('"&cMuser&"','"&cMpass&"','"&cAbout&"')")
	AlertMsg "会员添加成功!","mem_ber.asp"
	response.End()
elseif request("act1")="saveedit" then
	cuser=GetSafeStr(request.form("cuser"))
	if cMuser<>cuser then
		if ctuser(cMuser)>0 then
			AlertMsg "此会员名已存在,请填写其它名称!","2"
			response.End()
		end if
	end if
	id=GetSafeInt1(request("id"))
	set rsa=server.CreateObject("adodb.recordset")
	rsa.open "select cMuser,cMpass,cAbout from t_Member where id="&id,conn,1,3
	if rsa.eof and rsa.bof then
		rsa.close:set rsa=nothing
		AlertMsg "没有对应记录！","2"
		response.End()
	end if
	rsa("cMuser")=cMuser
	rsa("cMpass")=cMpass
	rsa("cAbout")=cAbout
	rsa.update
	rsa.close:set rsa=nothing
	AlertMsg "会员修改成功!","mem_ber.asp?page="&request.Cookies("m_pg")
	response.End()
end if

function ctuser(user)
	set rsa=conn.execute("select count(id) from t_Member where cMuser='"&user&"'")
	ct=rsa(0)
	rsa.close:set rsa=nothing
	ctuser=ct
end function

id_a=trim(request("che"))
if id_a<>"" then
	rid=Split(id_a, ",")
	for i=0 to UBound(rid)
		idf=GetSafeInt1(rid(i))
		set rs=conn.execute("select count(id) from t_Member where id="&idf)
		ct1=rs(0)
		rs.close:set rs=nothing
		if ct1=0 then
			AlertMsg "参数不正确!","2"
			Response.end()
		end if
	next
	if request("subok_1")<>"" then
		conn.execute("update t_Member set cZt='1' where id in ("&idf&")")
	end if
	if request("subok_2")<>"" then
		conn.execute("update t_Member set cZt='0' where id in ("&idf&")")
	end if
	if request("subok_3")<>"" then
		conn.execute("delete from t_Member where id in ("&idf&")")
	end if	
	AlertMsg "操作成功!","mem_ber.asp?page="&request.Cookies("m_pg")
end if
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
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:70%">当前位置：<a href="mem_ber.asp"><span class="MainMenu">会员中心</span></a></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:30%; text-align:right"><a href="mem_ber.asp"><span class="MainMenu">添加会员</span></a></td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下为会员信息内容列表，默认显示所有会员信息。<br>
注意：您可以进行开通选中、锁定选中、删除选中等操作，点击确认之后立即生效。</td>
</tr>
</table>
</td>
</tr>
<tr>
  <td colspan="2">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23">
    <div class="mydiv">
  <table width="100%" border="0" cellpadding="3" cellspacing="1" class="Form">
  <form name="form2" method="post" action="<%= tjstr %>" onSubmit="return confirm('你确定要提交结果吗?')">
    <tr>
      <td height="23" colspan="5" class="topbar">&nbsp;会员管理(注:点击列表中的用户名,此框的内容即为修改状态,按'提交'即可保存修改.)</td>
    </tr>
    <tr>
      <td width="13%" height="23" align="right" class="tab1_td">会员用户名：</td>
      <td width="24%" height="23" class="tab1_td"><input name="cMuser" type="text" id="cMuser" value="<%= cMuser1 %>" size="20" maxlength="30"></td>
      <td width="12%" height="23" align="right" class="tab1_td">会员密码：</td>
      <td width="21%" height="23" class="tab1_td"><input name="cMpass" type="text" id="cMpass" value="<%= cMpass1 %>" size="20" maxlength="30"></td>
      <td width="30%" height="23" class="tab1_td"><input type="submit" name="Submit" value="提 交" onClick="return Checks();">
	  </td>
    </tr>
    <tr>
      <td height="23" align="right" class="tab1_td">会员说明：</td>
      <td height="23" colspan="4" class="tab1_td"><input name="cAbout" type="text" id="cAbout" value="<%= cAbout1 %>" size="45" maxlength="65">
      此处是方便管理员对会员的辩别之用。
        <input name="cuser" type="hidden" id="cuser" value="<%= cMuser1 %>"></td>
    </tr>
	</form>
  </table>
</div>
    </td>
  </tr>
</table>
  </td>
</tr>
<tr>
<td colspan="2" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top"><form name="form1" method="post" action="mem_ber.asp">
<div class="mydiv">
  <table width="100%" border="0" cellpadding="0" cellspacing="1" class="Form">
    <tr>
      <th width="5%" align="center" class="topbar">选择</th>
      <th width="14%" height="23" align="center" class="topbar">会员用户名</th>
      <th width="16%" height="23" align="center" class="topbar">原始<span class="tabtop">密码</span></th>
      <th width="8%" height="23" align="center" class="topbar"><span class="tabtop">会员状态</span></th>
      <th width="31%" height="23" align="center" class="topbar"><span class="tabtop">会员说明</span></th>
      <th width="13%" height="23" align="center" class="topbar"><span class="tabtop">添加时间</span></th>
      <th width="13%" align="center" class="topbar">操作</th>
    </tr>
<% 
page=GetSafeInt(request("page"))
set rs=server.createobject("adodb.recordset")
sql="select * from t_Member order by id desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
Response.Write("<tr><td height='22' colspan='6' class='tab1_td'>·暂无会员帐号！</td></tr>")
else
rs.pagesize=20
pagecount1=rs.pagecount
if page>pagecount1 then page=pagecount1
rs.absolutepage=page
for i=1 to rs.pagesize
if rs.eof then exit for
if rs("cZt")="1" then
	stra="正常"
else
	stra="<font color='FF0000'>锁定</font>"
end if
%>	
    <tr class="tab1_td" onMouseOver="this.className='tab1_td1'" onMouseOut="this.className='tab1_td'">
      <td align="center"><input name="che" type="checkbox" id="che" value="<%= rs("id") %>"></td>
      <td height="23"><a href="?id=<%= rs("id") %>&act=edit"><%= rs("cMuser") %></a></td>
      <td height="23"><%= rs("cMpass") %></td>
      <td height="23"><%= stra %></td>
      <td height="23">&nbsp;<%= rs("cAbout") %></td>
      <td height="23"><%= FormatDateTime(rs("tDate"),1)%></td>
      <td><a href="?id=<%= rs("id") %>&act=edit">修改</a></td>
    </tr>
<% 
rs.movenext
next
end if
rs.close:set rs=nothing
response.Cookies("m_pg")=page
%>	
  </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
  <td height="3"></td>
</tr>
</table>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="Form">
    <tr>
      <th width="50%" height="25" align="left" class="tab1_td">
	  &nbsp;
        <input name="checall" type="checkbox" id="checall" onClick="cheall('che');">
      全选&nbsp;&nbsp;<input name="subok_1" type="submit" id="subok_1" onClick="return delok('che')" value="开通选中">
      &nbsp;
      <input name="subok_2" type="submit" id="subok_2" onClick="return delok('che')" value="锁定选中">
      &nbsp;
      <input name="subok_3" type="submit" id="subok_3" onClick="return delok('che')" value="删除选中">
	  <input name="chid" type="hidden" id="chid" value="yes">	  </th>
      <th width="50%" height="25" align="right" class="tab1_td">
	    <%
        if page>1 then
        	response.write "<a href=?page=1>〖首页〗</a>&nbsp;<a href=?page="&(page-1)&">〖上页〗</a>&nbsp;"
        else
        	response.write "〖首页〗&nbsp;〖上页〗&nbsp;"
        end if
        if page < pagecount1 then
        	response.write "<a href=?page="&(page+1)&">〖下页〗</a>&nbsp;<a href=?page="&pagecount1&">〖末页〗</a>"
        else
       	 response.write "〖下页〗&nbsp;〖末页〗"
        end if
        %>&nbsp;&nbsp; </th>
    </tr>
  </table> 
</div>
 </form></td>
  </tr>
</table>
</td>
</tr>
</table>
<script Language="javascript">
function Checks()
{
	if (document.form2.cMuser.value=="")
	{
      alert("会员用户名不能为空!");
	  document.form2.cMuser.focus();
	  return false;
     }
	if (document.form2.cMpass.value=="")
	{
      alert("会员密码不能为空!");
	  document.form2.cMpass.focus();
	  return false;
     }
}
function cheall(opj){
for(i=0;i<document.form1.elements.length;i++){
	var e=document.form1.elements[i];
	if (e.name==opj){
		if (document.form1.checall.checked)
			e.checked=true;		
		else
			e.checked=false;
	}
  }
}

function delok(opj){
var flag=0
for(i=0;i<document.form1.elements.length;i++)
{
	var e=document.form1.elements[i];
	if (e.name==opj){
		if (e.checked==true){
			flag=1;
			if(!confirm("确定执行此项操作吗?"))
				return false;
			break;
		}
	}
}
if(flag==0){
	alert('请选择至少一条以上的记录!');
	return false;
}
}
</script>
</body>
</html>