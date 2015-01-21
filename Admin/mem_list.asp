<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#include File="../Include/Class_MD5.asp"-->

<%
Call ISPopedom(UserName,"mem_list")
%>
<%
page=GetSafeInt(request("page"))
 
if request("act")="yes" then
	id=trim(request("che"))
	if id<>"" then
		rid=Split(id, ",")
		if request("del")<>"" then
			for i=0 to UBound(rid)
				ida=GetSafeInt1(rid(i))
				conn.execute("delete from t_Member where id="&ida)
			next
		elseif request("op")<>"" then
			conn.execute("update t_Member set cZt='1' where id in ("&id&")")
		elseif request("lk")<>"" then
			conn.execute("update t_Member set cZt='0' where id in ("&id&")")
		end if
		AlertMsg "操作成功!","mem_list.asp?page="&page
	end if
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
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:70%">当前位置：<a href="mem_list.asp"><span class="MainMenu">会员中心</span></a></td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:30%; text-align:right"></td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下为会员信息内容的列表。<br>
注意：您可以进行搜索、删除等操作，保存确定之后立即生效。</td>
</tr>
</table>
</td>
</tr>
<tr>
  <td colspan="2">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23">
   
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="Form">
<form name="form2" method="post" action="mem_list.asp">
<tr>
<td width="55%" height="30">

      &nbsp;会员名或姓名:
        <input name="seltxt" type="text" id="seltxt" size="35">
        <input type="submit" name="Submit2" value="搜 索">

</td>
</tr>
</form>
</table>    
      
    </td>
  </tr>
</table>
  </td>
</tr>
<tr>
<td colspan="2" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top"><form name="form1" method="post" action="mem_list.asp?act=yes&page=<%= page %>">
      <div class="mydiv">
      <table width="100%" border="0" cellpadding="0" cellspacing="1" class="Form">
        <tr>
          <th width="6%" align="center" class="topbar">选择</th>
          <th width="10%" height="22" align="center" class="topbar">会员用户名</th>
          <th width="11%" align="center" class="topbar">注册时间</th>
		  <th width="24%" align="center" class="topbar">会员说明</th>
          </tr>
        <% 
   seltxt=GetSafeStr(request("seltxt"))
   if seltxt<>"" then str=" and (cMuser like '%"&seltxt&"%' or cName like '%"&seltxt&"%')"
  set rs=server.createobject("adodb.recordset")
  sql="select * from t_Member where id>0"&str&" order by id desc"
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then
  Response.Write("<tr><td height='22' colspan='7' class='tab1_td'>·暂无信息！</td></tr>")
  else
  rs.pagesize=25
  pagecount1=rs.pagecount
  if page>pagecount1 then page=pagecount1
  rs.absolutepage=page
  for i=1 to rs.pagesize
  if rs.eof then exit for
  zt_a=rs("cZt")
   %>
        <tr class="tab1_td" onMouseOver="this.className='tab1_td1'" onMouseOut="this.className='tab1_td'">
          <td><input name="che" type="checkbox" id="che" value="<%= rs("id") %>"></td>
          <td height="22">&nbsp;<%= rs("cMuser") %></td>
          <td><%= FormatDateTime(rs("tDate"),2)%></td>
          <td height="22">
		  <%
		  if rs("cAbout")<>"" then
		  Response.Write(rs("cAbout"))
		  else
		  Response.Write("普通会员")
		  end if
		  %>
		  </td>
          </tr>
        <% 
  rs.movenext
  next
  end if
  rs.close:set rs=nothing
  
  function showzta(num)
	if num="0" then
		zt_str="<font color='#FF0000'>已锁定</font>"
	elseif num="1" then
		zt_str="正常"
	end if	
	showzta=zt_str
end function
   %>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
		  <td height="3"></td>
		</tr>
	  </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="Form">
        <tr>
          <th width="46%" height="23" class="tab1_td">&nbsp;&nbsp;<input name="checall" type="checkbox" id="checall" value="checkbox" onClick="cheall('che');">
          全选&nbsp;&nbsp;<!--<input name="lk" type="submit" id="lk" onClick="return delok('che')" value="锁定"> 
          <input name="op" type="submit" id="op" value="开通" onClick="return delok('che')">-->
          <input name="del" type="submit" id="del" value="删除" onClick="return delok('che')"></th>
          <th width="54%" align="right" class="tab1_td"><% call pager(page,pagecount1,"&sh="&sh&"&seltxt="&server.URLEncode(seltxt)) %>&nbsp;</th>
        </tr>
      </table>
    </div></form></td>
  </tr>
</table>
</td>
</tr>
</table>
<script language="javascript">
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
var flag=0;
for(i=0;i<document.form1.elements.length;i++)
{
	var e=document.form1.elements[i];
	if (e.name==opj){
		if (e.checked==true){
			flag=1;
			if(!confirm("你确定要操作选中项吗?"))
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