<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config2.asp"-->
<!--#include File="../Include/Class_Function.asp"-->
<!--#Include File="../Editor/fckeditor.asp"-->
<%
Call ISPopedom(UserName,"Job_ContentAdd")
Action=ReplaceBadChar(Trim(Request("Action")))
ClassID=ReplaceBadChar(Trim(Request("ClassID")))
If ClassID="" Then
	ClassID="0"
End If
If Action="Save" Then
	Jobs=ReplaceBadChar(Trim(Request("Jobs")))
	EnJobs=ReplaceBadChar(Trim(Request("EnJobs")))
	ClassID=ReplaceBadChar(Trim(Request("ClassID")))
	TitleRequest=Trim(Request("TitleRequest"))
	EnTitleRequest=Trim(Request("EnTitleRequest"))
	JobNumber=Request("JobNumber")
	RDepart=Trim(Request("RDepart"))
	EnRDepart=Trim(Request("EnRDepart"))
	Gender=ReplaceBadChar(Trim(Request("Gender")))
	EnGender=ReplaceBadChar(Trim(Request("EnGender")))
	Experience=Trim(Request("Experience"))
	EnExperience=Trim(Request("EnExperience"))
	Education=ReplaceBadChar(Trim(Request("Education")))
	EnEducation=ReplaceBadChar(Trim(Request("EnEducation")))
	Age=Trim(Request("Age"))
	Profession=ReplaceBadChar(Trim(Request("Profession")))
	EnProfession=ReplaceBadChar(Trim(Request("EnProfession")))
	WorkAreas=ReplaceBadChar(Trim(Request("WorkAreas")))
	EnWorkAreas=ReplaceBadChar(Trim(Request("EnWorkAreas")))
	EffectiveLimit=ReplaceBadChar(Trim(Request("EffectiveLimit")))
	EnEffectiveLimit=ReplaceBadChar(Trim(Request("EnEffectiveLimit")))
	ContactName=ReplaceBadChar(Trim(Request("ContactName")))
	Phone=ReplaceBadChar(Trim(Request("Phone")))
	JobContentAdd=ReplaceBadChar(Trim(Request("JobContentAdd")))
	Fax=Trim(Request("Fax"))
	Email=Trim(Request("Email"))
	Address=ReplaceBadChar(Trim(Request("Address")))
	EnAddress=Trim(Request("EnAddress"))
	RAT=Trim(Request("RAT"))
	EnRAT=ReplaceBadChar(Trim(Request("EnRAT")))
	PostTime=Trim(Request("PostTime"))
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From JobInfo where ID="&session("ID")
	Rs.Open Sql,Conn,1,3
	Rs("Jobs")=Jobs
	Rs("EnJobs")=EnJobs
	Rs("ClassID")=ClassID
	Rs("TitleRequest")=TitleRequest
	Rs("EnTitleRequest")=EnTitleRequest
	Rs("JobNumber")=JobNumber
	Rs("RDepart")=RDepart
	Rs("EnRDepart")=EnRDepart
	Rs("Gender")=Gender
	Rs("EnGender")=EnGender
	Rs("Experience")=Experience
	Rs("EnExperience")=EnExperience
	Rs("Education")=Education
	Rs("EnEducation")=EnEducation
	Rs("Age")=Age
	Rs("Profession")=Profession
	Rs("EnProfession")=EnProfession
	Rs("WorkAreas")=WorkAreas
	Rs("EnWorkAreas")=EnWorkAreas
	Rs("EffectiveLimit")=EffectiveLimit
	Rs("EnEffectiveLimit")=EnEffectiveLimit
	Rs("ContactName")=ContactName
	Rs("Phone")=Phone
	Rs("JobContentAdd")=JobContentAdd
	Rs("Fax")=Fax
	Rs("Email")=Email
	Rs("Address")=Address
	Rs("EnAddress")=EnAddress
	Rs("RAT")=RAT
	Rs("EnRAT")=EnRAT
	Rs("PostTime")=Now()
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	session("ID")=""
	Response.Write("<script>alert('\u4fe1\u606f\u5185\u5bb9\u6dfb\u52a0\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='Job_List.asp';</script>")
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
<script type="text/javascript">
$(document).ready(function(){
	$("#ClassID").val("<%=ClassID%>");
});
</script>
</head>
<body>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：招聘管理 >> 职位修改</td>
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:20%; text-align:right">&nbsp;</td>
</tr>
<tr>
<td height="80" colspan="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="60"><img src="Images/SiteInfo.jpg" width="60" height="61"></td>
<td width="90%" valign="top">以下所有项目带*均不能为空，请准确真实的填写相关信息。<br />
注意：您可以进行添加、修改、删除等操作，保存之后立即生效。<br />
  </td>
</tr>
</table></td>
</tr>
<tr>
<td colspan="2" valign="top">
<%
	session("ID")=Request("ID")
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From JobInfo where ID="&session("ID")
	Rs.Open Sql,Conn,1,3
	Jobs=Rs("Jobs")
	EnJobs=Rs("EnJobs")
	ClassID=Rs("ClassID")
	TitleRequest=Rs("TitleRequest")
	EnTitleRequest=Rs("EnTitleRequest")
	JobNumber=Rs("JobNumber")
	RDepart=Rs("RDepart")
	EnRDepart=Rs("EnRDepart")
	Gender=Rs("Gender")
	EnGender=Rs("EnGender")
	Experience=Rs("Experience")
	EnExperience=Rs("EnExperience")
	Education=Rs("Education")
	EnEducation=Rs("EnEducation")
	Age=Rs("Age")
	Profession=Rs("Profession")
	EnProfession=Rs("EnProfession")
	WorkAreas=Rs("WorkAreas")
	EnWorkAreas=Rs("EnWorkAreas")
	EffectiveLimit=Rs("EffectiveLimit")
	EnEffectiveLimit=Rs("EnEffectiveLimit")
	ContactName=Rs("ContactName")
	Phone=Rs("Phone")
	JobContentAdd=Rs("JobContentAdd")
	Fax=Rs("Fax")
	Email=Rs("Email")
	Address=Rs("Address")
	EnAddress=Rs("EnAddress")
	RAT=Rs("RAT")
	EnRAT=Rs("EnRAT")
	Rs.Close
	Set Rs=Nothing
%>
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#Jobs").val()=="")
	{
		alert("\u4fe1\u606f\u6807\u9898\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#Jobs").focus();
		return false;
	}
	return true;	
}
</script>
<form id="form1" name="form1" method="post" action="?Action=Save" onSubmit="return CheckForm();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
  <th colspan="4">职位修改</th>
</tr>
<tr>
<td width="11%" align="right" class="Right">职位名称：</td>
<td width="35%" class="Right">
  <div class="float_left_210txt">
	<input type="text" id="Jobs" name="Jobs" value="<%=Jobs%>" class="Input200px"/>*（必填）
  </div>
</td>
<td width="11%" align="right" class="Right">招聘类别：</td>
<td width="43%"><select id="ClassID" name="ClassID" style="width:200px;">
<option value="<%=ClassID%>">|--请选择栏目</option>
<%=GetSelect("JobClass",0)%>
</select></td>
</tr>

<tr>
<td class="Right" align="right">职称要求：</td>
<td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="TitleRequest" name="TitleRequest" value="<%=TitleRequest%>" class="Input200px"/>*（必填）
  </div>

</td>
<td class="Right" align="right">招聘人数：</td>
<td><span class="Right">
  <input type="text" id="JobNumber" name="JobNumber" value="<%=JobNumber%>" onKeyUp="value=value.replace(/[^0123456789-]/g,'')" class="Input200px"/> 
  人
</span></td>
</tr>
<tr>
<td class="Right" align="right">招聘部门：</td>
<td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="RDepart" name="RDepart" value="<%=RDepart%>" class="Input200px"/>*（必填）
  </div>

</td>
<td align="right" class="Right">性别要求：</td>
<td><select id="Gender" name="Gender" style="width:200px;">
  <option value="0">男</option>
  <option value="1">女</option>
  <option value="2">不限</option>
</select></td>
</tr>
<tr>
  <td class="Right" align="right">工作经验：</td>
  <td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="Experience" name="Experience" value="<%=Experience%>" class="Input200px"/>
  </div>
  </td>
  <td align="right" class="Right">学历要求：</td>
  <td>

  <div class="float_left_210txt">
	<input type="text" id="Education" name="Education" value="<%=Education%>" class="Input200px"/>
  </div>
  </td>
</tr>
<tr>
  <td class="Right" align="right"> 年龄要求：</td>
  <td class="Right"><input type="text" id="Age" name="Age" value="<%=Age%>" class="Input200px"/></td>
  <td align="right" class="Right">专业要求：</td>
  <td>
  <div class="float_left_210txt">
	<input type="text" id="Profession" name="Profession" value="<%=Profession%>" class="Input200px"/>
  </div>
  </td>
</tr>
<tr>
  <td class="Right" align="right">工作地区：</td>
  <td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="WorkAreas" name="WorkAreas" value="<%=WorkAreas%>" class="Input200px"/>
  </div>

  </td>
  <td align="right" class="Right">有效期：</td>
  <td><input type="text" id="EffectiveLimit" name="EffectiveLimit" value="<%=EffectiveLimit%>" class="Input200px"/></td>
</tr>
<tr>
  <td class="Right" align="right">要求与待遇：</td>
  <td colspan="3" class="Right">
<textarea name="JobContentAdd" cols="80" rows="10" id="JobContentAdd" style="padding:5px;"><%= JobContentAdd %></textarea>
  </td>
</tr>
  
<th colspan="4">职位应聘联系方式</th>
<tr>
  <td class="Right" align="right">联系人姓名：</td>
  <td><span class="Right">
    <input type="text" id="ContactName" name="ContactName" value="<%=ContactName%>" class="Input200px"/>
  </span></td>
  <td class="Right" align="right">联系电话：</td>
  <td class="Right" align="left"><input type="text" id="Phone" name="Phone" value="<%=Phone%>" class="Input200px"/></td>
<tr>
  <td class="Right" align="right">单位传真：</td>
  <td class="Right" align="left"><input type="text" id="Fax" name="Fax" value="<%=Fax%>" class="Input200px"/></td>
  <td class="Right" align="right">电子邮箱：</td>
  <td class="Right" align="left"><input type="text" id="Email" name="Email" value="<%=Email%>" class="Input200px"/></td>
<tr>
  <td class="Right" align="right">通信地址：</td>
  <td class="Right" align="left">
  <div class="float_left_210txt">
	<input type="text" id="Address" name="Address" value="<%=Address%>" class="Input200px"/>
  </div>

  </td>
  <td class="Right" align="right">&nbsp;</td>
  <td class="Right" align="right">&nbsp;</td>
<tr style="display:none;">
<td class="Right" align="right">浏览人群：</td>
<td class="Right"><input type="radio" id="JobVisit" name="JobVisit" value="0" checked="checked"/>所有人群<input type="radio" id="JobVisit" name="JobVisit" value="1"/>网站会员</td>
<td class="Right" align="right">信息状态：</td>
<td><input type="radio" id="JobLock" name="JobLock" value="0" checked="checked"/>解锁状态<input type="radio" id="JobLock" name="JobLock" value="1"/>锁定状态</td>
</tr>
<tr>
<td class="Right" align="right">&nbsp;</td>
<td colspan="3"><input type="submit" value="保 存" class="Button"> <input name="button" type="button" class="Button" onclick="window.location.href='Job_List.asp'" value="返 回" /></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>