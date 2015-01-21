<!--#include File="CheckLogin.asp"-->
<!--#include File="../Config/Config.asp"-->
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
	JobContentAdd=Trim(Request("JobContentAdd"))
	Fax=Trim(Request("Fax"))
	Email=Trim(Request("Email"))
	Address=ReplaceBadChar(Trim(Request("Address")))
	EnAddress=Trim(Request("EnAddress"))
	JobLock=ReplaceBadChar(Trim(Request("JobLock")))	
	JobVisit=ReplaceBadChar(Trim(Request("JobVisit")))
	PostTime=Trim(Request("PostTime"))
	'///////////排序/////////////
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select JobOrder From JobInfo Order By JobOrder desc"
	Rs.Open Sql,Conn,1,1
	If Not (Rs.Eof Or Rs.Bof) Then
		if len(Rs("JobOrder"))>0 then
			JobOrder=Cstr(Trim(Rs("JobOrder")))+1
		else
			JobOrder=1
		end if
	Else
		JobOrder=1
	End If
	Rs.Close
	Set Rs=Nothing
	'///////////////////////
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * From JobInfo"
	Rs.Open Sql,Conn,1,3
	Rs.AddNew
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
	Rs("JobClick")=1
	Rs("JobOrder")=JobOrder
	Rs("JobLock")=JobLock
	Rs("JobVisit")=JobVisit
	Rs("JobIndex")=0
	Rs("PostTime")=Now()
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Conn.Close
	Set Conn=Nothing
	Response.Write("<script>alert('\u4fe1\u606f\u5185\u5bb9\u6dfb\u52a0\u64cd\u4f5c\u6210\u529f\u3002');window.location.href='Job_ContentAdd.asp';</script>")
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
<td style="border-bottom:solid 1px #dde4e9;height:30px;width:80%">当前位置：招聘管理 >> 添加职位</td>
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
<script language="javascript" type="text/javascript">
function CheckForm()
{
	if ($("#Jobs").val()=="")
	{
		alert("\u804c\u4f4d\u540d\u79f0\u4e0d\u80fd\u4e3a\u7a7a\u3002");
		$("#Jobs").focus();
		return false;
	}
	if ($("#ClassID").val()==0)
	{
		alert("\u8bf7\u9009\u62e9\u804c\u4f4d\u7c7b\u522b\u0021");
		$("#ClassID").focus();
		return false;
	}
	return true;	
}
</script>
<form id="form1" name="form1" method="post" action="?Action=Save" onSubmit="return CheckForm();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Form">
<tr>
  <th colspan="5">添加职位</th>
</tr>
<tr>
<td width="9%" align="right" class="Right">职位名称：</td>
<td width="34%" class="Right">

  <div class="float_left_210txt">
	<input type="text" id="Jobs" name="Jobs" value="" class="Input200px"/> *（必填）
  </div>
</td>
<td width="9%" align="right" class="Right">招聘类别：</td>
<td colspan="2"><select id="ClassID" name="ClassID" style="width:200px;">
<option value="0">|--请选择栏目</option>
<%=GetSelect("JobClass",0)%>
</select></td>
</tr>

<tr>
<td class="Right" align="right">职称要求：</td>
<td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="TitleRequest" name="TitleRequest" value="" class="Input200px"/> *（必填）
  </div>
</td>
<td class="Right" align="right">招聘人数：</td>
<td colspan="2"><span class="Right">
  <input type="text" id="JobNumber" name="JobNumber" value="" onKeyUp="value=value.replace(/[^0123456789-]/g,'')" class="Input200px"/>
  名
</span></td>
</tr>
<tr>
<td class="Right" align="right">招聘部门：</td>
<td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="RDepart" name="RDepart" value="" class="Input200px"/> *（必填）
  </div>
</td>
<td align="right" class="Right">性别要求：</td>
<td colspan="2">
<select id="Gender" name="Gender" style="width:200px;">
<option value="0">男</option>
<option value="1">女</option>
<option value="2">不限</option>
</select>
</td>
</tr>
<tr>
  <td class="Right" align="right">工作经验：</td>
  <td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="Experience" name="Experience" value="" class="Input200px"/>
  </div>
 </td>
  <td align="right" class="Right">学历要求：</td>
  <td width="34%">
  <div class="float_left_210txt">
	<input type="text" id="Education" name="Education" value="" class="Input200px"/>
  </div>
  </td>
  <td width="14%">&nbsp;</td>
</tr>
<tr>
  <td class="Right" align="right"> 年龄要求：</td>
  <td class="Right"><input type="text" id="Age" name="Age" value="" class="Input200px"/></td>
  <td align="right" class="Right">专业要求：</td>
  <td>
  <div class="float_left_210txt">
	<input type="text" id="Profession" name="Profession" value="" class="Input200px"/>
  </div>
  </td>
  <td>&nbsp;</td>
</tr>
<tr>
  <td class="Right" align="right">工作地区：</td>
  <td class="Right">
  <div class="float_left_210txt">
	<input type="text" id="WorkAreas" name="WorkAreas" value="" class="Input200px"/>
  </div>
  </td>
  <td align="right" class="Right">有效期：</td>
  <td colspan="2"><input type="text" id="EffectiveLimit" name="EffectiveLimit" value="" class="Input200px"/></td>
</tr>
<tr>
  <td class="Right" align="right">要求与待遇：</td>
  <td colspan="4" class="Right">
  <textarea name="JobContentAdd" cols="80" rows="10" id="JobContentAdd"></textarea>
  </td>
</tr>
  
<th colspan="5">职位应聘联系方式</th>
<tr>
  <td class="Right" align="right">联系人姓名：</td>
  <td class="Right">
    <input type="text" id="ContactName" name="ContactName" value="" class="Input200px"/>
  </td>
  <td class="Right" align="right">联系电话：</td>
  <td colspan="2" align="left" class="Right"><input type="text" id="Phone" name="Phone" value="" class="Input200px"/></td>
<tr>
  <td class="Right" align="right">单位传真：</td>
  <td class="Right" align="left"><input type="text" id="Fax" name="Fax" value="" class="Input200px"/></td>
  <td class="Right" align="right">电子邮箱：</td>
  <td colspan="2" align="left" class="Right"><input type="text" id="Email" name="Email" value="" class="Input200px"/></td>
<tr>
  <td class="Right" align="right">通信地址：</td>
  <td class="Right" align="left">

  <div class="float_left_210txt">
	<input type="text" id="Address" name="Address" value="" class="Input200px"/>
  </div>
  </td>
  <td class="Right" align="right">&nbsp;</td>
  <td colspan="2" align="right" class="Right">&nbsp;</td>
<tr style="display:none;">
<td class="Right" align="right">浏览人群：</td>
<td class="Right"><input type="radio" id="JobVisit" name="JobVisit" value="0" checked="checked"/>所有人群<input type="radio" id="JobVisit" name="JobVisit" value="1"/>网站会员</td>
<td class="Right" align="right">信息状态：</td>
<td colspan="2"><input type="radio" id="JobLock" name="JobLock" value="0" checked="checked"/>解锁状态<input type="radio" id="JobLock" name="JobLock" value="1"/>锁定状态</td>
</tr>
<tr>
<td class="Right" align="right">&nbsp;</td>
<td colspan="4"><input type="submit" value="保 存" class="Button"> <input type="button" value="关闭窗口" class="Button" onClick="top.DeleteTabTitle('Job_ContentAdd')"></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>