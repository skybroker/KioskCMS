<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<%
ClassID	=	trim(request.QueryString("ID"))
if not IsNumeric(ClassID) and ClassID<>"" then
	response.Write("传递参数出错！")
	response.End()
end if
'------------------------------------------------------------------------------
if ClassID<>"" then
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql = "Select NavTitle,NavRemark From NewsClass Where ID="&ClassID&""
	Rs.Open Sql,Conn,1,1
	If Not (Rs.Eof Or Rs.Bof ) Then
		NavTitle	=	rs("NavTitle")
		NavRemark=	Rs("NavRemark")
	End If
	Rs.Close
	set rs = nothing
else
	NavTitle	=	"信息列表"
end if
'-----------------------------------------------------------------------------

title	=	NavTitle	'页面title
call showhead()		'-----显示头部导航
%>
  <div id="content">
    <div class="left">
		<%call newlistleft()%>
    </div>
    <div class="right">
      <div class="indexcontent">
        <div class="about">
          <div class="linetitle"><%=Title%></div>
          <div class="contacts">
<%
if ClassID="2" or ClassID="3" then
	call ShowPicList()
else
	call ShowTextList()
end if
%>	
          </div>
        </div>
        <div class="clear"></div>
      </div>
    </div>
  </div>
<%
'------显示页面尾部--
call showfoot()
%>