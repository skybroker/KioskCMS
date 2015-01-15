<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<%
ID	=	trim(request.QueryString("ID"))
if not IsNumeric(id) then
	response.Write("传递参数出错！")
	response.End()
end if
'------------------------------------------------------------------------------
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql = "Select NavTitle,NavContent,NavParent From SiteExplain Where ID="&ID&""
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof Or Rs.Bof ) Then
	NavTitle	=	rs("NavTitle")
	NavContent=	Rs("NavContent")
	NavParent	=	rs("NavParent")
End If
Rs.Close
set rs = nothing
'-----------------------------------------------------------------------------

title	=	NavTitle	'页面title
call showhead()		'-----显示头部导航
%>
  <div id="content">
    <div class="left">
		<%call showleft()%>
    </div>
    <div class="right">
      <div class="banner"><img src="images/banner.jpg" width="675" height="100" border="0" alt="<%=SiteName%>"/></div>
      <div class="indexcontent">
        <div class="about">
          <div class="linetitle"><%=Title%></div>
          <div class="contacts">
            <%=NavContent%>
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