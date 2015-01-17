<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<%
title	=	"首页"	'页面title
call showhead()		'-----显示头部导航
NavParent	=	1
%>
  <div id="content">
    <div class="left">
      <%call showleft()%>
    </div>
    <div class="right">
      <div class="banner"><img src="images/banner.jpg" width="675" height="100" border="0" alt="<%=SiteName%>"/></div>
      <div class="indexcontent">
        <div class="about">
          <div class="linetitle">展会概况</div>
          <div class="contacts">
<%
'------------------------------------------------------------------------------
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql = "Select NavContent From SiteExplain Where ID=2"
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof Or Rs.Bof ) Then
	NavContent=	Rs("NavContent")
End If
Rs.Close
set rs = nothing
SubNavContent	=	CutStr(NavContent,300)
response.Write(SubNavContent)
'-----------------------------------------------------------------------------
%>
          </div>
          <div class="linetitle">展会图片</div>
					<div class="contacts">
<%
	sqlstr = "select top 8 id,NewsTitle,PostTime,NewsSPic,NewsBPic from NewsInfo where ClassID in (2,3) and  NewsLock=0 order by NewsOrder desc "
	set rs=server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,1
		do until rs.eof
		id				=	rs("id")
		NewsTitle	=	rs("NewsTitle")
		PostTime	=	formatdatetime(rs("PostTime"),2)
		NewssPic	=	rs("NewssPic")
		NewsBPic	=	rs("NewsBPic")
		subNewsTitle	=	CutStr(NewsTitle,10)
%> 
  <dl>
    <dt><img alt="<%=NewsTitle%>" src="<%=NewssPic%>" border="0" /></dt>
    <dd><a href="news.asp?id=<%=id%>" target="_blank"><%=NewsTitle%></a></dd>
  </dl>
<%
		rs.movenext
		loop
	rs.close
	set rs = nothing
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