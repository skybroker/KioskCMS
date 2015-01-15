<!--#include File="Config/Config.asp"-->
<!--#include File="Include/Class_Function.asp"-->
<%
ID	=	trim(request.QueryString("ID"))
if not IsNumeric(id) then
	response.Write("传递参数出错！")
	response.End()
end if
'------------------------------------------------------------------------------
conn.execute("update NewsInfo set NewsClick=NewsClick+1 where id="&id)
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql = "Select ClassID,NewsAuthor,NewsTitle,NewsBPic,NewsContent,PostTime,NewsClick From NewsInfo Where NewsLock=0 and ID="&ID&""
Rs.Open Sql,Conn,1,1
If Not (Rs.Eof Or Rs.Bof ) Then
	ClassID		=	rs("ClassID")
	NewsAuthor=	rs("NewsAuthor")
	NewsTitle	=	trim(rs("NewsTitle"))
	NewsContent	=	rs("NewsContent")
	PostTime	=	rs("PostTime")
	NewsClick	=	rs("NewsClick")
	NewsBPic	=	rs("NewsBPic")
End If
Rs.Close
set rs = nothing
'-----------------------------------------------------------------------------

title	=	NewsTitle	'页面title
call showhead()		'-----显示头部导航
%>
  <div id="content">
    <div class="left">
		<%call newlistleft()%>
    </div>
    <div class="right">
      <div class="indexcontent">
        <div class="about">
          <div class="linetitle"><%=GetSubNavName("NewsClass",ClassID)%></div>
          <div class="contacts">
          	<div class="title">
            	<h1><%=Title%></h1>
              作者：<%=NewsAuthor%>&nbsp;&nbsp;&nbsp;&nbsp;浏览次数：<%=NewsClick%>&nbsp;&nbsp;&nbsp;&nbsp;发布时间：<%=PostTime%>
            </div>
            <%if NewsBPic<>"" then%>
            <img src="<%=NewsBPic%>" border="0" class="img_content" align="middle"/>
            <%end if%>
            <%=NewsContent%>
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