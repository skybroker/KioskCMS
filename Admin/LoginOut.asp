<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
Response.Cookies("CNVP_CMS2")("UserName")=""
Response.Redirect("Admin_Login.asp")
%>