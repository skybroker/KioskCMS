<!--#Include File="Conn2.asp"-->
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select SiteName,EnSiteName,SiyeKeys,EnSiyeKeys,SiteDes,EnSiteDes,SiteLogo,SiteICP,EnSiteICP,SiteCopy,EnSiteCopy,SiteAuthor,SMTPServer,SmtpFormMail,SMTPUserName,SMTPUserPass From SiteInfo"
Rs.Open Sql,Conn,1,1
	SiteName	=	Rs("SiteName")
	EnSiteName=	Rs("EnSiteName")
	SiyeKeys	=	Rs("SiyeKeys")
	EnSiyeKeys=	Rs("EnSiyeKeys")
	SiteDes		=	Rs("SiteDes")
	EnSiteDes	=	Rs("EnSiteDes")
	SiteLogo	=	Rs("SiteLogo")
	SiteICP		=	Rs("SiteICP")
	EnSiteICP	=	Rs("EnSiteICP")
	SiteCopy	=	Rs("SiteCopy")
	EnSiteCopy=	Rs("EnSiteCopy")
	SiteAuthor=Rs("SiteAuthor")
	SMTPServer=Rs("SMTPServer")
	SmtpFormMail=Rs("SmtpFormMail")
	SMTPUserName=Rs("SMTPUserName")
	SMTPUserPass=Rs("SMTPUserPass")
Rs.Close
Set Rs=Nothing
%>