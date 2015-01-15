<%
Function SendMail(ToMail,Subject,MailBody)
	Const cdoSendUsingMethod="http://schemas.microsoft.com/cdo/configuration/sendusing" 
	Const cdoSendUsingPort=2 
	Const cdoSMTPServer="http://schemas.microsoft.com/cdo/configuration/smtpserver" 
	Const cdoSMTPServerPort="http://schemas.microsoft.com/cdo/configuration/smtpserverport" 
	Const cdoSMTPConnectionTimeout="http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout" 
	Const cdoSMTPAuthenticate="http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" 
	Const cdoBasic=1 
	Const cdoSendUserName="http://schemas.microsoft.com/cdo/configuration/sendusername" 
	Const cdoSendPassword="http://schemas.microsoft.com/cdo/configuration/sendpassword" 
	
	Dim objConfig
	Dim objMessage
	Dim Fields 
	
	Set objConfig = Server.CreateObject("CDO.Configuration") 
	Set Fields = objConfig.Fields 
	
	With Fields 
	.Item(cdoSendUsingMethod) = cdoSendUsingPort 
	.Item(cdoSMTPServer) = SMTPServer 
	.Item(cdoSMTPServerPort) = 25 
	.Item(cdoSMTPConnectionTimeout) = 10 
	.Item(cdoSMTPAuthenticate) = cdoBasic 
	.Item(cdoSendUserName) = SMTPUserName
	.Item(cdoSendPassword) = SMTPUserPass 
	.Update 
	End With 
	
	Set objMessage = Server.CreateObject("CDO.Message") 
	Set objMessage.Configuration = objConfig 
	
	With objMessage 
	.To = ToMail
	.From = SmtpFormMail
	.Subject = Subject
	.HTMLBody = MailBody
	'.TextBody = MailBody 
	
	.Send 
	End With 
	
	Set Fields = Nothing 
	Set objMessage = Nothing 
	Set objConfig = Nothing
End Function
Function SendJMail(ToMail,Subject,MailBody)
	Set Msg=Server.CreateObject("Jmail.Message") 
	Msg.Silent=true 
	Msg.Logging = true 
	Msg.Charset = "gb2312" 
	Msg.MailServerUserName = SMTPUserName'����smtp��������֤��½�� 
	Msg.MailServerPassword = SMTPUserPass'����smtp��������֤���� 
	Msg.From = SmtpFormMail'������ 
	Msg.FromName = SmtpFormMail'���������� 
	Msg.AddRecipient ToMail'�ռ���  
	Msg.Subject = Subject'���� 
	Msg.Body = MailBody'���� 
	Msg.Priority = 1'�趨�ʼ����ȼ�1Ϊ������3Ϊ������5Ϊ������
	Msg.Send (SMTPServer)'�ʼ������� 
	Set Msg = Nothing
End Function
%>