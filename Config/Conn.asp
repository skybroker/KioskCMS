<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
'On Error Resume Next
If Trim(Request.QueryString) <> "" Then strTemp =Trim(Request.QueryString)
strTemp = LCase(strTemp)
hk=0
If Instr(strTemp,"count(")<>0 then hk=1
If Instr(strTemp,"asc(")<>0 then hk=1
If Instr(strTemp,"mid(")<>0 then hk=1
If Instr(strTemp,"char(")<>0 then hk=1
If Instr(strTemp,"xp_cmdshell")<>0 then hk=1
If Instr(strTemp,"'")<>0 then hk=1
if hk=1 then
response.write("数据库连接失败，请联系网站管理员。")
response.end
hk=0
End If
Dim Conn
Dim ConnStr
ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./DataBase/#DataBase#.mdb")
'Response.write connstr
'Response.end
Set Conn=Server.CreateObject("ADODB.Connection")
ON ERROR RESUME NEXT
Conn.Open ConnStr

if Err.Number<>0 then
Err.Clear

	On Error Goto 0
	'' it is x64
	providerStr = "Microsoft.ACE.OLEDB.12.0"
	conn.open "provider="& providerStr &"; data source=" & Server.MapPath("./DataBase/#DataBase#.mdb")

	if Err.Number<>0 then
	Err.Clear
	Response.write("数据库连接失败，请联系网站管理员。")
	Response.End
	End if
End if
%>