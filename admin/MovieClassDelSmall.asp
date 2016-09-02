<%@language=vbscript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<%
dim SmallClassID,sql
SmallClassID=trim(Request("SmallClassID"))
if SmallClassID<>"" then
	sql="delete from SmallClass_Movie where SmallClassID="&Cint(SmallClassID)&""
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Admin_MovieClassManage.asp"
%>


