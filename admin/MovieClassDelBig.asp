<%@language=vbscript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<%
dim BigClassName,sql
BigClassName=trim(Request("BigClassName"))
if BigClassName<>"" then
	sql="delete from BigClass_Movie where BigClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from SmallClass_Movie where BigClassName='" & BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "Admin_MovieClassManage"
%>


