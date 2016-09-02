<%@ LANGUAGE="VBScript" %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
dim SQL, Rs, OrderNum,CurrentPage
CurrentPage=request("Page")
OrderNum=trim(request("OrderNum"))

set rs=server.createobject("adodb.recordset")
sqltext="delete from OrderList where OrderNum='"&OrderNum&"'"
rs.open sqltext,conn,3,3
set rs=nothing

set rs=server.createobject("adodb.recordset")
sqltext="delete from OrderDetail where OrderNum='"&OrderNum&"'"
rs.open sqltext,conn,3,3
set rs=nothing
conn.close

response.redirect "Admin_Order.asp?page="&CurrentPage
%>