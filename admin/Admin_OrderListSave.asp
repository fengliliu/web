<%@ language="VBScript"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
flag="Yes"
set rs=server.createobject("adodb.recordset")
sqltext="select Flag from OrderList where OrderNum='"&request("OrderNum")&"'"
rs.open sqltext,conn,1,1
if rs("Flag")="Yes" then
  rs.close
  Response.Redirect "Admin_OrderListSave.asp?msg=��ѯ�۵����Ѿ�������"
Else
set rs=server.createobject("adodb.recordset")
 sqltext="Update OrderList set Flag='"&flag&"' where OrderNum='"&request("OrderNum")&"'"
 rs.open sqltext,conn,1,3
 response.redirect "Admin_Ordermessagebox.asp?msg=ѯ�۵�����������ϣ�"
 rs.close
end if
%>