<%@ LANGUAGE="VBScript" %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/articlechar.inc"-->
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ���˼�¼��"))
     return true;
   else
     return false;
}
</script>
<%
if Request.QueryString("no")="yes" then
 id= Trim(Request("id"))
 Set Rs = Server.CreateObject("ADODB.Recordset")
 Rs.Open "DELETE FROM Affiche Where id="&id,Conn,2,3,1
 Set Rs= Nothing
 Response.Redirect "Admin_AddAffiche.asp"
end if

if Request.QueryString("mark")="southidc" then
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û���빫�����ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select top 1 * from Affiche"
rs.open sql,conn,1,3

rs.addnew
if request.form("html")="on" then
rs("content")=request.form("content")
else
rs("content")=htmlencode2(request.form("content"))
end if
rs("title")=request.form("title")
rs("time")=date()
rs.update
rs.close
response.redirect "Admin_AddAffiche.asp"
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select  * from Affiche"
rs.open sql,conn,1,1
%>

<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <br> <br> 
      <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>������վ���� <br>
              </strong></div></td>
        </tr>
        <tr class="tr_southidc"> 
          <form method="POST" action="Admin_AddAffiche.asp?mark=southidc">
            <input type=hidden readonly name="Name"   value="����Ա">
            <td bgcolor="#A4B6D7">
<div align="center"> 
                <table width="100%" border="0" cellspacing="2" cellpadding="3">
                  <tr> 
                    <td width="22%" height="25" bgcolor="#A4B6D7"> 
                      <div align="center">�������:</div></td>
                    <td width="78%">
<input name="title" type="text" size="45" maxlength="100">
                    </td>
                  </tr>
                  <tr> 
                    <td height="22" bgcolor="#A4B6D7"> 
                      <div align="center">��������:</div></td>
                    <td>
<textarea rows="8" name="Content" cols="45"></textarea>
                    </td>
                  </tr>
                  <tr> 
                    <td height="25" colspan="2" bgcolor="#A4B6D7"> 
                      <div align="center">�Ƿ�֧��html 
                        <input type="checkbox" name="html" value="on">
                        <input type="submit" value="�ύ����" name="B1">
                        <input type="reset" value="ȫ����д" name="B2">
                      </div></td>
                  </tr>
                </table>
                <table width="100%" border="0" cellspacing="2" cellpadding="3">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="63%" height="25" bgcolor="#ECF5FF"> <div align="center">�������</div>
                      <div align="center"></div></td>
                    <td width="18%"> <div align="center">����ʱ��</div></td>
                    <td width="9%"> <div align="center">����</div></td>
                    <td width="10%"> <div align="center">����</div></td>
                  </tr>
                  <% do while not Rs.eof %>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22" bgcolor="#ECF5FF"> <div align="center"> <%=rs("Title")%></div>
                      <div align="center"></div></td>
                    <td> <div align="center"><%= FormatDateTime(rs("time"),2) %></div></td>
                    <td> <div align="center"><a href="Admin_editAffiche.asp?id=<%=Rs("id")%>">�޸�</a></div></td>
                    <td> <div align="center"><a href="Admin_AddAffiche.asp?id=<%=Rs("id")%>&amp;no=yes" onClick="return ConfirmDel();">ɾ��</a></div></td>
                  </tr>
                  <%Rs.MoveNext 
loop 
%>
                </table>
              </div></td>
          </form>
        </tr>
      </table>
      <br> <br> </td>
  </tr>
</table>
<%
Set Rs = Nothing 
Set Conn = Nothing 
rs.close
Conn.close
%>
<!-- #include file="Inc/Foot.asp" -->