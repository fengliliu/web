<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../ckeditor/ckeditor.asp"-->
<script type="text/javascript" src="../ckeditor/ckeditor.js" charset="gb2312"></script>
<%if Request.QueryString("mark")="southidc" then

id=request("id")
Title=Trim(Request("Title"))
Content=Request("Content")
Aboutusorder=Request("Aboutusorder")
Language=Request("Language")

If Title="" Then
response.write "SORRY <br>"
response.write "��������Ŀ����!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If Content="" Then
response.write "SORRY <br>"
response.write "��������Ŀ����!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Aboutus where id="&id
rs.open sql,conn,1,3

rs("Title")=Title
rs("Content")=Content
rs("Aboutusorder")=Aboutusorder
rs("Language")= 1
rs.update
rs.close
response.redirect "AdminAboutus.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Aboutus where id="&id, conn,1,3
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>�޸���Ŀ<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" name="myform" action="AdminAboutusModify.asp?mark=southidc">
            <input type=hidden name=id value=<%=id%>>
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="10%" height="25"> <div align="center">��Ŀ����:</div></td>
                    <td width="33%" height="25" bgcolor="#ECF5FF"> <input type="text" name="title" maxlength="30" size="25" value="<%=rs("Title")%>"> 
                    </td>
                    <td width="14%" bgcolor="#ECF5FF"><div align="center">�����:</div></td>
                    <td width="12%" bgcolor="#ECF5FF"><input name="Aboutusorder" type="text" id="Aboutusorder" value="<%=rs("Aboutusorder")%>" size="4"></td> 
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="6"> <div align="center">��Ŀ����</div></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="300" colspan="6"> 
                      <textarea id="Content" name="Content" class="ckeditor"><%=Server.HTMLEncode(rs("Content"))%></textarea>
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="6"> <div align="center"> 
                        <input name="submit" type="submit" value="ȷ��">
                        &nbsp; 
                        <input name="reset" type="reset" value="����">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->