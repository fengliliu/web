<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../ckeditor/ckeditor.asp"-->
<script type="text/javascript" src="../ckeditor/ckeditor.js"></script>

<%if Request.QueryString("mark")="southidc" then%>
<%
Title=Trim(Request("Title"))
Content=Request("Content")
Aboutusorder=Request("Aboutusorder")
Language=Request("Language")
%>
<%
If Title="" Then
response.write "SORRY <br>"
response.write "请输入栏目名称!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If Content="" Then
response.write "SORRY <br>"
response.write "请输入栏目内容!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Aboutus"
rs.open sql,conn,1,3
rs.addnew
rs("Title")=Title
rs("Content")=Content
rs("Aboutusorder")=Aboutusorder
rs("Language")=1
rs.update
rs.close
response.redirect "AdminAboutus.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>添加栏目<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" name="myform" action="AdminAboutusAdd.asp?mark=southidc">
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="16%" height="23" bgcolor="#ECF5FF"> <div align="center">栏目名称:</div></td>
                    <td width="34%" height="25"> <input type="text" name="Title" class="inputtext" maxlength="30" size="25"> 
                    </td>
                    <td width="12%"><div align="center">排序号:</div></td>
                    <td width="8%"><input name="Aboutusorder" type="text" id="Aboutusorder" size="4"></td> 
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="6"><div align="center">栏目内容</div></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="300" colspan="6"> <textarea id="Content" name="Content" class="ckeditor"></textarea>
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td colspan="6"> <div align="center"> 
                        <input name="submit" type="submit" value="确定">
                        &nbsp; 
                        <input name="reset" type="reset" value="取消">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table>
      <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->