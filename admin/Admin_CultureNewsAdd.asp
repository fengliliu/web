<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
title=Trim(Request("title"))
content=Request("content")
Language=Request("Language")
%>
<%if Request.QueryString("mark")="southidc" then%>
<%
If title="" Then
response.write "SORRY <br>"
response.write "请输入内容标题!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If content="" Then
response.write "SORRY <br>"
response.write "请输入内容!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from culture "
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=content
rs("Language")=Language
rs("counter")=0
rs("time")=Date()
rs.update
rs.close
response.redirect "Admin_Culture.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<table width="560" border="0" align="center" cellpadding="2" cellspacing="1" class="table_southidc">
  <tr> 
    <td class="back_southidc" height="25"> <div align="center"><strong>添加企业文化文章 
        <br>
        </strong></div></td>
  </tr>
  <tr> 
    <form method="post" action="Admin_CultureNewsAdd.asp?mark=southidc">
      <td><div align="center"> 
          <table width="100%" border="0" cellpadding="0" cellspacing="1">
            <tr bgcolor="#E7E7E7"> 
              <td width="9%" height="25" bgcolor="#ECF5FF"><div align="center">标&nbsp;&nbsp;题:</div></td>
              <td width="59%" bgcolor="#ECF5FF"><input type="text" name="title" class="inputtext" maxlength="60" size="45"></td>
              <td width="13%" bgcolor="#ECF5FF"><div align="center">语言:</div></td>
              <td width="19%" bgcolor="#ECF5FF"><select name="Language" id="Language">
                  <option value="1">中文</option>
                  <option value="0">英文</option>
                </select></td>
            </tr>
            <tr bgcolor="#E7E7E7"> 
              <td height="25" colspan="4" bgcolor="#ECF5FF"> <div align="center">内&nbsp;&nbsp;容</div></td>
            </tr>
            <tr bgcolor="#E7E7E7"> 
              <td height="300" colspan="4" bgcolor="#ECF5FF"> <input type="hidden" name="content" value=""> 
                <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe> 
              </td>
            </tr>
            <tr bgcolor="#ECF5FF"> 
              <td height="25" colspan="4"> <div align="center"> 
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
<!-- #include file="Inc/Foot.asp" -->