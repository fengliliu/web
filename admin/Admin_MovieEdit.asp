<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%if Request.QueryString("mark")="southidc" then%>
<%
id=request("id")
Readme=server.htmlencode(Trim(Request("Readme")))
Types=server.htmlencode(Trim(Request("Types")))
MovieAddr=server.htmlencode(Trim(Request("MovieAddr")))
%>

<%
If Readme="" Then
response.write "SORRY <br>"
response.write "视频说明不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If MovieAddr="" Then
response.write "SORRY <br>"
response.write "视频地址不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Movie where id="&id
rs.open sql,conn,1,3
rs("Readme")=Readme
rs("DefaultPicUrl")=DefaultPicUrl
rs("MovieAddr")=MovieAddr
rs.update
rs.close
response.redirect "Admin_Movie.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Movie where id="&id, conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <br>      
      <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc"> <div align="center"><strong>添加视频<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" name="myform" action="Admin_MovieEdit.asp?mark=southidc">
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="27%" height="25"> <div align="center">视频说明</div></td>
                    <td width="73%"> <textarea name="Readme" cols="40" rows="5" id="Readme"><%=rs("Readme")%></textarea></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22" bgcolor="#ECF5FF"><font color="#FF0000">视频广告,请同时上传一张默认图片,方便调用</font></td>
                    <td> 
                    	<input name="DefaultPicUrl" value="<%=rs("DefaultPicUrl")%>" type="text"  size="40" maxlength="50"><br>
                    	<iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=6" frameborder=0 scrolling=no width="300" height="25"></iframe>
                      </td>
                  </tr> 
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22" bgcolor="#ECF5FF"> <div align="center">视频地址</div></td>
                    <td> <input name="MovieAddr" type="text" value="<%=rs("MovieAddr")%>"  size="40" maxlength="50">
                      <font color="#FF0000"> 视频地址</font></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="2"> <div align="center"> 
                        <input name="submit" type="submit" value="确定">
                        &nbsp; 
                        <input name="reset" type="reset" value="取消">
                      </div></td>
                  </tr>
                  <tr> 
                    <td colspan="2">视频上传</td>
                  </tr>
                  <tr> 
                    <td colspan="2"> <div align="left"> 
                        <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=7" frameborder=0 scrolling=no width="300" height="25"></iframe>
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