<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%if Request.QueryString("mark")="southidc" then%>
<%
explain=server.htmlencode(Trim(Request("explain")))
CompVisualize=server.htmlencode(Trim(Request("CompVisualize")))
OrderID=server.htmlencode(Trim(Request("OrderID")))
if OrderID="" then OrderID=1
%>
<%
If explain="" Then
response.write "SORRY <br>"
response.write "广告名称不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If CompVisualize="" Then
response.write "SORRY <br>"
response.write "广告图片不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CompVisualize"
rs.open sql,conn,1,3
rs.addnew
rs("Explain")=Explain
rs("OrderID")=OrderID
rs("CompVisualize")=CompVisualize
rs("Time")=Date()
rs.update
rs.close
response.redirect "Admin_CompVisualize.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc"> <div align="center"><strong>添加广告 <br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" name="myform" action="Admin_CompVisualizeAdd.asp?mark=southidc">
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td> <div align="center">排序ID</div></td>
                    <td> <input name="OrderID" type="text" id="OrderID" value="1" size="10" maxlength="10"></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td width="25%"> <div align="center">广告名称</div></td>
                    <td width="75%"> <textarea name="explain" cols="40" rows="8" ></textarea></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td> <div align="center">广告图片</div></td>
                    <td> <input name="CompVisualize" type="text" id="CompVisualize" size="40" maxlength="50"> 
                      <font color="#FF0000"> * 图片地址</font></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td colspan="2"> <div align="center"> 
                        <input name="submit" type="submit" value="确定">
                        &nbsp; 
                        <input name="reset" type="reset" value="取消">
                      </div></td>
                  </tr>
                  <tr> 
                    <td colspan="2">图片上传</td>
                  </tr>
                  <tr> 
                    <td colspan="2"> <div align="left"> 
                        <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=3" frameborder=0 scrolling=no width="300" height="25"></iframe>
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