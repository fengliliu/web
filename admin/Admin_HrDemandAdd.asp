<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->

<%if Request.QueryString("mark")="southidc" then%>
<%
HrName=Trim(Request("HrName"))
HrRequireNum=Trim(Request("HrRequireNum"))
HrAddress=Trim(Request("HrAddress"))
HrSalary=Trim(Request("HrSalary"))
HrDate=Trim(Request("HrDate"))
HrValidDate=Trim(Request("HrValidDate"))
HrDetail=Request("Content")
%>
<%
If HrName="" Then
response.write "SORRY <br>"
response.write "招聘对象内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If HrRequireNum="" Then
response.write "SORRY <br>"
response.write "招聘人数内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If HrAddress="" Then
response.write "SORRY <br>"
response.write "工作地点内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If HrSalary="" Then
response.write "SORRY <br>"
response.write "工资待遇内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If HrValidDate="" Then
response.write "SORRY <br>"
response.write "有效期限内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if
If HrDetail="" Then
response.write "SORRY <br>"
response.write "招聘要求内容不能为空!<a href=""javascript:history.go(-1)"">返回重输</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from HrDemand "
rs.open sql,conn,1,3
rs.addnew
rs("HrName")=HrName
rs("HrRequireNum")=HrRequireNum
rs("HrAddress")=HrAddress
rs("HrSalary")=HrSalary
rs("HrDate")=HrDate
rs("HrValidDate")=HrValidDate
rs("HrDetail")=HrDetail
rs.update
rs.close
response.redirect "Admin_HrDemand.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> 
      <table width="650" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>发布招聘信息 <br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" action="Admin_HrDemandAdd.asp?mark=southidc">
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="19%" height="25"> 
                      <div align="center">招聘对象:</div></td>
                    <td width="81%"> 
                      <input name="HrName" type="text" id="HrName" size="30" maxlength="30"></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">招聘人数:</div></td>
                    <td> 
                      <input name="HrRequireNum" type="text" id="HrRequireNum" size="5" maxlength="30">
                      人</td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">工作地点:</div></td>
                    <td height="-4"> 
                      <input name="HrAddress" type="text" id="HrAddress" size="50" maxlength="60"> 
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">工资待遇:</div></td>
                    <td height="-1" bgcolor="#ECF5FF"> 
                      <input name="HrSalary" type="text" id="HrSalary" size="20" maxlength="50">
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">发布日期:</div></td>
                    <td>                    <input name="HrDate" type="text" id="HrDate" value="<%=date()%>" size="12" maxlength="50"></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">有效期限:</div></td>
                    <td height="0"> 
                      <input name="HrValidDate" type="text" id="HrValidDate" size="5" maxlength="30">
                      天</td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">招聘要求:</div></td>
                    <td height="10"> 
                       <input type="hidden" name="content" value=""> <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="400"></iframe> </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="2"> 
                      <div align="center"> 
                        <input type="submit" value="确定">
                        &nbsp; 
                        <input type="reset" value="取消">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table>
      <br> <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->