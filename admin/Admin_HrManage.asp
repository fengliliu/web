<%@ LANGUAGE="VBScript" %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/articlechar.inc"-->
<!-- #include file="Inc/Head.asp" -->

<%
'ɾ��ӦƸ����
if Request.QueryString("action")="Del" then
 dim sql
 dim rs
 on error resume next
 set rs=server.createobject("adodb.recordset")
 sql="delete from HrDemandAccept where ID="&request("id")
 rs.open sql,conn,1,1
 set rs=nothing
 conn.close
 set conn=nothing
 response.redirect "Admin_HrManage.asp"
end if
%>

<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ���˼�¼��"))
     return true;
   else
     return false;
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <br> <table width="560" border="0" cellpadding="0" cellspacing="0" >
        <tr> 
          <td class="back_southidc" height="28"> 
            <div align="center"><strong>ӦƸ���� <br>
              </strong></div></td>
        </tr>
        <tr> 
          <td><div align="center"> 
<% 
set rs=server.createobject("adodb.recordset") 
sql="select * from HrDemandAccept order by id desc" 
rs.open sql,conn,1,1 


dim MaxPerPage
MaxPerPage=10

dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call list

Sub list()%>

<%dim i 
i=0 
do while not rs.eof%>
              <table class="table_southidc" width="559" border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td width="17%" height="25" bgcolor="#A4B6D7"> <div align="center">ӦƸ��λ</div></td>
                  <td colspan="2" bgcolor="#A4B6D7">&nbsp;&nbsp;<b><%=rs("Quarters")%></b> &nbsp; </td>
                  <td bgcolor="#A4B6D7"><a href="Admin_HrManage.asp?id=<%=rs("ID")%>&amp;action=Del" onClick="return ConfirmDel();">ɾ��</a>&nbsp;&nbsp;&nbsp;<a href="AdminMaillist.asp?Email=<%=rs("email")%>&Name=<%=rs("Name")%>">����</a> 
                  </td>
                </tr>
                <%if rs("email")<>"" then%>
                <tr> 
                  <td height="10" bgcolor="#ECF5FF"> <div align="center">�� ��</div></td>
                  <td width="34%" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Name")%></td>
                  <td width="13%" bgcolor="#ECF5FF"> <div align="center">�� ��</div></td>
                  <td width="36%" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("sex")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"> <div align="center">��������</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("birthday")%></td>
                  <td bgcolor="#ECF5FF"> <div align="center">����״��</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("marry")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"><div align="center">����</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Stature")%></td>
                  <td bgcolor="#ECF5FF"><div align="center">�����</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Residence")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"> <div align="center">��ҵԺУ</div></td>
                  <td colspan="3" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("school")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"> <div align="center">ѧ ��</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("studydegree")%></td>
                  <td bgcolor="#ECF5FF"> <div align="center">ר ҵ</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("specialty")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"> <div align="center">��ҵʱ��</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("gradyear")%></td>
                  <td bgcolor="#ECF5FF"> <div align="center"><font color="#FF0000">ӦƸ����</font></div></td>
                  <td bgcolor="#ECF5FF"><font color="#FF0000">&nbsp;&nbsp;<%=rs("Adddate")%></font></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"> <div align="center">�� ��</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("phone")%></td>
                  <td bgcolor="#ECF5FF"><div align="center">�ֻ�</div></td>
                  <td bgcolor="#ECF5FF"><font color="#FF0000">&nbsp;&nbsp;<%=rs("Mobile")%></font></td>
                </tr>
                <tr> 
                  <td height="24" bgcolor="#ECF5FF"> <div align="center">��ϵ��ַ</div></td>
                  <td colspan="3" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("add")%></td>
                </tr>
                <tr> 
                  <td height="22" bgcolor="#ECF5FF"><div align="center">�ʱ�</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Postcode")%></td>
                  <td bgcolor="#ECF5FF"><div align="center">E-mail</div></td>
                  <td bgcolor="#ECF5FF">&nbsp;&nbsp;<a href="AdminMaillist.asp?Email=<%=rs("email")%>"><%=rs("email")%></a></td>
                </tr>
                <%end if%>
                <tr> 
                  <td height="25" bgcolor="#ECF5FF"> <div align="center">��������</div></td>
                  <td height="25" colspan="3" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Edulevel")%></td>
                </tr>
                <tr> 
                  <td height="25" bgcolor="#ECF5FF"> <div align="center">���˼���</div></td>
                  <td height="25" colspan="3" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("Experience")%></td>
                </tr>
              </table>			  
			  
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <% 
i=i+1 
if i >= MaxPerpage then exit do 
rs.movenext 
loop 
%>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="30" bgcolor="#A4B6D7"> 
                    <div align="right"> 
                      <%
Response.write "ȫ��"
Response.write "��" & Cstr(Rs.RecordCount) & "��ӦƸ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='Admin_HrManage.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='Admin_HrManage.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='Admin_HrManage.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='Admin_HrManage.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='Admin_HrManage.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='Admin_HrManage.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>ҳ"
%>
                    </div></td>
                </tr>
              </table>
<%
End sub
rs.close
%>			  
            </div></td>
        </tr>
      </table>
      <br> <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td><div align="center">

            </div></td>
        </tr>
      </table>
      <br> <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->