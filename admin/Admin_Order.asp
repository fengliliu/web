<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<SCRIPT language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�еĲ�Ʒ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;	 
}
</SCRIPT>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"><div align="center"><strong>ѯ�۹��� 
              <br>
              </strong></div></td>
        </tr>
        <tr> 
          <td height="16"> 
            <%
flag="��δ����"
set rs=server.createobject("adodb.recordset")
sqltext="select * from OrderList  order by OrderTime desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=20

'ȡ��ҳ��,���ж��û�������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
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

'��ʾ���ӵ��ӳ���
Sub list()%> 
            <table width="100%" border="0" cellspacing="1" cellpadding="4">
              <tr bgcolor="#A4B6D7" class="tr_southidc"> 
                <td width="18%" height="25"> 
                  <div align="center">�������</div></td>
                <td width="20%"> 
                  <div align="center">�ͻ�����</div></td>
                <td width="22%"> 
                  <div align="center">��ϵ�绰</div></td>
                <td width="15%"> 
                  <div align="center">�������</div></td>
                <td width="15%"> 
                  <div align="center">ѯ������</div></td>
                <td width="10%"> 
                  <div align="center">����</div></td>
              </tr>
              <%
if not rs.eof then
i=0
do while not rs.eof
%>
              <tr bgcolor="#ECF5FF"> 
                <td> 
                  <div align="center"><%=rs("OrderNum")%></div></td>
                <td> 
                  <div align="center"><%=rs("UserName")%></div></td>
                <td> 
                  <div align="center"><%=rs("Phone")%></div></td>
                <td> 
                  <div align="center"> 
                    <%If rs("Flag")="Yes" Then%>
                    �Ѵ���
                    <%else%>
                    <b><font color="#FF0000">δ����</font></b> 
                    <%End If%>
                  </div></td>
                <td> 
                  <div align="center"> 
                    <%response.write "<a href='Admin_OrderDetail.asp?ID="&rs("OrderNum")&"&page="&CurrentPage&"'  >��ϸ����</a>"%>
                  </div></td>
                <td> 
                <div align="center">
				 <a href="Admin_OrderDel.asp?OrderNum=<%=rs("OrderNum")%>&page=<%=CurrentPage%>&Action=Del" onClick="return ConfirmDel();">ɾ��</a>
				  </div></td>
              </tr>
              <%
i=i+1
if i >= MaxPerpage then exit do
rs.movenext
loop
end if
%>
              <tr bgcolor="#A4B6D7"> 
                <td height="25" colspan="6"> 
                  <div align="right"> 
                    <%
Response.write "ȫ��-"
Response.write "��" & Cstr(Rs.RecordCount) & "��ѯ����Ϣ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='Admin_Order.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='Admin_Order.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='Admin_Order.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='Admin_Order.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='Admin_Order.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='Admin_Order.asp?page="&i&"&id="&id&"'>"&i&"</option>"
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
conn.close
%> </td>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->