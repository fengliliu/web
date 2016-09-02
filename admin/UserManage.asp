<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim strFileName
const MaxPerPage=10
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="UserManage.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from [User] order by UserID desc"
rs.Open sql,conn,1,1
%>
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ�����û���"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"><br>
      <br>
      <%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ��ע���û�"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then        
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark        		
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	else
	        	currentPage=1        
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%> 
      <table width="600" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr>
          <td class="back_southidc" height="25"><div align="center">ע���Ա����</div></td>
        </tr>
        <tr> 
          <td height="25"><div align="center">
              <table width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border">
                <tr bgcolor="#ECF5FF" class="title"> 
                  <td width="30" height="25" align="center"><strong> ���</strong></td>
                  <td width="59" height="25" align="center"><strong> �û���</strong></td>
                  <td width="29" height="25" align="center"><strong> �Ա�</strong></td>
                  <td width="87" height="25" align="center"><strong>Email</strong></td>
                  <td width="185" height="25" align="center"><strong>��˾����</strong></td>
                  <td width="39" height="25" align="center"><strong> ״̬</strong></td>
                  <td width="104" height="25" align="center"><strong> ����</strong></td>
                </tr>
                <%do while not rs.EOF %>
                <tr class="tdbg"> 
                  <td height="22" align="center" bgcolor="#ECF5FF"><%=rs("UserID")%></td>
                  <td align="center" bgcolor="#ECF5FF"><%=rs("username")%></td>
                  <td align="center" bgcolor="#ECF5FF"> 
                    <%if rs("Sex")=1 then
	    response.write "��"
	  else
	    response.write "Ů"
	  end if%>
                  </td>
                  <td bgcolor="#ECF5FF"><a href="AdminMaillist.asp?Email=<%=rs("email")%>"><%=rs("email")%></a></td>
                  <td bgcolor="#ECF5FF"> <%=rs("CompanyName")%> </td>
                  <td align="center" bgcolor="#ECF5FF"> 
                    <%
	  if rs("LockVIPUser")=true then
	  	response.write "������"
	  else
	  	response.write "����"
	  end if
	  %>
                  </td>
                  <td align="center" bgcolor="#ECF5FF"><a href="UserModify.asp?UserID=<%=rs("UserID")%>">�޸�</a>&nbsp; 
				<%if rs("VIPUser")=true then %>
                    <%if rs("LockVIPUser")=False then %>
                    <a href="UserLock.asp?Action=Lock&UserID=<%=rs("UserID")%>">VIP����</a> 
                    <%else%>
                    <a href="UserLock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">VIP����</a> 
                    <%end if
				end if%>
                    &nbsp;<a href="UserDel.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
                </tr>
                <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
              </table>
            </div></td>
        </tr>
      </table>
      <%
end sub 
%> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>