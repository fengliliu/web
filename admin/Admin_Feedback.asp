<%@ LANGUAGE="VBScript" %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/articlechar.inc"-->
<!-- #include file="Inc/Head.asp" -->
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
    <td align="center" valign="top"> <br> 
      <table width="760" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" >
        <tr> 
          <td height="28" class="back_southidc"> 
            <div align="center"><strong>���Թ��� <br>
              </strong></div></td>
        </tr>
        <tr> 
          <td> 
						 <table width="100%" border="0" cellspacing="1" cellpadding="4">
						              <tr bgcolor="#A4B6D7" class="tr_southidc"> 
						                <td width="20%" height="25"> 
						                  <div align="center">����ʱ��</div></td>
						                <td width="10%"> 
						                  <div align="center">������</div></td>
						                <td width="20%"> 
						                  <div align="center">�绰</div></td>
						                <td width="15%"> 
						                  <div align="center">��˾</div></td>
						                <td width="30%"> 
						                  <div align="center">����</div></td>
						                <td> 
						                  <div align="center" >����</div></td>
						              </tr>
<% 
const MaxPerPage=5 '��ҳ��ʾ�ļ�¼���� 
dim sql 
dim rs 
dim totalPut 
dim CurrentPage 
dim TotalPages 
dim i,j 
 
 
if not isempty(request("page")) then 
  currentPage=cint(request("page")) 
else 
  currentPage=1 
end if 
set rs=server.createobject("adodb.recordset") 
sql="select * from Feedback order by id desc" 
rs.open sql,conn,1,1 
 
if rs.eof and rs.bof then 
  response.write "<p align='center'>��û���κ�����!</p>" 
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
  showpages 
  showContent 
  showpages 
else 
 if (currentPage-1)*MaxPerPage<totalPut then 
  rs.move (currentPage-1)*MaxPerPage 
  dim bookmark 
  bookmark=rs.bookmark 
  showpages 
  showContent 
  showpages 
 else 
  currentPage=1 
  showContent 
 end if 
end if 
rs.close 
end if 
set rs=nothing 
conn.close 
set conn=nothing 


sub showContent 
dim i 
i=0 
do while not (rs.eof or err) %>
             <tr bgcolor="#ECF5FF"> 
                <td> 
                  <div align="center"><%=rs("Time")%></div></td>
                <td> 
                  <div align="center"><%=rs("Receiver")%></div></td>
                <td> 
                  <div align="center"><%=rs("Phone")%></div></td>
                <td> 
                  <div align="center"><%=rs("CompanyName")%></div></td>
                <td> 
                  <div align="center"><%=rs("Content")%></div></td>
                <td> <a href="Admin_FeedbackDel.asp?id=<%=rs("id")%>&Action=Del" onClick="return ConfirmDel();">ɾ��</a></td>
              </tr>
              <% i=i+1 
 if i>=MaxPerPage then exit do 
 rs.movenext 
loop 
end sub 
sub showpages() 
 dim n 
 if (totalPut mod MaxPerPage)=0 then 
   n= totalPut \ MaxPerPage 
 else 
   n= totalPut \ MaxPerPage + 1 
 end if 
 if n=1 then 
  exit sub 
 end if 
 dim k 
 response.write "<p align='left'>&gt;&gt; ���Է�ҳ " 
 for k=1 to n 
 if k=currentPage then 
   response.write "[<b>"+Cstr(k)+"</b>] " 
 else 
   response.write "[<b>"+"<a href='Admin_Feedback.asp?page="+cstr(k)+"'>"+Cstr(k)+"</a></b>] " 
 end if 
 next 
end sub 
%>
      </table>
            </td>
        </tr>
      </table>
      <br> <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->