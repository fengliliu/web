<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->

<script language="javascript">
<!--

function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�е���Ŀ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;	 
}

function form_onsubmit()
{
if (document.form_admin.Password.value!=document.form_admin.ConPassword.value)
{
alert ("�������ȷ�����벻һ�¡�");
document.form_admin.Password.value='';
document.form_admin.ConPassword.value='';
document.form_admin.Password.focus();
return false;
}
}
//-->
</SCRIPT>

<%if Request.QueryString("mark")="southidc" then
dim sql,rs,ID
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="delete from admin where Id="& ID
rs.open sql,conn,1,3
set rs=nothing
conn.close
response.redirect "Admin_Manage.asp"
end if
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" width="598" height="25" > <div align="center"><strong>���ӹ���Ա</strong></div></td>
        </tr>
        <tr class="tr_southidc"> 
          <FORM name="form_admin" method="post"  onsubmit="return form_onsubmit()" action="Admin_ManageSave.asp" >	
            <td><table width="50%" border="0" align="center" >
                <tr> 
                  <td height="25" colspan="2">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="29%" height="22"> <div align="right">����Ա�ʺţ�</div></td>
                  <td width="71%"><input name="UserName" type="text" id="UserName" size="16" maxlength="20"></td>
                </tr>
                <tr> 
                  <td height="22"> <div align="right">����Ա���룺</div></td>
                  <td><input name="Password" type="password" size="16" maxlength="20"></td>
                </tr>
                <tr> 
                  <td height="22"> <div align="right">����ȷ�ϣ�</div></td>
                  <td><input name="ConPassword" type="password" size="16" maxlength="20"></td>
                </tr>
                <tr> 
                  <td height="22" colspan="2"><div align="center">
                      <INPUT type=submit value='ȷ������' name=Submit2>
                    </div></td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <br> 
      <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td width="553" height="25" class="back_southidc"> <div align="center">����Ա�ʺŹ���(<font color="#FF0000">ɾ�����ܹر�</font>)</div></td>
        </tr>
        <tr class="tr_southidc">          
            <td><br> 
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="2">
                <tr bgcolor="#A4B6D7"> 
                  <td width="28%" height="25"> 
                    <div align="center">����Ա�ʺ�</div></td>
                  <td width="28%"> 
                    <div align="center">����Ա����</div></td>
                  <td width="24%"> 
                    <div align="center">����</div></td>
                  <td width="20%"> 
                    <div align="center">ɾ��</div></td>
                </tr>
<%				
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,1
do while not rs.eof
%>
                <tr bgcolor="#DFEBF2"> 
                  <td height="22"> 
                    <div align="center"><%=rs("UserName")%></div></td>
                  <td> 
                    <div align="center"><%=rs("PassWord")%></div></td>
                  <td> 
                    <div align="center">
                      <a href="admin_manageEdit.asp?id=<%=rs("ID")%>" style="cursor:pointer">�޸�����</a>
                    </div></td>
                  <td> 
                    <div align="center">
					     <a href="admin_manage.asp?mark=southidc&id=<%=rs("ID")%>" style="cursor:pointer">ɾ��</a>                
                    </div></td>
                </tr>
<%
rs.movenext
loop
rs.close
%>
              </table></td>         
        </tr>
      </table>
      </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->