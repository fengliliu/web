<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass_down order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    



function CheckForm()
{
     if (document.myform.title.value.length == 0) {
		alert("��������û����д.");
		document.myform.title.focus();
		return false;
	}
	if (document.myform.System.value.length == 0) {
		alert("����ϵͳû����д");
		document.myform.System.focus();
		return false;
	}
		if (document.myform.DownloadUrl.value.length == 0) {
		alert("���ص�ַû����д");
		document.myform.DownloadUrl.focus();
		return false;
	}
		
	return true;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="650" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="Down_Save.asp?action=Add" target="_self">
          <tr> 
            <td class="back_southidc" height="30" colspan="3" align="center">��������</td>
          </tr>
          <tr> 
            <td width="168" height="25" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������ƣ�</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="title" type="text" class="input" size="30"></td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������</td>
            <td colspan="2" bgcolor="#ECF5FF"> <%		
        sql = "select * from BigClass_down"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "����������Ŀ��"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <select name="SmallClassName">
                <option value="" selected>��ָ��С��</option>
                <%
			sql="select * from SmallClass_down where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <% rs.movenext
				do while not rs.eof%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
                <%
			ranNum=int(9*rnd)+10
			iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			%>
              </select></td>
          </tr>
          <tr> 
            <td height="25" valign="top" bgcolor="#ECF5FF"><font color="#FF0000">*</font>����˵����</td>
            <td colspan="2" valign="top" bgcolor="#ECF5FF"> <input type="hidden" name="content" value=""> 
              <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe> 
            </td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#ECF5FF"><font color="#FF0000">*</font>����ϵͳ��</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="System" type="text"  id="System" value="Win98, Win2000, WinXP" size="30"></td>
          </tr>
          <tr>
            <td height="25" bgcolor="#ECF5FF">�������ԣ�</td>
            <td colspan="2" bgcolor="#ECF5FF"><input name="Language" type="text"  id="Language" value="��������" size="30"></td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������ͣ�</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="Softclass" type="text"  id="Softclass" value="��Ʒ����" size="30"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25">��ƷͼƬ��</td>
            <td width="224"> <input name="PhotoUrl" type="text" id="PhotoUrl" size="30" maxlength="200"></td>
            <td width="242"> <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=0" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25"><font color="#FF0000">*</font>���ص�ַ��</td>
            <td> <input name="DownloadUrl" type="text" id="PhotoUrl2" size="30" maxlength="200"></td>
            <td> <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=1" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr> 
            <td height="25" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�ļ���С��</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="FileSize" type="text" class="input" id="fileSize" size="30">
              K </td>
          </tr>
          <tr> 
            <td height="30" align="center" bgcolor="#ECF5FF"><div align="left">¼��ʱ�䣺</div></td>
            <td height="30" colspan="2" align="center" bgcolor="#ECF5FF"><div align="left"> 
                <input name="AddDate" type="text" id="AddDate" value="<%=date()%>" maxlength="50">
              </div></td>
          </tr>
          <tr> 
            <td height="30" colspan="3" align="center" bgcolor="#ECF5FF"> <input type="submit" name="Submit" value="�ύ" class="input">
              �� 
              <input type="reset" name="Submit2" value="����" class="input"> </td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->