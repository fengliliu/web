<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim ID,rs_down,FoundErr,ErrMsg
dim sql
dim count
ID=trim(request("ID"))
FoundErr=False
if ID="" then 
	response.Redirect("Down_Manage.asp")
end if
sql="select * from Download where ID=" & ID & ""
Set rs_down= Server.CreateObject("ADODB.Recordset")
rs_down.open sql,conn,1,1
if FoundErr=True then
	call WriteErrMsg()
else
%>
<%
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass_Down order by SmallClassID asc"
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

function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}
function CheckForm()
{
     if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
     else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

	if (document.myform.title.value.length == 0) {
		alert("��������û����д.");
		document.myform.title.focus();
		return false;
	}
	if (document.myform.System.value.length == 0) {
		alert("����ϵͳû����д");
		document.addNEWS.System.focus();
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
    <td align="center" valign="top">
      <table width="650" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="Down_Save.asp?action=Modify" target="_self">
          <tr align="center"> 
            <td class="back_southidc" height="30" colspan="3"><font color="#0000FF"><strong>�޸�����</strong></font></td>
          </tr>
          <tr> 
            <td width="157" height="24" align="right" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������ƣ�</td>
            <td colspan="2" valign="top" bgcolor="#ECF5FF">�� 
              <input name="title" type="text" class="input" value="<%=rs_down("title")%>" size="30"></td>
          </tr>
          <tr> 
            <td height="24" align="right" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������</td>
            <td colspan="2" valign="top" bgcolor="#ECF5FF"> �� 
              <%
			
        sql = "select * from BigClass_down"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "����������Ŀ��"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <%
		    do while not rs.eof
			%>
                <option <% if rs("BigClassName")=rs_down("BigClassName") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <select name="SmallClassName">
                <option value="" <%if rs_down("SmallClassName")="" then response.write "selected"%>>��ָ��С��</option>
                <%
			sql="select * from SmallClass_down where BigClassName='" & rs_down("BigClassName") & "'" 
			rs.open sql,conn,1,1 
			if not(rs.eof and rs.bof) then  
				do while not rs.eof %>
                <option <% if rs("SmallClassName")=rs_down("SmallClassName") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
              </select> </td>
          </tr>
          <tr> 
            <td height="19" align="right" valign="top" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������ݣ�</td>
            <td colspan="2" valign="top" bgcolor="#ECF5FF"> <input type="hidden" name="content" value="<%=Server.HTMLEncode(rs_down("Content"))%>"> 
              <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe> 
            </td>
          </tr>
          <tr> 
            <td height="24" align="right" bgcolor="#ECF5FF"><font color="#FF0000">*</font>����ϵͳ��</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="System" type="text" class="input" id="System" value="<%=rs_down("System")%>" size="30"></td>
          </tr>
          <tr>
            <td height="24" align="right" bgcolor="#ECF5FF">�������ԣ�</td>
            <td colspan="2" bgcolor="#ECF5FF"><input name="Language" type="text"  id="Language" value="<%=rs_down("Language")%>" size="30"></td>
          </tr>
          <tr> 
            <td height="24" align="right" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�������ͣ�</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="Softclass" type="text" class="input" id="Softclass" value="<%=rs_down("Softclass")%>" size="30"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right">��ƷͼƬ��</td>
            <td width="229"> <input name="PhotoUrl" type="text" id="PhotoUrl" value="<%=rs_down("PhotoUrl")%>" size="30" maxlength="200"></td>
            <td width="248"> <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=0" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><font color="#FF0000">*</font>���ص�ַ��</td>
            <td> <input name="DownloadUrl" type="text" id="PhotoUrl2" value="<%=rs_down("DownloadUrl")%>" size="30" maxlength="200"></td>
            <td> <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=1" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr> 
            <td height="24" align="right" bgcolor="#ECF5FF"><font color="#FF0000">*</font>�ļ���С��</td>
            <td colspan="2" bgcolor="#ECF5FF"> <input name="FileSize" type="text" class="input" id="fileSize" value="<%=rs_down("fileSize")%>" size="30">
              K</td>
          </tr>
          <tr> 
            <td height="24" align="right" bgcolor="#ECF5FF">¼��ʱ�䣺</td>
            <td colspan="2" bgcolor="#ECF5FF"><input name="AddDate" type="text" id="AddDate" value="<%=rs_down("AddDate")%>" maxlength="50"></td>
          </tr>
          <tr align="center"> 
            <td height="35" colspan="3"> <input type="submit" name="Submit" value="�ύ" class="input"> 
              <input type="hidden" name="Id" value="<%=Id%>">
              �� 
              <input type="reset" name="Submit2" value="����" class="input"> </td>
          </tr>
        </form>
      </table>
      <% End If
rs_down.close
set rs_down=nothing
 %></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->