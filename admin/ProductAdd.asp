<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<!--#include file="../ckeditor/ckeditor.asp"-->
<script type="text/javascript" src="../ckeditor/ckeditor.js"></script>
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
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

  if (document.myform.Title.value=="")
  {
    alert("��Ʒ���Ʋ���Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Product_Id.value=="")
  {
    alert("��Ʒ��Ų���Ϊ�գ�");
	document.myform.Key.focus();
	return false;
  }
  
  if (document.myform.Content.value=="")
  {
    alert("��Ʒ���ݲ���Ϊ�գ�");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}

</script>
<!-- #include file="Inc/Head.asp" -->

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> <b><br>
      </b>
<form method="POST" name="myform" action="ProductSave.asp?action=add" target="_self">
		<input name="Passed" type="hidden" id="Passed" value="yes" >
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
          <tr align="center"> 
            <td class="tdbg"> <table width="100%" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
                <tr> 
                  <td class="back_southidc" height="22" colspan="3" align="right" bgcolor="#A4B6D7"><div align="center"><b>�� 
                      �� �� Ʒ</b> </div></td>
                </tr>
                <tr> 
                  <td width="159" height="22" align="right" bgcolor="#A4B6D7">�������</td>
                  <td colspan="2" bgcolor="#E3E3E3"> <strong> 
                    <%
        sql = "select * from BigClass"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "����������Ŀ��"
		else
		%>
                    <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
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
                    </select>
                    <select name="SmallClassName">
                      <option value="" selected>��ָ��С��</option>
                      <%
			sql="select * from SmallClass where BigClassName='" & selclass & "'"
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
                    </select>
                    </strong></td>
                </tr>
                <tr> 
                  <td width="159" height="22" align="right" bgcolor="#A4B6D7">��Ʒ��ţ�</td>
                  <td colspan="2" bgcolor="#E3E3E3"> <strong> 
                    <input name="Product_Id" type="text"
           id="Product_Id" value="<%=iddata%>" size="10" maxlength="10">
                    <font color="#FF0000">*��Ʒ��Ų�������ͬ�����㲻��ȷ�����ظ�������Ķ�����</font></strong></td>
                </tr>
                <tr> 
                  <td width="159" height="21" align="right" bgcolor="#A4B6D7">��Ʒ���ƣ�</td>
                  <td colspan="2" bgcolor="#E3E3E3"> <strong> 
                    <input name="Title" type="text"
           id="Title2" size="50" maxlength="80">
                    <font color="#FF0000">*</font></strong></td>
                </tr>      
                <tr>
                  <td bgcolor="#A4B6D7"><div align="right">��Ʒ˵����</div></td>
                  <td colspan="2" bgcolor="#E3E3E3"><textarea id="Content" name="Content" class="ckeditor"></textarea></td>
                </tr>
                <tr> 
                  <td align="right" bgcolor="#A4B6D7">��ƷͼƬ�� </td>
                  <td width="214" height="21" bgcolor="#E3E3E3"> <input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="30" maxlength="50"> 
                  </td>
                  <td width="231" bgcolor="#E3E3E3"><iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=6" frameborder=0 scrolling=no width="300" height="25"></iframe> 
                  </td>
                </tr> 
				<tr> 
                  <td height="22" align="right" bgcolor="#A4B6D7">��ҳ��ʾ��</td>
                  <td colspan="2" bgcolor="#E3E3E3"><input name="Elite" type="checkbox" id="Elite" value="yes" checked>
                    ��<font color="#0000FF">�����ѡ�еĻ�������ҳ��ʾ��</font></td>
                </tr>
                <tr> 
                  <td height="22" align="right" bgcolor="#A4B6D7">¼��ʱ�䣺</td>
                  <td colspan="2" bgcolor="#E3E3E3"><input name="UpdateTime" type="text" id="UpdateTime2" value="<%=now()%>" maxlength="50">
                    ��ǰʱ��Ϊ��<%=date()%> ע�ⲻҪ�ı��ʽ��</td>
                </tr>
                <tr bgcolor="#E3E3E3"> 
                  <td height="22" colspan="3" align="right"> <div align="center"> 
                      <input
  name="Add" type="submit"  id="Add2" value=" �� �� " onClick="document.myform.action='ProductSave.asp?action=add';document.myform.target='_self';">
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <div align="center">
          <p>&nbsp;  </p>
        </div>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->