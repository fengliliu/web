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
sql = "select * from SmallClass_New order by SmallClassID asc"
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
			if(locationid=='婚礼套系'){
				document.getElementById('price').style.display='';
				}else{
					document.getElementById('price').style.display='none';
					document.getElementById('txtprice').value='';
					}
    }    

function CheckForm()
{
    if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
    else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

	if (document.myform.title.value.length == 0) {
		alert("信息标题没有填写.");
		document.myform.title.focus();
		return false;
	}
	if (document.myform.BigClassName.value.length == 0) {
		alert("信息大类没有选");
		document.myform.BigClassName.focus();
		return false;
	}
	
		if (document.myform.user.value.length == 0) {
		alert("信息发布人没有填写");
		document.myform.user.focus();
		return false;
	}
	return true;
}
</script>
<!-- #include file="Inc/Head.asp" -->

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="News_save.asp?action=Add" target="_self">
          <tr> 
            <td class="back_southidc" height="30" colspan="2" align="center">增加信息</td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td align="right" width="119" height="25"><font color="#FF0000">*</font>信息标题：</td>
            <td width="476"> <input name="title" type="text" class="input" size="50" maxlength="200"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25" align="right"><font color="#FF0000">*</font>信息类别：</td>
            <td> <%		
        sql = "select * from BigClass_New"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
                <option value="" selected>不指定小类</option>
                <%
			sql="select * from SmallClass_New where BigClassName='" & selclass & "'"
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
          <tr id="price" bgcolor="#ECF5FF" style="display:none"> 
            <td height="25" valign="top" align="right">套系价格：</td>
            <td valign="top" colspan="2"> 
              <input name="price" id="txtprice" type="text"></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25" align="right" valign="top"><font color="#FF0000">*</font>信息内容：</td>
            <td valign="top"> 
              <textarea id="Content" name="Content" class="ckeditor"></textarea></td>
          </tr>
          <tr> 
              <td align="right" bgcolor="#ECF5FF">缩略图： </td>
              <td width="214" height="21" bgcolor="#ECF5FF"> <input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="30" maxlength="50"> <br>
              <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=6" frameborder=0 scrolling=no width="300" height="25"></iframe>
              </td>
            </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="25"><font color="#FF0000">*</font>发布人：</td>
            <td> <input name="user" type="text" class="input" value="<%=session("AdminName")%>" size="30"></td>
          </tr> 
          <tr> 
            <td height="30" align="center" bgcolor="#ECF5FF"><div align="left">录入时间：</div></td>
            <td height="30" align="center" bgcolor="#ECF5FF"><div align="left"> 
                <input name="AddDate" type="text" id="AddDate" value="<%=now()%>" maxlength="50">
              </div></td>
          </tr>
          <tr> 
            <td height="30" colspan="2" align="center" bgcolor="#ECF5FF"> <input type="submit" name="Submit" value="提交" class="input">
              　 
              <input type="reset" name="Submit" value="重置" class="input"> </td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->