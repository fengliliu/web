<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
dim BigClassID,Action,rs,NewBigClassName,OldBigClassName,FoundErr,ErrMsg
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(Request("NewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))

if BigClassID="" then
  response.Redirect("Admin_MovieClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from BigClass_Movie where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�˲�Ʒ���಻���ڣ�</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��Ʒ����������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("BigClassName")=NewBigClassName
    	 	rs.update
			rs.Close
	     	set rs=Nothing
			if NewBigClassName<>OldBigClassName then
				conn.execute "Update SmallClass_Movie set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				conn.execute "Update Movie set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
     		end if		
			call CloseConn()
     		Response.Redirect "Admin_MovieClassManage.asp" 
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center" valign="top"><br>
            <strong>�� �� �� �� �� ��</strong> <br> <br> <form name="form1" method="post" action="MovieClassModifyBig.asp">
              <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
                <tr bgcolor="#A4B6D7" class="title"> 
                  <td height="25" colspan="2" align="center"><strong>�޸Ĵ�������</strong></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" bgcolor="#C0C0C0"><strong>����ID��</strong></td>
                  <td width="204" bgcolor="#E3E3E3"><%=rs("BigClassID")%> <input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("BigClassID")%>"> 
                    <input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("BigClassName")%>"></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" bgcolor="#C0C0C0"><strong>�������ƣ�</strong></td>
                  <td bgcolor="#E3E3E3"> <input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("BigClassName")%>" size="20" maxlength="30"></td>
                </tr>
                <tr class="tdbg"> 
                  <td align="center" bgcolor="#C0C0C0">&nbsp;</td>
                  <td align="center" bgcolor="#E3E3E3"> <div align="left"> 
                      <input name="Action" type="hidden" id="Action" value="Modify">
                      <input name="Save" type="submit" id="Save" value=" �� �� ">
                    </div></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <%
	end if
end if
rs.close
set rs=nothing
%> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->