<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim UserID,Action,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserID=trim(request("UserID"))
if UserID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserID=" & Clng(UserID)
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
	else
		if Action="Modify" then
			dim UserName,Password,Question,Answer,Sex,Email,Homepage,LockUser,Comane,Add,Somane,Zip,Phone,Fox
			UserName=trim(request("UserName"))
			Password=trim(request("Password"))
			Question=trim(request("Question"))
			Answer=trim(request("Answer"))
			Sex=trim(Request("Sex"))
			Email=trim(request("Email"))
			Homepage=trim(request("Homepage"))
			CompanyName=trim(request("CompanyName"))
			Add=trim(request("Add"))
			Receiver=trim(request("Receiver"))
			Postcode=trim(request("Postcode"))
			Phone=trim(request("Phone"))
			Mobile=trim(request("Mobile"))
			Fax=trim(request("Fax"))
			LockUser=trim(request("LockUser"))
			if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
				founderr=true
				errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
			else
  				if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
					errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
					founderr=true
				else
					dim sqlReg,rsReg
					sqlReg="select * from [User] where UserName='" & Username & "' and UserID<>" & UserID
					set rsReg=server.createobject("adodb.recordset")
					rsReg.open sqlReg,conn,1,1
					if not(rsReg.bof and rsReg.eof) then
						founderr=true
						errmsg=errmsg & "<br><li>�û����Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
					end if
					rsReg.Close
					set rsReg=nothing
				end if
			end if
			if Password<>"" then
				if strLength(Password)>12 or strLength(Password)<6 then
					founderr=true
					errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)���粻���޸ģ������գ�</li>"
				else
					if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
						errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
						founderr=true
					end if
				end if
			end if
			if Question="" then
				founderr=true
				errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
			end if
			if Sex="" then
				founderr=true
				errmsg=errmsg & "<br><li>�Ա���Ϊ��</li>"
			else
				sex=cint(sex)
				if Sex<>0 and Sex<>1 then
					Sex=1
				end if
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
			else
				if IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>����Email�д���</li>"
			   		founderr=true
				end if
			end if
			if LockUser="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û�״̬����Ϊ�գ�</li>"
			end if
			if FoundErr<>true then
				rsUser("UserName")=UserName
				if Password<>"" then
					rsUser("Password")=md5(Password)
				end if
				rsUser("Question")=Question
				if Answer<>"" then
					rsUser("Answer")=md5(Answer)
				end if
				rsUser("Sex")=Sex
				rsUser("Email")=Email
				rsUser("HomePage")=HomePage
				rsUser("CompanyName")=CompanyName
				rsUser("Add")=Add
				rsUser("Receiver")=Receiver
				rsUser("Postcode")=Postcode
				rsUser("Phone")=Phone
				rsUser("Mobile")=Mobile
				rsUser("Fax")=Fax
				rsUser("LockUser")=LockUser
				rsUser.update
				rsUser.Close
				set rsUser=nothing
				call CloseConn() 
				response.redirect "UserManage.asp"
			end if
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> 
      <FORM name="Form1" action="UserModify.asp" method="post">
        <table width="560" border="0" cellpadding="2" cellspacing="1" bgcolor="#000000" >
          <TR align=center bgcolor="#FFFFFF" class='title'> 
            <TD height=25 colSpan=2 class="back_southidc"><b>�޸�ע���û���Ϣ</b></TD>
          </TR>
          <TR bgcolor="#FFFFFF" > 
            <TD width="120" align="right">�� �� ����</TD>
            <TD> <INPUT name=UserName value="<%=rsUser("UserName")%>" size=30   maxLength=14></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">����(����6λ)��</TD>
            <TD> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������</font> 
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">�������⣺</TD>
            <TD> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30> 
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">����𰸣�</TD>
            <TD> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������</font></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">�Ա�</TD>
            <TD> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then response.write "CHECKED"%>>
              �� &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then response.write "CHECKED"%>>
              Ů</TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">Email��ַ��</TD>
            <TD> <INPUT name=Email value="<%=rsUser("Email")%>" size=30   maxLength=50> 
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">��ҳ��</TD>
            <TD> <INPUT   maxLength=100 size=30 name=HomePage value="<%=rsUser("HomePage")%>"></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">��˾���ƣ�</TD>
            <TD> <INPUT name=CompanyName value="<%=rsUser("CompanyName")%>" size=30 maxLength=20></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">�ջ���ַ��</TD>
            <TD> <INPUT name=Add value="<%=rsUser("Add")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD align="right">�ջ��ˣ�</TD>
            <TD> <INPUT name=Receiver value="<%=rsUser("Receiver")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD align="right">�������룺</TD>
            <TD> <INPUT name=Postcode value="<%=rsUser("Postcode")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD align="right">��ϵ�绰��<br></TD>
            <TD> <INPUT name=Phone value="<%=rsUser("Phone")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" >
            <TD align="right">�ֻ���</TD>
            <TD><INPUT name=Mobile value="<%=rsUser("Mobile")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD align="right">�� �棺</TD>
            <TD> <INPUT name=Fax value="<%=rsUser("Fax")%>" size=30 maxLength=50></TD>
          </TR>
          <TR bgcolor="#FFFFFF" class="tdbg" > 
            <TD width="120" align="right">�û�״̬��</TD>
            <TD> <input type="radio" name="LockUser" value="False" <%if rsUser("LockUser")=False then response.write "checked"%>>
              ����&nbsp;&nbsp; <input type="radio" name="LockUser" value="True" <%if rsUser("LockUser")=True then response.write "checked"%>>
              ����</TD>
          </TR>
          <TR align="center" bgcolor="#FFFFFF" class="tdbg" > 
            <TD height="40" colspan="2"> <input name="Action" type="hidden" id="Action" value="Modify"> 
              <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
          </TR>
        </TABLE>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
	end if
	rsUser.close
	set rsUser=nothing
end if
call CloseConn()
%>