<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<!--#include file="../inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim ObjInstalled,Action,FoundErr,ErrMsg
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowInfo"
end if
	set rs = server.CreateObject("adodb.recordset")
	sql = "select * from siteconfig"
	rs.open sql,conn,1,1
	if not rs.eof then
		SiteName=rs("SiteName")
		SiteTitle=rs("SiteTitle")
		SiteDescription=rs("SiteDescription")
		SiteKeyword=rs("SiteKeyword")
		SiteUrl=rs("SiteUrl")
		LinkMan=rs("LinkMan")
		CompanyName=rs("CompanyName")
		Mobile=rs("Mobile")
		Tel=rs("Tel")
		QQ=rs("QQ")
		Address=rs("Address")
		Copyright=rs("Copyright")
	end if
	rs.close
	set rs=nothing
%>
<html>
<head>
<title>��վ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0"> 
<%
if Action="SaveConfig" then
	call SaveConfig()
	Response.Redirect "SiteConfig.asp" 
else
	call ShowConfig()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub ShowConfig()
%>
<form method="POST" action="SiteConfig.asp" id="form" name="form">
  <table width="650" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" >
    <tr> 
      <td height="24" colspan="4" class="back_southidc"> <div align="center"><strong>�� 
          վ �� ��</strong></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"> <strong>��վ��Ϣ����</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ���ƣ�</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ���⣺</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ������</strong></td>
	<td colspan="3" class="tdbg"><textarea name="SiteDescription" cols="50" rows="8" id="SiteDescription"><%=SiteDescription%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ�ؼ��֣�</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteKeyword" type="text" id="SiteKeyword" value="<%=SiteKeyword%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ��ַ��</strong><br>
        ����д����URL��ַ(������վͳ��)</td>
      <td colspan="3" class="tdbg"> <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��˾���ƣ�</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="CompanyName" type="text" id="CompanyName" value="<%=CompanyName%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��ϵ�ˣ�</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="LinkMan" type="text" id="LinkMan" value="<%=LinkMan%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>�ƶ��绰��</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Mobile" type="text" id="Mobile" value="<%=Mobile%>" size="40" maxlength="255" > 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>�������ߣ�</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Tel" type="text" id="Tel" value="<%=Tel%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>QQ��</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="QQ" type="text" id="QQ" value="<%=QQ%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��˾��ַ��</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Address" type="text" id="Address" value="<%=Address%>" size="40" maxlength="255">
      </td>
    </tr>
<tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>LOGO��ַ��</strong><br> </td>
      <td colspan="3" class="tdbg"> <input name="LogoUrl" type="text" id="LogoUrl" value="<%=LogoUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
<tr bgcolor="#FFFFFF"> 
      <td width="224" height="20" class="tdbg"><strong>��ҳBanner��ַ��</strong><br>
        ͼƬ��ʽΪ��jpg,gif,bmp,png,swf <br> </td>
      <td width="186" class="tdbg"> <input name="BannerUrl" type="text" id="BannerUrl" value="<%=BannerUrl%>" size="25" maxlength="255"> 
      </td>
      <td width="42" class="tdbg">����*�߶ȣ�</td>
      <td width="127" class="tdbg">  <input name="width" type="text" id="width" value="<%=width%>" size="6" maxlength="5">*<input name="height" type="text" id="High" value="<%=height%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="23" class="tdbg"><strong>�Ƿ�ΪFLASH��</strong></td>
      <td colspan="3" class="tdbg"> <input name="IsFlash" type="radio" value="Yes" <%if IsFlash="Yes" then response.write "checked"%>>
        ��&nbsp;&nbsp;&nbsp;&nbsp; <input name="IsFlash" type="radio" value="No" <%if IsFlash="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҵ�ʾ֣�</strong><br>
        ����д����URL��ַ</td>
      <td colspan="3" class="tdbg"> <input name="EnterpriseMail" type="text" id="EnterpriseMail" value="<%=EnterpriseMail%>" size="40" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>վ��������</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterName" type="text" id="WebmasterName" value="<%=WebmasterName%>" size="40" maxlength="20"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>վ�����䣺</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=WebmasterEmail%>" size="40" maxlength="100"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��Ȩ��Ϣ��</strong><br>
        ֧��HTML��ǣ�����ʹ��˫����</td>
      <td colspan="3" class="tdbg"><textarea name="Copyright" cols="50" rows="8" id="Copyright"><%=Copyright%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>��վѡ������</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҳ����ͼƬ����</strong></td>
      <td colspan="3" class="tdbg"> <input name="ComPic_count" type="text" id="ComPic_count" value="<%=ComPic_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>����ҳÿҳ��������</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxPerPage" type="text" id="MaxPerPage" value="<%=MaxPerPage%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҳ������Ѷ������</strong></td>
      <td colspan="3" class="tdbg"> <input name="New_count" type="text" id="New_count" value="<%=New_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�Ƿ����ò�Ʒ��˹��ܣ�</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableProductCheck" value="Yes" <%if EnableProductCheck="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableProductCheck" value="No" <%if EnableProductCheck="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�Ƿ񿪷��ļ��ϴ���</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableUploadFile" value="Yes" <%if EnableUploadFile="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableUploadFile" value="No" <%if EnableUploadFile="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�ϴ��ļ���С���ƣ�</strong><br>
        ���鲻Ҫ����1024K������Ӱ�����������<strong>��</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>����ϴ��ļ���Ŀ¼��</strong><br>
        �������������ҳ��Default.asp�������·��</td>
      <td colspan="3" class="tdbg"> <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�������ϴ��ļ����ͣ�</strong><br>
        ֻ������չ����ÿ���ļ�������"|"�ŷֿ���</td>
      <td colspan="3" class="tdbg"> <input name="UpFileType" type="text" id="UpFileType2" value="<%=UpFileType%>" size="50" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ���</strong><br>
        �˹�����ҪFSO֧�֡�</td>
      <td colspan="3" class="tdbg"> <input type="radio" name="DelUpFiles" value="Yes" <%if DelUpFiles="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="DelUpFiles" value="No" <%if DelUpFiles="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>Session�Ự�ı���ʱ�䣺</strong><br>
        ��Ҫ���ں�̨����Ա��¼��Ϊ�˰�ȫ���벻Ҫ��ʱ�����̫����������Ϊ10����</td>
      <td colspan="3" class="tdbg"> <input name="SessionTimeout" type="text" id="SessionTimeout" value="<%=SessionTimeout%>" size="6" maxlength="5">
        ����</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg">&nbsp;</td>
      <td colspan="3" class="tdbg">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>�ʼ�������ѡ��</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>�ʼ����������</strong><br>
        ��һ��Ҫѡ����������Ѱ�װ�����</td>
      <td colspan="3" class="tdbg"> <select name="MailObject" id="MailObject">
          <option value="Jmail" selected>Jmail</option>
        </select> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��������ַ��</strong><br>
        ���������ʼ���SMTP������<br>
        ����㲻����˲������壬����ϵ��Ŀռ��� </td>
      <td colspan="3" class="tdbg"> <input name="MailServer" type="text" id="MailServer" value="<%=MailServer%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��¼�û�����</strong><br>
        ����ķ�������ҪSMTP������֤ʱ�������ô˲���</td>
      <td colspan="3" class="tdbg"> <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=MailServerUserName%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��¼���룺</strong><br>
        ����ķ�������ҪSMTP������֤ʱ�������ô˲��� </td>
      <td colspan="3" class="tdbg"> <input name="MailServerPassWord" type="password" id="MailServerPassWord" value="<%=MailServerPassWord%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP����</strong>��<br>
        �����"name@domain.com"�������û�����¼ʱ����ָ��domain.com</td>
      <td colspan="3" class="tdbg"> <input name="MailDomain" type="text" id="MailDomain" value="<%=MailDomain%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="SaveConfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " <% If ObjInstalled=false Then response.write "disabled" %>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ����ܡ�<br>��ֱ���޸ġ�Inc/config.asp���ļ��е����ݡ�</font></b></td></tr>"
End If
%>
  </table>
<%
end sub
%>
</form>
</body>
</html>
<%
 Function InPutStr(Input)'�����ݿ��б����ַ���ʱ�� 
     IF IsEmpty(Input) Then Input="" 
     IF IsNull(Input) Then Input="" 
     IF instr(input,chr(34))>0 Then input=replace(input,chr(34),chr(39))'��"�滻��',�Ա��������ݿ����� 
     InPutStr=input 
End function 
sub SaveConfig()
	If ObjInstalled=false Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ķ�������֧�� FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("../inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const EnterpriseMail=" & chr(34) & trim(request("EnterpriseMail")) & chr(34) & "        '��ҵ�ʾ�" & vbcrlf
	hf.write "Const LogoUrl=" & chr(34) & trim(request("LogoUrl")) & chr(34) & "        'Logo��ַ" & vbcrlf
	hf.write "Const BannerUrl=" & chr(34) & trim(request("BannerUrl")) & chr(34) & "        'Banner��ַ" & vbcrlf
	hf.write "Const IsFlash=" & chr(34) & trim(request("IsFlash")) & chr(34) & "        '�Ƿ�ΪFLASH" & vbcrlf
	hf.write "Const width=" & chr(34) & trim(request("width")) & chr(34) & "        '����" & vbcrlf
	hf.write "Const height=" & chr(34) & trim(request("height")) & chr(34) & "        '�߶�" & vbcrlf
	hf.write "Const WebmasterName=" & chr(34) & trim(request("WebmasterName")) & chr(34) & "        'վ������" & vbcrlf
	hf.write "Const WebmasterEmail=" & chr(34) & trim(request("WebmasterEmail")) & chr(34) & "        'վ������" & vbcrlf
	hf.write "Const ComPic_count=" & chr(34) & InPutStr(trim(request("ComPic_count"))) & chr(34) & "        '��ҳ����ͼ��" & vbcrlf	
	hf.write "Const New_count=" & trim(request("New_count")) &  "        '����������Ѷ����" & vbcrlf
	hf.write "Const MaxPerPage=" & trim(request("MaxPerPage")) & "        '��������ҳÿҳ������" & vbcrlf
	hf.write "Const EnableProductCheck=" & chr(34) & trim(request("EnableProductCheck")) & chr(34) & "        '�Ƿ�����������˹���" & vbcrlf
	hf.write "Const EnableUploadFile=" & chr(34) & trim(request("EnableUploadFile")) & chr(34) & "        '�Ƿ񿪷��ļ��ϴ�" & vbcrlf
	
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '�ϴ��ļ���С����" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '����ϴ��ļ���Ŀ¼" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '�������ϴ��ļ�����" & vbcrlf
	hf.write "Const DelUpFiles=" & chr(34) & trim(request("DelUpFiles")) & chr(34) & "        'ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ�" & vbcrlf
	hf.write "Const SessionTimeout=" & trim(request("SessionTimeout")) & "        'Session�Ự�ı���ʱ��" & vbcrlf
	
	hf.write "Const MailObject=" & chr(34) & trim(request("MailObject")) & chr(34) & "        '�ʼ��������" & vbcrlf
	hf.write "Const MailServer=" & chr(34) & trim(request("MailServer")) & chr(34) & "        '���������ʼ���SMTP������" & vbcrlf
	hf.write "Const MailServerUserName=" & chr(34) & trim(request("MailServerUserName")) & chr(34) & "        '��¼�û���" & vbcrlf
	hf.write "Const MailServerPassWord=" & chr(34) & trim(request("MailServerPassWord")) & chr(34) & "        '��¼����" & vbcrlf
	hf.write "Const MailDomain=" & chr(34) & trim(request("MailDomain")) & chr(34) & "        '����" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing	
	set rs = server.CreateObject("adodb.recordset")
	sql = "select * from siteconfig"
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("SiteName") = request("SiteName")
		rs("SiteTitle") = request("SiteTitle")
		rs("SiteDescription") = request("SiteDescription")
		rs("SiteKeyword") = request("SiteKeyword")
		rs("SiteUrl") = request("SiteUrl")
		rs("LinkMan") = request("LinkMan")
		rs("CompanyName") = request("CompanyName")
		rs("Mobile") = request("Mobile")
		rs("Tel") = request("Tel")
		rs("QQ") = request("QQ")
		rs("Address") = request("Address")
		rs("Copyright") = request("Copyright")
		rs.update()
	end if
	rs.close
	set rs=nothing
end sub
%>