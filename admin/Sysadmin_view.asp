<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<%
'=================================
'
'    SOUTHIDC.NET(�Ϸ�����)
'    QQ:635495 MSN:jxhwq@hotmail.com
'    Email:hdz2008@163.com  
'    http://www.southidc.net
'    copyright(c)2004 �ƴ�����
'
'=================================
%>
<%

Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"    
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
%>
<%
set rs = Server.CreateObject("ADODB.Recordset")
sqltext="select * from Admin where username='"+Session("AdminName")+"'"
rs.Open sqltext,conn,1,1
if not rs.EOF then	
	LoginTimes=rs("LoginTimes")	
	LastLoginTime=rs("LastLoginTime")
	idcount=rs(0)	
end if
rs.Close
%>
<!--#include file="inc/head.asp"-->
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
  <tr> 
    <td class="back_southidc" colspan="2" height="25" align="center"><b>������ݷ�ʽ</b></td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="20%" height="23">��ݹ�������</td>
    <td width="80%" height="23"> 
      | <a href="Admin_Manage.asp"><font color="000000">����Ա����</font></a> </td>
  </tr>
</table>
<br>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
  <tr> 
    <td class="back_southidc" colspan="2" height="25" align="center"><b>ϵͳ��Ϣ</b></td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="48%" height="23">�û�����<font class="t4"> <%=Session("AdminName")%></font></td>
    <td width="52%">IP��<font class="t4"> <%=Request.ServerVariables("REMOTE_ADDR")%></font></td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="48%" height="23">���ݹ��ڣ�<font class="t4"><%=Session.timeout%> ����</font></td>
    <td width="52%">����ʱ�䣺<font class="t4"> <%=year(now())%>��<%=month(now())%>��<%=day(now())%>��<%=hour(now())%>:<%=minute(now())%></font></td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="48%" height="23">���ߴ����� <font class="t4"><%=LoginTimes%></font></td>
    <td width="52%">����ʱ�䣺<font class="t4"> <%=LastLoginTime%></font></td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="48%" height="23">������������<font class="t4"> <%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("Http_HOST")%></font></td>
    <td width="52%">�ű��������棺<font class="t4"> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></font></td>
  </tr>
  <tr class="tr_southidc"> 
    <td height="23">���������������ƣ�<font class="t4"> <%=Request.ServerVariables("SERVER_SOFTWARE")%></font></td>
    <td>������汾��<font class="t4"> <%=Request.ServerVariables("Http_User_Agent")%></font></td>
  </tr>
  <tr class="tr_southidc">
    <td height="23">FSO�ı���д�� 
      <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <font color="red"><b>��</b></font> 
      <%else%>
      <b>��</b> 
      <%end if%>
    </td>
    <td>���ݿ�ʹ�ã� 
      <%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <font color="red"><b>��</b></font> 
      <%else%>
      <b>��</b> 
      <%end if%>
    </td>
  </tr>
  <tr class="tr_southidc"> 
    <td width="48%" height="23">Jmail���֧�֣� 
      <%If Not IsObjInstalled(theInstalledObjects(13)) Then%> <font color="red"><b>��</b></font> <%else%> <b>��</b> <%end if%> </td>
    <td width="52%">CDONTS���֧�֣� 
      <%If Not IsObjInstalled(theInstalledObjects(14)) Then%> <font color="red"><b>��</b></font> <%else%> <b>��</b> <%end if%> </td>
    <!-- <td width="50%">ACCESS ���ݿ�·����<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td> -->
  </tr>
</table>
<br>
<%
sub discreteness
%>
<table border=0 width="95%" cellspacing=1 cellpadding=3 class=table_southidc align=center>
<tr class=back_southidc>
<td width="57%" height="25">&nbsp;�������</td><td width="41%" height="25">֧�ּ��汾</td>
</tr>
<%
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "MSWC.PermissionChecker"
theInstalledObjects(7) = "ADODB.Stream"
theInstalledObjects(8) = "adodb.connection"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "SoftArtisans.FileUp"
theInstalledObjects(11) = "SoftArtisans.FileManager"
theInstalledObjects(12) = "JMail.Message"
theInstalledObjects(13) = "CDONTS.NewMail"
theInstalledObjects(14) = "Persits.MailSender"
theInstalledObjects(15) = "LyfUpload.UploadFile"
theInstalledObjects(16) = "Persits.Upload.1"
theInstalledObjects(17) = "w3.upload"
For i=0 to 17
Response.Write "<TR class=td_southidc><TD>&nbsp;" & theInstalledObjects(i) & "<font color=888888>&nbsp;"
select case i
case 8
Response.Write "(ACCESS ���ݿ�)"
case 9
Response.Write "(FSO �ı��ļ���д)"
case 10
Response.Write "(SA-FileUp �ļ��ϴ�)"
case 11
Response.Write "(SA-FM �ļ�����)"
case 12
Response.Write "(JMail �ʼ�����)"
case 13
Response.Write "(WIN����SMTP ����)"
case 14
Response.Write "(ASPEmail �ʼ�����)"
case 15
Response.Write "(LyfUpload �ļ��ϴ�)"
case 16
Response.Write "(ASPUpload �ļ��ϴ�)"
case 17
Response.Write "(w3 upload �ļ��ϴ�)"
end select
Response.Write "</font></td><td height=25>"
If Not IsObjInstalled(theInstalledObjects(i)) Then
Response.Write "<font color=red><b>��</b></font>"
Else
Response.Write "<b>��</b> " & getver(theInstalledObjects(i)) & ""
End If
Response.Write "</td></TR>" & vbCrLf
Next
%>
</table>
<%
end sub

''''''''''''''''''''''''''''''
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
''''''''''''''''''''''''''''''
Function getver(Classstr)
On Error Resume Next
getver=""
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(Classstr)
If 0 = Err Then getver=xtestobj.version
Set xTestObj = Nothing
Err = 0
End Function

%>
<%htmlend%>
<%
sub htmlend
%><p>
<table cellspacing=0 cellpadding=0 width=95% align=center><tr>
    <td align=middle> Copyright 2013 <a target=_blank href="http://www.pxsww.net"><font color=000000>http://www.pxsww.net</font></a> 
      <a href="http://www.southidc.cn" target="_blank">southidc.cn</a><br>
      Powered by <font color=ffffff> <a target=_blank href="http://www.pxsww.net"><font color=000000>�������� 
      Ver9.0</font></a></font>&nbsp;&copy; 2013<br>
Script Execution Time:<%=fix((timer()-startime)*1000)%>ms
</td></tr></table>
<%
end sub
%>