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
<title>网站配置</title>
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
      <td height="24" colspan="4" class="back_southidc"> <div align="center"><strong>网 
          站 配 置</strong></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"> <strong>网站信息配置</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>网站名称：</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>网站标题：</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>网站描述：</strong></td>
	<td colspan="3" class="tdbg"><textarea name="SiteDescription" cols="50" rows="8" id="SiteDescription"><%=SiteDescription%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>网站关键字：</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteKeyword" type="text" id="SiteKeyword" value="<%=SiteKeyword%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>网站地址：</strong><br>
        请添写完整URL地址(用于网站统计)</td>
      <td colspan="3" class="tdbg"> <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>公司名称：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="CompanyName" type="text" id="CompanyName" value="<%=CompanyName%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>联系人：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="LinkMan" type="text" id="LinkMan" value="<%=LinkMan%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>移动电话：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Mobile" type="text" id="Mobile" value="<%=Mobile%>" size="40" maxlength="255" > 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>服务热线：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Tel" type="text" id="Tel" value="<%=Tel%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>QQ：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="QQ" type="text" id="QQ" value="<%=QQ%>" size="40" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>公司地址：</strong><br></td>
      <td class="tdbg" colspan="3"> <input name="Address" type="text" id="Address" value="<%=Address%>" size="40" maxlength="255">
      </td>
    </tr>
<tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>LOGO地址：</strong><br> </td>
      <td colspan="3" class="tdbg"> <input name="LogoUrl" type="text" id="LogoUrl" value="<%=LogoUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
<tr bgcolor="#FFFFFF"> 
      <td width="224" height="20" class="tdbg"><strong>首页Banner地址：</strong><br>
        图片格式为：jpg,gif,bmp,png,swf <br> </td>
      <td width="186" class="tdbg"> <input name="BannerUrl" type="text" id="BannerUrl" value="<%=BannerUrl%>" size="25" maxlength="255"> 
      </td>
      <td width="42" class="tdbg">宽度*高度：</td>
      <td width="127" class="tdbg">  <input name="width" type="text" id="width" value="<%=width%>" size="6" maxlength="5">*<input name="height" type="text" id="High" value="<%=height%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="23" class="tdbg"><strong>是否为FLASH：</strong></td>
      <td colspan="3" class="tdbg"> <input name="IsFlash" type="radio" value="Yes" <%if IsFlash="Yes" then response.write "checked"%>>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input name="IsFlash" type="radio" value="No" <%if IsFlash="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>企业邮局：</strong><br>
        请添写完整URL地址</td>
      <td colspan="3" class="tdbg"> <input name="EnterpriseMail" type="text" id="EnterpriseMail" value="<%=EnterpriseMail%>" size="40" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>站长姓名：</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterName" type="text" id="WebmasterName" value="<%=WebmasterName%>" size="40" maxlength="20"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>站长信箱：</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=WebmasterEmail%>" size="40" maxlength="100"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>版权信息：</strong><br>
        支持HTML标记，不能使用双引号</td>
      <td colspan="3" class="tdbg"><textarea name="Copyright" cols="50" rows="8" id="Copyright"><%=Copyright%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>网站选项配置</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>首页形象图片数：</strong></td>
      <td colspan="3" class="tdbg"> <input name="ComPic_count" type="text" id="ComPic_count" value="<%=ComPic_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>文章页每页文章数：</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxPerPage" type="text" id="MaxPerPage" value="<%=MaxPerPage%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>首页新闻资讯条数：</strong></td>
      <td colspan="3" class="tdbg"> <input name="New_count" type="text" id="New_count" value="<%=New_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>是否启用产品审核功能：</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableProductCheck" value="Yes" <%if EnableProductCheck="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableProductCheck" value="No" <%if EnableProductCheck="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>是否开放文件上传：</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableUploadFile" value="Yes" <%if EnableUploadFile="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableUploadFile" value="No" <%if EnableUploadFile="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>上传文件大小限制：</strong><br>
        建议不要超过1024K，以免影响服务器性能<strong>：</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>存放上传文件的目录：</strong><br>
        请输入相对于首页（Default.asp）的相对路径</td>
      <td colspan="3" class="tdbg"> <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>允许的上传文件类型：</strong><br>
        只输入扩展名。每种文件类型用"|"号分开。</td>
      <td colspan="3" class="tdbg"> <input name="UpFileType" type="text" id="UpFileType2" value="<%=UpFileType%>" size="50" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>删除文章时是否同时删除文章中的上传文件：</strong><br>
        此功能需要FSO支持。</td>
      <td colspan="3" class="tdbg"> <input type="radio" name="DelUpFiles" value="Yes" <%if DelUpFiles="Yes" then response.write "checked"%>>
        是 &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="DelUpFiles" value="No" <%if DelUpFiles="No" then response.write "checked"%>>
        否</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>Session会话的保持时间：</strong><br>
        主要用于后台管理员登录，为了安全，请不要将时间设得太长。建议设为10分钟</td>
      <td colspan="3" class="tdbg"> <input name="SessionTimeout" type="text" id="SessionTimeout" value="<%=SessionTimeout%>" size="6" maxlength="5">
        分钟</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg">&nbsp;</td>
      <td colspan="3" class="tdbg">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>邮件服务器选项</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>邮件发送组件：</strong><br>
        请一定要选择服务器上已安装的组件</td>
      <td colspan="3" class="tdbg"> <select name="MailObject" id="MailObject">
          <option value="Jmail" selected>Jmail</option>
        </select> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP服务器地址：</strong><br>
        用来发送邮件的SMTP服务器<br>
        如果你不清楚此参数含义，请联系你的空间商 </td>
      <td colspan="3" class="tdbg"> <input name="MailServer" type="text" id="MailServer" value="<%=MailServer%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP登录用户名：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数</td>
      <td colspan="3" class="tdbg"> <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=MailServerUserName%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP登录密码：</strong><br>
        当你的服务器需要SMTP身份验证时还需设置此参数 </td>
      <td colspan="3" class="tdbg"> <input name="MailServerPassWord" type="password" id="MailServerPassWord" value="<%=MailServerPassWord%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP域名</strong>：<br>
        如果用"name@domain.com"这样的用户名登录时，请指明domain.com</td>
      <td colspan="3" class="tdbg"> <input name="MailDomain" type="text" id="MailDomain" value="<%=MailDomain%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="SaveConfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " <% If ObjInstalled=false Then response.write "disabled" %>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能。<br>请直接修改“Inc/config.asp”文件中的内容。</font></b></td></tr>"
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
 Function InPutStr(Input)'向数据库中保存字符串时用 
     IF IsEmpty(Input) Then Input="" 
     IF IsNull(Input) Then Input="" 
     IF instr(input,chr(34))>0 Then input=replace(input,chr(34),chr(39))'将"替换成',以便在向数据库里存放 
     InPutStr=input 
End function 
sub SaveConfig()
	If ObjInstalled=false Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>你的服务器不支持 FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("../inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const EnterpriseMail=" & chr(34) & trim(request("EnterpriseMail")) & chr(34) & "        '企业邮局" & vbcrlf
	hf.write "Const LogoUrl=" & chr(34) & trim(request("LogoUrl")) & chr(34) & "        'Logo地址" & vbcrlf
	hf.write "Const BannerUrl=" & chr(34) & trim(request("BannerUrl")) & chr(34) & "        'Banner地址" & vbcrlf
	hf.write "Const IsFlash=" & chr(34) & trim(request("IsFlash")) & chr(34) & "        '是否为FLASH" & vbcrlf
	hf.write "Const width=" & chr(34) & trim(request("width")) & chr(34) & "        '宽度" & vbcrlf
	hf.write "Const height=" & chr(34) & trim(request("height")) & chr(34) & "        '高度" & vbcrlf
	hf.write "Const WebmasterName=" & chr(34) & trim(request("WebmasterName")) & chr(34) & "        '站长姓名" & vbcrlf
	hf.write "Const WebmasterEmail=" & chr(34) & trim(request("WebmasterEmail")) & chr(34) & "        '站长信箱" & vbcrlf
	hf.write "Const ComPic_count=" & chr(34) & InPutStr(trim(request("ComPic_count"))) & chr(34) & "        '首页形象图数" & vbcrlf	
	hf.write "Const New_count=" & trim(request("New_count")) &  "        '首面新闻资讯条数" & vbcrlf
	hf.write "Const MaxPerPage=" & trim(request("MaxPerPage")) & "        '文章搜索页每页文章数" & vbcrlf
	hf.write "Const EnableProductCheck=" & chr(34) & trim(request("EnableProductCheck")) & chr(34) & "        '是否启用文章审核功能" & vbcrlf
	hf.write "Const EnableUploadFile=" & chr(34) & trim(request("EnableUploadFile")) & chr(34) & "        '是否开放文件上传" & vbcrlf
	
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '上传文件大小限制" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '存放上传文件的目录" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '允许的上传文件类型" & vbcrlf
	hf.write "Const DelUpFiles=" & chr(34) & trim(request("DelUpFiles")) & chr(34) & "        '删除文章时是否同时删除文章中的上传文件" & vbcrlf
	hf.write "Const SessionTimeout=" & trim(request("SessionTimeout")) & "        'Session会话的保持时间" & vbcrlf
	
	hf.write "Const MailObject=" & chr(34) & trim(request("MailObject")) & chr(34) & "        '邮件发送组件" & vbcrlf
	hf.write "Const MailServer=" & chr(34) & trim(request("MailServer")) & chr(34) & "        '用来发送邮件的SMTP服务器" & vbcrlf
	hf.write "Const MailServerUserName=" & chr(34) & trim(request("MailServerUserName")) & chr(34) & "        '登录用户名" & vbcrlf
	hf.write "Const MailServerPassWord=" & chr(34) & trim(request("MailServerPassWord")) & chr(34) & "        '登录密码" & vbcrlf
	hf.write "Const MailDomain=" & chr(34) & trim(request("MailDomain")) & chr(34) & "        '域名" & vbcrlf
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