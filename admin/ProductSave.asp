<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim rs,sql,ErrMsg,FoundErr
dim ID,Product_Id,BigClassName,EnBigClassName,SmallClassName,EnSmallClassName 
dim Title,EnTitle,Spec,EnSpec,Unit,EnUnit,Memo,EnMemo,Content,EnContent
dim UpdateTime,Hits,DefaultPicUrl,Elite,Passed
FoundErr=false
ID=Trim(Request.Form("ID"))
Product_Id=trim(request.form("Product_Id"))
BigClassName=trim(request.form("BigClassName"))
SmallClassName=trim(request.form("SmallClassName"))
Title=trim(request.form("Title"))
Content=trim(request.form("Content"))
UpdateTime=trim(request.form("UpdateTime"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
Passed=trim(request.form("Passed"))
Elite=trim(request.form("Elite"))
Newproduct=trim(request.form("Newproduct"))
Hits=trim(request.form("Hits"))

sqlBig="select * from BigClass where BigClassName='" & BigClassName & "'"
Set rsBig= Server.CreateObject("ADODB.Recordset")
rsBig.open sqlBig,conn,1,1
EnBigClassName=rsBig("EnBigClassName")
rsBig.close

if SmallClassName<>"" then
sqlSmall="select * from SmallClass where SmallClassName='" & SmallClassName & "' and  BigClassName='" & BigClassName & "'"
Set rsSmall= Server.CreateObject("ADODB.Recordset")
rsSmall.open sqlSmall,conn,1,1
EnSmallClassName=rsSmall("EnSmallClassName")
rsSmall.close
end if


if BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>未指定产品所属大类</li>"
end if
if Title="" then
	founderr=true
	errmsg="<li>产品标题不能为空</li>"
end if


if founderr=false then	
	if UpdateTime<>"" and IsDate(UpdateTime)=true then
		UpdateTime=CDate(UpdateTime)
	else
		UpdateTime=now()
	end if
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select * from Product where (ID is null)"
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()		
		rs.update
		ID=rs("ID")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then	  
  		if ID<>"" then
			sql="select * from Product where id=" & ID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到此产品，可能已经被删除。</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>不能确定产品ID的值</li>"
			call WriteErrMsg()
		end if
	else
		founderr=true
		errmsg=errmsg+"<li>没有选定参数</li>"
		call WriteErrMsg()
	end if

	call CloseConn()
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> <br> <br> <b> </b><br> <br> <br> 
      <table class="border" align=center width="70%" border="0" cellpadding="0" cellspacing="2" bordercolor="#999999">
        <tr align=center bgcolor="#A4B6D7"> 
          <td  height="25" colspan="2" class="title"><b> 
            <%if request("action")="add" then%>
            添加 
            <%else%>
            修改 
            <%end if%>
            产品成功</b></td>
        </tr>
        <tr> 
          <td width="24%" height="22" bgcolor="#C0C0C0" class="tdbg"> <p align="right">产品序号：</p></td>
          <td width="76%" bgcolor="#E3E3E3" class="tdbg"><%=ID%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品编号：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Product_Id%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">产品名称：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"><%=Title%></td>
        </tr>
        <tr> 
          <td height="22" bgcolor="#C0C0C0" class="tdbg"><div align="right">所属类别：</div></td>
          <td bgcolor="#E3E3E3" class="tdbg"> <%response.write BigClassName
		if SmallClassName<>"" then response.write " &gt;&gt; " & SmallClassName		
		 %> </td>
        </tr>
        <tr> 
          <td height="22" colspan="2" bgcolor="#E3E3E3" class="tdbg"> <p align="center"> 
              【<a href="ProductModify.asp?ID=<%=ID%>">修改产品</a>】&nbsp;【<a href="ProductAdd.asp">继续添加产品</a>】&nbsp;【<a href="ProductManage.asp">产品管理</a>】&nbsp;【<a href="../ProductShow.asp?ID=<%=ID%>" target="_blank">预览产品内容</a>】</p></td>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
else
	WriteErrMsg
end if

sub SaveData()
	rs("Product_Id")=Product_Id
	rs("BigClassName")=BigClassName
	rs("SmallClassName")=SmallClassName
	rs("Title")=Title
	rs("Content")=Content
	rs("Hits")=Hits
	if Passed="yes" then
		rs("Passed")=True
	else
		if EnableProductCheck="No" then
			rs("Passed")=True
		else
			rs("Passed")=False
		end if
	end if
	if Elite="yes" then
		rs("Elite")=True
	else
		rs("Elite")=False
	end if
	if Newproduct="yes" then
		rs("Newproduct")=True
	else
		rs("Newproduct")=False
	end if
	rs("UpdateTime")=UpdateTime	
	rs("DefaultPicUrl")=DefaultPicUrl
	if DefaultPicUrl<>"" then
	rs("IncludePic")= true	
End If
end sub	
%>