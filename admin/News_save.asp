<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="../Inc/Config.asp"-->
<!--#include file="../Inc/Ubbcode.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
ID=Request.Form("ID")
Title=Trim(request.form("Title"))
BigClassName=trim(request.form("BigClassName"))
SmallClassName=trim(request.form("SmallClassName"))
Content=trim(request.form("Content"))
price=trim(request.form("price"))
User=Trim(request.form("User"))
FirstImageName=trim(request.form("DefaultPicUrl"))
Ok=Trim(request.form("Ok"))
AddDate=trim(request.form("AddDate"))
Hits=trim(request.form("Hits"))
Action=trim(request("Action"))
if BigClassName="" then
	founderr=true
	errmsg=errmsg+"<li>δָ����Ϣ��������</li>"
end if
if Title="" then
	founderr=true
	errmsg="<li>��Ϣ���ⲻ��Ϊ��</li>"
end if
if Content="" then
	founderr=true
	errmsg=errmsg+"<li>��Ϣ���ݲ���Ϊ��</li>"
end if

if founderr=false then
	Title=dvhtmlencode(Title)	
	if AddDate<>"" and IsDate(AddDate)=true then
		AddDate=CDate(AddDate)
	else
		AddDate=now()
	end if
	
	if Hits<>"" then
		Hits=CLng(Hits)
	else
		Hits=0
	end if
	
	set rs=server.createobject("adodb.recordset")
	select case Action
	  case "Add"	
		sql="select * from News where (id is null)" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()		
		rs.update	
		rs.close
		set rs=nothing
		response.redirect "News_Manage.asp"
	 case "Modify"	
  		if ID<>"" then
			sql="select * from News where id="&ID
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
			    
				call SaveData()				
				rs.update
				rs.close
				set rs=nothing
	            response.redirect "News_Manage.asp"
 			else
				founderr=true
				errmsg=errmsg+"<li>�Ҳ�������Ϣ�������Ѿ���ɾ����</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>����ȷ��ID��ֵ</li>"
			call WriteErrMsg()
		end if
	 Case else
		founderr=true
		errmsg=errmsg+"<li>û��ѡ������</li>"
		call WriteErrMsg()
	end select
	call CloseConn()
else
	WriteErrMsg
end if
%>

<%
sub SaveData()  
  rs("Title")=Title
  rs("Content")=Content
  rs("price")=price 
  rs("SmallContent")=left(loseHtml(Content),100)
  rs("User")=User
  rs("BigClassName")=BigClassName
  rs("SmallClassName")=SmallClassName
  rs("AddDate")=AddDate
  rs("Hits")=Hits
  if Ok="Yes" then
	rs("Ok")=True
  else
	rs("Ok")=False
  end if 
   
  '***************************************
	'ɾ�����õ��ϴ��ļ�
	if ObjInstalled=True and UploadFiles<>"" then
		dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(UploadFiles,"|")>1 then
			dim arrUploadFiles,intTemp
			arrUploadFiles=split(UploadFiles,"|")
			UploadFiles=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Content,arrUploadFiles(intTemp))<=0 and arrUploadFiles(intTemp)<>FirstImageName then
					strRubbishFile=server.MapPath("../" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "����Ϣ��û���õ���Ҳû�б���Ϊ��ҳͼƬ�������Ѿ���ɾ����</li>"
					end if
				else
					if intTemp=0 then
						UploadFiles=arrUploadFiles(intTemp)
					else
						UploadFiles=UploadFiles & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Content,UploadFiles)<=0 and UploadFiles<>FirstImageName then
				strRubbishFile=server.MapPath("../" & UploadFiles)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & UploadFiles & "����Ϣ��û���õ���Ҳû�б���Ϊ��ҳͼƬ�������Ѿ���ɾ����</li>"
				end if
				UploadFiles=""
			end if
		end if
		set fso=nothing
	end If
	'����
	'***************************************
  rs("FirstImageName")=FirstImageName     
end sub
%>
