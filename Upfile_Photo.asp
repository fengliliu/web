<%@language=vbscript codepage=936 %>
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upfile_class.asp"-->
<%
const upload_type=0   '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload

dim upload,oFile,formName,SavePath,filename,fileExt,oFileSize
dim EnableUpload
dim arrUpFileType
dim ranNum
dim msg,FoundErr
dim PhotoUrlID
msg=""
FoundErr=false
EnableUpload=false
dim PageSource
	PageSource = request("CKEditor")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY{
BACKGROUND-COLOR: #E1F4EE;
font-size:9pt
}
.tx1 { height: 20px;font-size: 9pt; border: 1px solid; border-color: #000000; color: #0000FF}
-->
</style>
</head>
<body leftmargin="2" topmargin="5" marginwidth="0" marginheight="0" >
<%
if EnableUploadFile="No" then
	response.write "ϵͳδ�����ļ��ϴ�����"
else
	if session("AdminName")="" then
		response.Write("���¼����ʹ�ñ����ܣ�")
	else
		select case upload_type
			case 0
				call upload_0()  'ʹ�û���������ϴ���
			case else
				'response.write "��ϵͳδ���Ų������"
				'response.end
		end select
	end if
end if
%>
</body>
</html>
<%
sub upload_0()    'ʹ�û���������ϴ���
	set upload=new upfile_class ''�����ϴ�����
	upload.GetData(104857600)   'ȡ���ϴ�����,��������ϴ�100M
	if upload.err > 0 then  '�������
		select case upload.err
			case 1
				response.write "����ѡ����Ҫ�ϴ����ļ���"
			case 2
				response.write "���ϴ����ļ��ܴ�С������������ƣ�100M��"
		end select
		response.end
	end if
	PhotoUrlID = trim(upload.form("PhotoUrlID"))
	if PhotoUrlID ="" Then PhotoUrlID =0
	PhotoUrlID=Clng(PhotoUrlID)
	if PhotoUrlID>0 then
		SavePath = SaveUpFilesPath   '����ϴ��ļ���Ŀ¼
	else
		SavePath = SaveUpFilesPath   '����ϴ��ļ���Ŀ¼
	end if
	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
		
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set ofile=upload.file(formName)  '����һ���ļ�����
		oFileSize=ofile.filesize
		if oFileSize<100 then
			msg="����ѡ����Ҫ�ϴ����ļ���"
			FoundErr=True
		else
		    if oFileSize>(10000*1024) then
         	 msg="�ļ���С���������ƣ����ֻ���ϴ�10M���ļ���"
			 FoundErr=true
		    end if	
		end if

		fileExt=lcase(ofile.FileExt)
		arrUpFileType=split(UpFileType,"|")
		for i=0 to ubound(arrUpFileType)
			if fileEXT=trim(arrUpFileType(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" or fileEXT="php" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			FoundErr=true
		end if
		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			ofile.SaveToFile Server.mappath(FileName)   '�����ļ�

			response.write "�ļ��ϴ��ɹ����ļ���СΪ��" & cstr(round(oFileSize/1024)) & "K"
			if PageSource<>"Content" Then
				select case PhotoUrlID
				case "1"
				strJS=strJS & "parent.document.myform.DefaultPicUrl.value='" & fileName & "';" & vbcrlf
				case "2"
					strJS=strJS & "parent.document.myform.CompHonor.value='" & fileName & "';" & vbcrlf		
				case "3"
					strJS=strJS & "parent.document.myform.CompVisualize.value='" & fileName & "';" & vbcrlf						
				case "6"
				strJS=strJS & "parent.document.myform.DefaultPicUrl.value='" & fileName & "';" & vbcrlf
				case "7"
					strJS=strJS & "parent.document.myform.MovieAddr.value='" & fileName & "';" & vbcrlf
				end select 
			else
				strJS=strJS & "var img=new Image();img.src='/" & fileName & "';" & vbcrlf
				strJS=strJS & "imgwidth=img.width;imgheight=img.height;" & vbcrlf
				strJS=strJS & "parent.document.getElementById('cke_94_textInput').value='/"&fileName&"';" & vbcrlf
				strJS=strJS & "parent.document.getElementById('cke_90_previewImage').style.display='block';" & vbcrlf
				strJS=strJS & "parent.document.getElementById('cke_90_previewImage').src='/"&fileName&"';" & vbcrlf
				strJS=strJS & "parent.document.getElementById('cke_104_textInput').value=600;" & vbcrlf
				strJS=strJS & "parent.document.getElementById('cke_110_textInput').value=400;" & vbcrlf
				
			End IF
		else
			strJS=strJS & "alert('" & msg & "');" & vbcrlf
		  	strJS=strJS & "history.go(-1);" & vbcrlf
		end if
		strJS=strJS & "</script>" & vbcrlf
		response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>
