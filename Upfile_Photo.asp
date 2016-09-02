<%@language=vbscript codepage=936 %>
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upfile_class.asp"-->
<%
const upload_type=0   '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

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
	response.write "系统未开放文件上传功能"
else
	if session("AdminName")="" then
		response.Write("请登录后再使用本功能！")
	else
		select case upload_type
			case 0
				call upload_0()  '使用化境无组件上传类
			case else
				'response.write "本系统未开放插件功能"
				'response.end
		end select
	end if
end if
%>
</body>
</html>
<%
sub upload_0()    '使用化境无组件上传类
	set upload=new upfile_class ''建立上传对象
	upload.GetData(104857600)   '取得上传数据,限制最大上传100M
	if upload.err > 0 then  '如果出错
		select case upload.err
			case 1
				response.write "请先选择你要上传的文件！"
			case 2
				response.write "你上传的文件总大小超出了最大限制（100M）"
		end select
		response.end
	end if
	PhotoUrlID = trim(upload.form("PhotoUrlID"))
	if PhotoUrlID ="" Then PhotoUrlID =0
	PhotoUrlID=Clng(PhotoUrlID)
	if PhotoUrlID>0 then
		SavePath = SaveUpFilesPath   '存放上传文件的目录
	else
		SavePath = SaveUpFilesPath   '存放上传文件的目录
	end if
	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
		
	for each formName in upload.file '列出所有上传了的文件
		set ofile=upload.file(formName)  '生成一个文件对象
		oFileSize=ofile.filesize
		if oFileSize<100 then
			msg="请先选择你要上传的文件！"
			FoundErr=True
		else
		    if oFileSize>(10000*1024) then
         	 msg="文件大小超过了限制，最大只能上传10M的文件！"
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
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
			FoundErr=true
		end if
		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			ofile.SaveToFile Server.mappath(FileName)   '保存文件

			response.write "文件上传成功！文件大小为：" & cstr(round(oFileSize/1024)) & "K"
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
