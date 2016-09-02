<%

'================================================
'函数名：FormatDate
'作　用：格式化日期
'参　数：DateAndTime            (原日期和时间)
'       Format                 (新日期格式)
'返回值：格式化后的日期
'================================================
Function FormatDate(DateAndTime, Format)
  On Error Resume Next
  Dim yy,y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(Format) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  yy = CStr(Year(DateAndTime))
  y = Mid(CStr(Year(DateAndTime)),3)
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
   
  Select Case Format
  Case "1"
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
    strDateTime = yy & m & d & h & mi & s
    '返回12位 直到秒 的时间字符串
  Case "3"
    strDateTime = yy & m & d & h & mi    
    '返回12位 直到分 的时间字符串
  Case "4"
    strDateTime = yy & "年" & m & "月" & d & "日"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "月" & d & "日"
  Case "8"
    strDateTime = y & "年" & m & "月"
  Case "9"
    strDateTime = y & "-" & m
  Case "10"
    strDateTime = y & "/" & m
  Case "11"
    strDateTime = y & "-" & m & "-" & d
  Case "12"
    strDateTime = y & "/" & m & "/" & d
  Case "13"
    strDateTime = yy & "." & m & "." & d
  Case "14"
    strDateTime = yy & "-" & m & "-" & d
  Case "15"
    strDateTime = m & "-" & d & "&nbsp;" & h & ":" & mi
    '返回12位 直到分 的时间字符串
  Case Else
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function
'==================================================
'过程名：MainMenu
'作  用：显示主导航
'参  数: 
'==================================================
sub MainMenu()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select ClassID,ClassName,Depth,NextID,LinkUrl,Child From MenuClass Where Depth=0 and ShowOnTop=True order by RootID"
  rs.open sql,conn,1,1
  do while not rs.eof
	response.Write("<li><a href='/"&rs("LinkUrl")&"'>"&rs("ClassName")&"</a></li> ")
  	rs.movenext
  loop 
  rs.close
  set rs=nothing
end sub
'==================================================
'过程名：ListFileName
'作  用：列表页面url处理
'参  数：无
'==================================================
function ListFileName()
    const MaxPerPage=10
	totalPut = 0
	currentPage = 0
	TotalPages = 0	
	if request("page_no")<>"" And isnumeric(request("page_no")) then
		currentPage=cint(request("page_no"))
	else
		currentPage=1
	end if
	QUERY_STRING_arr = split(QUERY_STRING,"&")
	ListFileName = "/"
	if PageName<>"" Then
	ListFileName = ListFileName&PageName
	End IF
	if ubound(QUERY_STRING_arr)>0 Then
		for i=0 to ubound(QUERY_STRING_arr)-1
			param_arr = split(QUERY_STRING_arr(i),"=")
			ListFileName = "/"&param_arr(1)&ListFileName
		next
	end if
	if ubound(QUERY_STRING_arr)=0 Then
			param_arr = split(QUERY_STRING_arr(0),"=")
			if param_arr(0)<>"page_no" Then ListFileName = "/"&param_arr(1)&ListFileName
	End if
	if isnumeric(currentPage) And currentPage>1 Then
	ListFileName = replace(ListFileName,".html","/"&currentPage&".html")
	End IF
	if isnumeric(currentPage) And currentPage=1 Then
	ListFileName = ListFileName&PageName
	End IF
End function

'***********************************************
'过程名：showpage
'作  用：显示"上一页 下一页"等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages)
	dim n, i,strTemp,strUrl,strUnit
	strUnit = "条信息"
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td style='background-color:#FFF;color:#000'>"
	if ShowTotal=true then
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=sfilename
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	PageName2 = replace(PageName,".html","")
	re.Pattern= PageName2&"(.*)\.html"
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName) & "'>首页</a>&nbsp;"
			if CurrentPage=2 Then
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName) & "'>上一页</a>&nbsp;"
			else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&""&CurrentPage-1&".html") & "'>上一页</a>&nbsp;"
			End If
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&""&CurrentPage+1&".html")& "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&""&n&".html")&"'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=true then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:goto(this,showpages)'>"
    	for i = 1 to n
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	if n>1 Then
		response.write strTemp
	End If
end sub
%>
