<%
'*************************************************
'��������loseHtml
'��  �ã�ȥ��HTML
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ��loseHtml
'*************************************************
function loseHtml(str)
	if str="" then
		loseHtml=""
		exit function
	end if
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="\<[^\>]+\>"
str=re.Replace(str,"")
str=replace(replace(str,"<",""),">","")
loseHtml=str
end function

function StrLen(Str)
  if Str="" or isnull(Str) then 
    StrLen=0
    exit function 
  else
    dim regex
    set regex=new regexp
    regEx.Pattern ="[^\x00-\xff]"
    regex.Global =true
    Str=regEx.replace(Str,"^^")
    set regex=nothing
    StrLen=len(Str)
  end if
end function

function StrLeft(Str,StrLen)
  dim L,T,I,C
  if Str="" then
    StrLeft=""
    exit function
  end if
  Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
  L=Len(Str)
  T=0
  for i=1 to L
    C=Abs(AscW(Mid(Str,i,1)))
    if C>255 then
      T=T+2
    else
      T=T+1
    end if
    if T>=StrLen then
      StrLeft=Left(Str,i) & "��"
      exit for
    else
      StrLeft=Str
    end if
  next
  StrLeft=Replace(Replace(Replace(replace(StrLeft," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

function StrReplace(Str)'�������滻�ַ�
  if Str="" or isnull(Str) then 
    StrReplace=""
    exit function 
  else
    StrReplace=replace(str," ","&nbsp;") '"&nbsp;"
    StrReplace=replace(StrReplace,chr(13),"&lt;br&gt;")'"<br>"
    StrReplace=replace(StrReplace,"<","&lt;")' "&lt;"
    StrReplace=replace(StrReplace,">","&gt;")' "&gt;"
  end if
end function

function ReStrReplace(Str)'д����滻�ַ�
  if Str="" or isnull(Str) then 
    ReStrReplace=""
    exit function 
  else
  	ReStrReplace=replace(Str,"&amp;","&") '"&amp;"
    ReStrReplace=replace(ReStrReplace,"&nbsp;"," ") '"&nbsp;"
    ReStrReplace=replace(ReStrReplace,"<br>",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;br&gt;",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;","<")' "&lt;"
    ReStrReplace=replace(ReStrReplace,"&gt;",">")' "&gt;"
  end if
end function

function HtmlStrReplace(Str)'д��Html��ҳ�滻�ַ�
  if Str="" or isnull(Str) then 
    HtmlStrReplace=""
    exit function 
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br>")'"<br>"
  end if
end function

function ViewNoRight(GroupID,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from WrzcNet_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  ViewNoRight=true
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then   ' ���û�Ȩ������С�ڲ�ƷȨ������, ����Ȩ����
	    ViewNoRight=false
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	    ViewNoRight=false
      end if
  end select
end function

Function GetUrl()
  GetUrl="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
  If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL=GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function

function HtmlSmallPic(GroupID,PicPath,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from WrzcNet_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  HtmlSmallPic=PicPath
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then HtmlSmallPic="Images/NoRight.jpg"
    case "="
      if not session("GroupLevel") = GroupLevel then HtmlSmallPic="Images/NoRight.jpg"
  end select
  if HtmlSmallPic="" or isnull(HtmlSmallPic) then HtmlSmallPic="Images/NoPicture.jpg"
end function

function IsValidMemName(memname)
  dim i, c
  IsValidMemName = true
  if not (3<=len(memname) and len(memname)<=16) then
    IsValidMemName = false
    exit function
  end if  
  for i = 1 to Len(memname)
    c = Mid(memname, i, 1)
    if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 and not IsNumeric(c) then
      IsValidMemName = false
      exit function
    end if
  next
end function

function IsValidEmail(email)
  dim names, name, i, c
  IsValidEmail = true
  names = Split(email, "@")
  if UBound(names) <> 1 then
    IsValidEmail = false
    exit function
  end if
  for each name in names
	if Len(name) <= 0 then
	  IsValidEmail = false
      exit function
    end if
    for i = 1 to Len(name)
      c = Mid(name, i, 1)
      if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 and not IsNumeric(c) then
        IsValidEmail = false
        exit function
      end if
	next
	if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
    end if
  next
  if InStr(names(1), ".") <= 0 then
    IsValidEmail = false
    exit function
  end if
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then
    IsValidEmail = false
    exit function
  end if
  if InStr(email, "..") > 0 then
    IsValidEmail = false
  end if
end function

'================================================
'��������FormatDate
'�����ã���ʽ������
'�Ρ�����DateAndTime            (ԭ���ں�ʱ��)
'       Format                 (�����ڸ�ʽ)
'����ֵ����ʽ���������
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
    '����12λ ֱ���� ��ʱ���ַ���
  Case "3"
    strDateTime = yy & m & d & h & mi    
    '����12λ ֱ���� ��ʱ���ַ���
  Case "4"
    strDateTime = yy & "��" & m & "��" & d & "��"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "��" & d & "��"
  Case "8"
    strDateTime = y & "��" & m & "��"
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
    '����12λ ֱ���� ��ʱ���ַ���
  Case Else
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function
'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function


'********************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ



'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'***************************************************
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


'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'****************************************************
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����ServerAddress  ----��������ַ
'        AddRecipient  ----�����˵�ַ
'        Subject       ----����
'        Body          ----�ż�����
'        Sender        ----�����˵�ַ
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
		err.clear
		exit function
	end if
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.Subject=Subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	Set JMail=nothing 
	if err then 
		SendMail=err.description
		err.clear
	else
		SendMail="OK"
	end if
end function

'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>�����ء�</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'****************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td height='20' class='title'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr><td height='100' class='tdbg' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>�����ء�</a></td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

function getFileExtName(fileName)
    dim pos
    pos=instrrev(filename,".")
    if pos>0 then 
        getFileExtName=mid(fileName,pos+1)
    else
        getFileExtName=""
    end if
end function 

'==================================================
'��������MainMenu
'��  �ã���ʾ������
'��  ��: 
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
'��������ChildMenu
'��  �ã���ʾ�ӵ���
'��  ��:  ParentID ����ID
'==================================================
sub ChildMenu(ParentID)
	dim rs,sql
	sql="select ClassID,ClassName,Depth,NextID,LinkUrl,Child From MenuClass where ParentID=" & ParentID & " order by OrderID asc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
	response.Write("<ul>")
	do while not rs.eof
	if instr(rs("LinkUrl"),PageName)>0  Then
		Readme = rs("ClassName")
		If PageID="" Then
		  PageID = ParentID
	  	end If
	End IF  
	response.Write("<li><a href='"&rs("LinkUrl")&"'>"&rs("ClassName")&"</a> ")
	if rs("Child")>0 Then
		call ChildMenu(rs("ClassID"))
	End IF
	response.Write("</li>")
	rs.movenext
	loop
	rs.close
	Set rs = nothing
  response.Write("</ul>")
End sub

'==================================================
'��������ShowLeftNav
'��  �ã���ʾ��ߵ���
'��  ��:  ID ͷ������ID
'==================================================
sub ShowLeftNav(ID)
	dim sqlClass,rsClass
	if ID<>"" Then
		  sqlClass="select ClassID,ClassName,LinkUrl,ParentID From MenuClass"
		  sqlClass= sqlClass & " where ParentID=" & ID & " order by OrderID asc"
		  Set rsClass = server.CreateObject("adodb.Recordset")
		  rsClass.open sqlClass,conn,1,1
		  response.write("<ul>")
		  Do While Not rsClass.Eof
		  	if instr(rsClass("LinkUrl"),PageName)>0 Then
		  		response.write("<li class='cur'><h2><a href='"&rsClass("LinkUrl")&"'>"&rsClass("ClassName")&"</a></h2></li>")
			else
				response.write("<li><h2><a href='"&rsClass("LinkUrl")&"'>"&rsClass("ClassName")&"</a></h2></li>")
			End if
		  rsClass.Movenext
		  Loop
		  rsClass.Close
		  Set rsClass = Nothing
		  		  response.write("</ul>")
	End If	  
end sub
'==================================================
'��������PageSqlID
'��  �ã���ȡ��ҳ����ID
'��  ��:  ID ͷ������ID
'==================================================
function PageSqlID(SqlTable,SqlWhere,HideSort,Page)
  dim idCount'��¼����
  dim pages'ÿҳ����
      pages=MaxPerPage_Search
  dim pagec'��ҳ��
      if Page="" Or Not isnumeric(Page) Then Page=1
      page=clng(Page)
	  
  dim datawhere'��������
	  datawhere=SqlWhere&" "&HideSort& " "
  dim Myself,PATH_INFO,QUERY_STRING'��ҳ��ַ�Ͳ���
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" then
	    Myself = PATH_INFO & "?"
	  elseif Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself= PATH_INFO & "?" & QUERY_STRING & "&"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'�������� asc,desc
      taxis="order by id desc "
  dim i'����ѭ��������
  '��ȡ��¼����
  dim PageSql,rsPage
	  PageSql="select count(ID) as idCount from ["& SqlTable &"]" & datawhere
	  set rsPage=server.createobject("adodb.recordset")
	  rsPage.open PageSql,conn,0,1
	  idCount=rsPage("idCount")
	  '��ȡ��¼����
	  if(idcount>0) then'�����¼����=0,�򲻴���
		if(idcount mod pages=0)then'�����¼��������ÿҳ����������,��=��¼����/ÿҳ����+1
		  pagec=int(idcount/pages)'��ȡ��ҳ��
		else
		  pagec=int(idcount/pages)+1'��ȡ��ҳ��
		end if
		'��ȡ��ҳ��Ҫ�õ���id============================================
		'��ȡ���м�¼��id��ֵ,��Ϊֻ��id�����ٶȺܿ�
		PageSql="select id from ["& SqlTable &"] " & datawhere & taxis
		set rsPage=server.createobject("adodb.recordset")
		rsPage.open PageSql,conn,1,1
		rsPage.pagesize = pages 'ÿҳ��ʾ��¼��
		if page < 1 then page = 1
		if page > pagec then page = pagec
		if pagec > 0 then rsPage.absolutepage = page  
		for i=1 to rsPage.pagesize
		  if rsPage.eof then exit for  
		  if(i=1)then
			PageSqlID=rsPage("id")
		  else
			PageSqlID=PageSqlID &","&rsPage("id")
		  end if
		  rsPage.movenext
		next
	  '��ȡ��ҳ��Ҫ�õ���id����============================================
	  end if
End function
'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
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
'==================================================
'��������ListFileName
'��  �ã��б�ҳ��url����
'��  ������
'==================================================
function ListFileName()
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
			ListFileName = "/"&param_arr(1)&ListFileName
	End if
	if isnumeric(currentPage) And currentPage>1 Then
	ListFileName = replace(ListFileName,".html","/"&currentPage&".html")
	End IF
	if isnumeric(currentPage) And currentPage=1 Then
	ListFileName = ListFileName&PageName
	End IF
End function

'***********************************************
'��������showpage
'��  �ã���ʾ"��һҳ ��һҳ"����Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages)
	dim n, i,strTemp,strUrl,strUnit
	strUnit = "����Ϣ"
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td style='background-color:#FFF;color:#000'>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=sfilename
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	PageName2 = replace(PageName,".html","")
	re.Pattern= PageName2&"(.*)\.html"
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName) & "'>��ҳ</a>&nbsp;"
			if CurrentPage=2 Then
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName) & "'>��һҳ</a>&nbsp;"
			else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&"/"&CurrentPage-1&".html") & "'>��һҳ</a>&nbsp;"
			End If
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&"/"&CurrentPage+1&".html")& "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & re.Replace(strUrl,PageName2&"/"&n&".html")&"'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=true then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:goto(this,showpages)'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	if n>1 Then
		response.write strTemp
	End If
end sub
'==================================================
'��������RollFriendLinks
'��  �ã�������ʾ��������վ��
'��  ������
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //��������rolllink1��rolllink2����ʱ
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink�������
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��
   rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��
</script>
<%
end sub
%>