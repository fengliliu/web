<!--#Include File="Check_Sql.asp"-->
<!--#include file="Conn.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="Function.asp"-->

<%
dim strFileName,MaxPerPage,ShowSmallClassType
dim totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime
dim founderr, errmsg
dim BigClassName,SmallClassName,keyword,strField
dim rs,sql,sqlProduct,rsProduct,sqlSearch,rsSearch,sqlBigClass,rsBigClass
BeginTime=Timer
BigClassName=Trim(request("BigClassName"))
SmallClassName=Trim(request("SmallClassName"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=replace(replace(replace(replace(keyword,"'","��"),"<","&lt;"),">","&gt;")," ","&nbsp;")
end if
strField=trim(request("Field"))

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

sqlBigClass="select * from BigClass order by BigClassID"
Set rsBigClass= Server.CreateObject("ADODB.Recordset")
rsBigClass.open sqlBigClass,conn,1,1

'=================================================
'��������ShowSmallClass_Tree
'��  �ã�����Ŀ¼��ʽ��ʾ��Ŀ
'��  ������
'=================================================
Sub ShowSmallClass_Tree()
%>
<SCRIPT language=javascript>
function opencat(cat,img){
	if(cat.style.display=="none"){
	cat.style.display="";
	img.src="img/class2.gif";
	}	else {
	cat.style.display="none"; 
	img.src="img/class1.gif";
	}
}
</Script>
<TABLE cellSpacing=0 cellPadding=0 width="99%" border=0>
<%
dim i
set rsbig = server.CreateObject ("adodb.recordset")
		sql="select * from BigClass"
		rsbig.open sql,conn,1,1
if rsbig.eof and rsbig.bof then
	Response.Write "��Ŀ���ڽ����С���"
else
	i=1
	do while not rsbig.eof
%>
	<TR>
		<TD language=javascript onmouseup="opencat(cat10<%=i%>000,&#13;&#10; img10<%=i%>000);" id=item$pval[catID]) style="CURSOR: hand" width=34 height=24 align=center><IMG id=img10<%=i%>000 src="img/class1.gif" width=20 height=20></TD>
		<TD width="662"><a href='Product.asp?BigClassName=<%=rsbig("BigClassName")%>'><%=rsbig("BigClassName")%></a></TD>
	</TR>
	<TR>
		<TD id=cat10<%=i%>000 <%if rsbig("BigClassName")=BigClassName then 
		     response.write "style='DISPLAY'"   
		    else  
		     response.write "style='DISPLAY: none'"
		    end if%> colspan="2">
<%
dim rsSmall,sqls,j
set rsSmall = server.CreateObject ("adodb.recordset")
sqls="select * from SmallClass where BigClassName='" & rsbig("BigClassName") & "' order by SmallClassID"
rsSmall.open sqls,conn,1,1
if rsSmall.eof and rsSmall.bof then
	Response.Write "û��С��"
else
	j=1
	do while not rsSmall.eof
%>
&nbsp;<IMG height=20 src="img/class3.gif" width=26 align=absMiddle border=0><a href="Product.asp?BigClassName=<%=rsSmall("BigClassName")%>&Smallclassname=<%=rsSmall("SmallClassName")%>"><%=rsSmall("SmallClassName")%></a><BR>
<%
	rsSmall.movenext
	j=j+1
	loop
end if
rsSmall.close
set rsSmall=nothing
%>
		</TD>
	</TR>
<%
	rsbig.movenext
	i=i+1
	loop
	rsbig.close
    set rsbig=nothing
end if
%>
</TABLE>
<%
end Sub 

'=================================================
'��������ShowVote
'��  �ã���ʾ��վ����
'��  ������
'=================================================
sub ShowVote()
	dim sqlVote,rsVote,i
	sqlVote="select top 1 * from Vote where IsSelected=True"
	Set rsVote= Server.CreateObject("ADODB.Recordset")
	rsVote.open sqlVote,conn,1,1
	if rsVote.bof and rsVote.eof then 
		response.Write "&nbsp;û���κε���"
	else
		response.write "<form name='VoteForm' method='post' action='vote.asp' target='_blank'><td>"
		response.write rsVote("Title") & "<br><br>"
		if rsVote("VoteType")="Single" then
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='radio' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		else
			for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
				response.Write "<input type='checkbox' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
			next
		end if
		response.write "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
		response.write "<input name='Action' type='hidden' value='Vote'>"
		response.write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
		response.write "<a href='javascript:VoteForm.submit();'><img src='images/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
        response.write "<a href='Vote.asp?Action=Show' target='_blank'><img src='images/voteView.gif' width='52' height='18' border='0'></a>"
		response.write "</td></form>"
	end if
	rsVote.close
	set rsVote=nothing
end sub

'=================================================
'��������ShowClassGuide
'��  �ã���ʾ��Ŀ����λ��
'��  ������
'=================================================
sub ShowClassGuide()
	response.write  "&nbsp;<a href='Products.asp'>��Ʒ</a>&nbsp;&gt;&gt;"
	if BigClassName="" and SmallClassName="" then
		response.write "&nbsp;���в�Ʒ"
	else
		if BigClassName<>"" then
			response.write "&nbsp;<a href='Product.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
			if SmallClassName<>"" then
				response.write "<a href='Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>"
			else
				response.write "����С��"
			end if
		end if
	end if
end sub

'=================================================
'��������ShowProductTotal
'��  �ã���ʾ��������
'��  ������
'=================================================
sub ShowProductTotal()
	dim sqlTotal
	dim rsTotal
	sqlTotal="select Count(*) from Product where Passed=True "
	if BigClassName<>"" then
		sqlTotal=sqlTotal & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlTotal=sqlTotal & " and SmallClassName='" & SmallClassName & "' "
		end if	
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "���� 0 ����Ʒ"
	else
		totalPut=rsTotal(0)
		response.Write "���� " & totalPut & " ����Ʒ"
	end if
	rsTotal.close
	set rsTotal=nothing
end sub

'=================================================
'��������ShowProduct
'=================================================
sub ShowProduct(TitleLen)
	if TitleLen<0 or TitleLen>200 then
		TitleLen=50
	end if
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlProduct="select top " & MaxPerPage	
	else
		sqlProduct="select "
	end if

	sqlProduct=sqlProduct & " ID,Product_Id,BigClassName,SmallClassName,IncludePic,Title,Price,Spec,Unit,Memo,DefaultPicUrl,UpdateTime,Hits from Product where Passed=True "
	
	if BigClassName<>"" then
		sqlProduct=sqlProduct & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlProduct=sqlProduct & " and SmallClassName='" & SmallClassName & "' "
		end if
	end if
	sqlProduct=sqlProduct & " order by UpdateTime desc"
	Set rsProduct= Server.CreateObject("ADODB.Recordset")
	rsProduct.open sqlProduct,conn,1,1
	if rsProduct.bof and  rsProduct.eof then
		response.Write("<br><li>û���κβ�Ʒ</li>")
	else
		if currentPage=1 then
			call ProductContent(TitleLen)
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsProduct.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsProduct.bookmark
            	call ProductContent(TitleLen)
        	else
	        	currentPage=1
           		call ProductContent(TitleLen)
	    	end if
		end if
	end if
	rsProduct.close
	set rsProduct=nothing
end sub

sub ProductContent(intTitleLen)
   	dim i,strTemp
    i=0
	do while not rsProduct.eof
		strTemp=""		
		strTemp= strTemp & "<table width=100% border=0 cellspacing=3 cellpadding=0>"
                strTemp= strTemp & "<tr>"
                strTemp= strTemp & "<td width=25% rowspan=6>"
                strTemp= strTemp & "<div align=center><a href=ProductShow.asp?ID=" & rsProduct("ID") & ">" 
				
				fileExt=lcase(getFileExtName(rsProduct("DefaultPicUrl")))
				if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
                strTemp= strTemp & "<img style='BORDER-LEFT-COLOR: #cccccc; BORDER-BOTTOM-COLOR: #cccccc; BORDER-TOP-COLOR: #cccccc; BORDER-RIGHT-COLOR: #cccccc' src=" & rsProduct("DefaultPicUrl") & " width=105 height=80 onload='javascript:DrawImage(this);'>" 
				else
				 if fileext="swf" then
				    strTemp= strTemp & "<object  classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='105' height='84'>"
					strTemp= strTemp &"<param name=movie value='"&rsProduct("DefaultPicUrl")&"'>"
					strTemp= strTemp &"<param name=quality value=high>"
					strTemp= strTemp &"<param name='Play' value='-1'>"
					strTemp= strTemp &"<param name='Loop' value='0'>"
					strTemp= strTemp &"<param name='Menu' value='-1'>"					
					strTemp= strTemp &"<embed  src='"&rsProduct("DefaultPicUrl")&"' width='105' height='84' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed> </object>"												
			   end if
		      end if			 
				 
                strTemp= strTemp & "</a></div></td>"
                strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ���ƣ�</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=ProductShow.asp?ID=" & rsProduct("ID") & ">" & rsProduct("Title") & ""
                strTemp= strTemp & "</a></td>"						
				
				'strTemp= strTemp & "</tr><tr>" 
                'strTemp= strTemp & "<td width=12% height=12>"
                'strTemp= strTemp & "��Ʒ�ۼۣ�</td>"
                'strTemp= strTemp & "<td>" & rsProduct("Price") & "Ԫ</td>"	
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ���</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsProduct("Spec") & ""
                strTemp= strTemp & "</a></td>"
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ��ע��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsProduct("Memo") & ""
                strTemp= strTemp & "</a></td>"			
				
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td height=12>"
                strTemp= strTemp & "��Ʒ���</td>"
                strTemp= strTemp & "<td><a href=Product.asp?BigClassName="& rsProduct("BigClassName")&">"&rsProduct("BigClassName")&"</a> �� "
                strTemp= strTemp & "<a href=Product.asp?BigClassName=" & rsProduct("BigClassName") & "&SmallClassName=" & rsProduct("SmallClassName") & ">" & rsProduct("SmallClassName") & ""
                strTemp= strTemp & "</a></td>"
                strTemp= strTemp & "</tr><tr>" 				

			   
                strTemp= strTemp & "<td height=12>��Ʒ��Ϣ��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=ProductShow.asp?ID=" & rsProduct("ID") & " target=_blank><img src=Img/arrow_7.gif border=0></a></td>"
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td colspan=2>"
                strTemp= strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0>"
                strTemp= strTemp & "<tr>" 
                strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center></div></td>"
                
				strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center><input name='Product_Id' type='checkbox'    id='Product_Id' value="&cstr(rsProduct("Product_Id"))&"> ѡȡ"
                strTemp= strTemp & "</div></td>"
				
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
                strTemp= strTemp & "</td>"
                strTemp= strTemp & "</tr><tr>" 
                strTemp= strTemp & "<td height=1 colspan=3 bgcolor=#CCCCCC></td>"
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
		response.write strTemp
		rsProduct.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 


'=================================================
'��������ShowSearchTerm
'��  �ã���ʾ����������Ϣ
'��  ������
'=================================================
sub ShowSearchTerm()
	response.write "&nbsp;��Ʒ����&nbsp;&gt;&gt; "
	if BigClassName<>"" then
		response.write "<a href='Product.asp?BigClassName=" & BigClassName & "'>" & BigClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		if SmallClassName<>"" then
			response.write "<a href='Product.asp?BigClassName=" & BigClassName & "&SmallClassName=" & SmallClassName & "'>" & SmallClassName & "</a>&nbsp;&gt;&gt;&nbsp;"
		else
			response.write "����С��&nbsp;&gt;&gt;&nbsp;"
		end if
	end if
	if keyword<>"" then
	  select case strField
		case "Title"
			response.Write "�����к��� <font color=red>"&keyword&"</font> �Ĳ�Ʒ"
		case "Content"
			response.Write "˵������ <font color=red>"&keyword&"</font> �Ĳ�Ʒ"
		case else
			response.Write "�����к��� <font color=red>"&keyword&"</font> �Ĳ�Ʒ"
	  end select
	else
	  response.write "&nbsp;���в�Ʒ"
	end if
end sub

'=================================================
'��������SearchResultTotal
'��  �ã���ʾ�����������
'��  ������
'=================================================
sub SearchResultTotal()
	dim rsTotal,sqlTotal
	sqlTotal="select count(*) from Product where Passed=True "
	if BigClassName<>"" then
		sqlTotal=sqlTotal & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlTotal=sqlTotal & " and SmallClassName='" & SmallClassName & "' "
		end if	
	end if
	if keyword<>"" then
		select case strField
			case "Title"
				sqlTotal=sqlTotal & " and Title like '%" & keyword & "%' "
			case "Content"
				sqlTotal=sqlTotal & " and Content like '%" & keyword & "%' "
			case else
				sqlTotal=sqlTotal & " and Title like '%" & keyword & "%' "
		end select
	end if
	Set rsTotal= Server.CreateObject("ADODB.Recordset")
	rsTotal.open sqlTotal,conn,1,1
	if rsTotal.eof and rsTotal.bof then
		totalPut=0
		response.write "���� 0 ����Ʒ"
	else
		totalPut=rsTotal(0)
		response.Write "���ҵ� " & totalPut & " ����Ʒ"
	end if
end sub

'=================================================
'��������ShowSearchResult
'��  �ã���ҳ��ʾ�������
'��  ������
'=================================================
sub ShowSearchResult()
    if currentpage<1 then
	   	currentpage=1
    end if
	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	   		currentpage= totalPut \ MaxPerPage
		else
		   	currentpage= totalPut \ MaxPerPage + 1
		end if
   	end if
	if currentPage=1 then
        sqlSearch="select top " & MaxPerPage	
	else
		sqlSearch="select "
	end if

	sqlSearch=sqlSearch & " * from Product where Passed=True "
	if BigClassName<>"" then
		sqlSearch=sqlSearch & " and BigClassName='" & BigClassName & "' "
		if SmallClassName<>"" then
			sqlSearch=sqlSearch & " and SmallClassName='" & SmallClassName & "' "
		end if	
	end if
	if keyword<>"" then
		select case strField
			case "Title"
				sqlSearch=sqlSearch & " and Title like '%" & keyword & "%' "
			case "Content"
				sqlSearch=sqlSearch & " and Content like '%" & keyword & "%' "
			case else
				sqlSearch=sqlSearch & " and Title like '%" & keyword & "%' "
		end select
	end if
	sqlSearch=sqlSearch & " order by ID desc"
	Set rsSearch= Server.CreateObject("ADODB.Recordset")
	rsSearch.open sqlSearch,conn,1,1
 	if rsSearch.eof and rsSearch.bof then 
       		response.write "<p align='center'><br><br>û�л�û���ҵ��κβ�Ʒ</p>" 
   	else 
   		if currentPage=1 then 
       		call SearchResultContent()
   		else 
       		if (currentPage-1)*MaxPerPage<totalPut then 
       			rsSearch.move  (currentPage-1)*MaxPerPage 
       			dim bookmark 
       			bookmark=rsSearch.bookmark 
       			call SearchResultContent()
      		else 
        		currentPage=1 
       			call SearchResultContent()
      		end if 
	   	end if 
   	end if 
   	rsSearch.close 
   	set rsSearch=nothing   
end sub

sub SearchResultContent()
	dim i,strTemp
    i=0
	do while not rsSearch.eof
		strTemp=""		
		strTemp= strTemp & "<table width=100% border=0 cellspacing=3 cellpadding=0>"
                strTemp= strTemp & "<tr>"
                strTemp= strTemp & "<td width=25% rowspan=6>"
                strTemp= strTemp & "<div align=center><a href=ProductShow.asp?ID=" & rsSearch("ID") & ">" 
				
				fileExt=lcase(getFileExtName(rsSearch("DefaultPicUrl")))
				if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then
                strTemp= strTemp & "<img border=0 src=" & rsSearch("DefaultPicUrl") & " width=105 height=80>" 
				else
				 if fileext="swf" then
				    strTemp= strTemp & "<object  classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='105' height='84'>"
					strTemp= strTemp &"<param name=movie value='"&rsSearch("DefaultPicUrl")&"'>"
					strTemp= strTemp &"<param name=quality value=high>"
					strTemp= strTemp &"<param name='Play' value='-1'>"
					strTemp= strTemp &"<param name='Loop' value='0'>"
					strTemp= strTemp &"<param name='Menu' value='-1'>"
					strTemp= strTemp &"<param name='wmode' value='transparent'>"
					strTemp= strTemp &"<embed src='"&rsSearch("DefaultPicUrl")&"' width='105' height='84' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed> </object>"												
			   end if
		      end if			 
				 
                strTemp= strTemp & "</a></div></td>"
                strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ���ƣ�</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=ProductShow.asp?ID=" & rsSearch("ID") & ">" & rsSearch("Title") & ""
                strTemp= strTemp & "</a></td>"
				
				'strTemp= strTemp & "</tr><tr>" 
                'strTemp= strTemp & "<td width=12% height=12>"
                'strTemp= strTemp & "��Ʒ�ۼۣ�</td>"
                'strTemp= strTemp & "<td>" & rsSearch("Price") & "Ԫ</td>"				
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ���</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsSearch("Spec") & ""
                strTemp= strTemp & "</a></td>"
				
				strTemp= strTemp & "</tr><tr>"
				strTemp= strTemp & "<td width=12% height=12>"
                strTemp= strTemp & "��Ʒ��ע��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & rsSearch("Memo") & ""
                strTemp= strTemp & "</a></td>"				
				
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td height=12>"
                strTemp= strTemp & "��Ʒ���</td>"
                strTemp= strTemp & "<td><a href=Product.asp?BigClassName=" & rsSearch("BigClassName") & ">" & rsSearch("BigClassName") & "</a> �� "
                strTemp= strTemp & "<a href=Product.asp?BigClassName=" & rsSearch("BigClassName") & "&SmallClassName=" & rsSearch("SmallClassName") & ">" & rsSearch("SmallClassName") & ""
                strTemp= strTemp & "</a></td>"
                strTemp= strTemp & "</tr><tr>" 

			   
                strTemp= strTemp & "<td height=12>��Ʒ��Ϣ��</td>"
                strTemp= strTemp & "<td>"
                strTemp= strTemp & "<a href=ProductShow.asp?ID=" & rsSearch("ID") & "><img src=Img/arrow_7.gif border=0></a></td>"
                strTemp= strTemp & "</tr><tr>"
                strTemp= strTemp & "<td colspan=2>"
                strTemp= strTemp & "<table width=100% border=0 cellpadding=0 cellspacing=0>"
                strTemp= strTemp & "<tr>" 
                strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center></div></td>"
                
				strTemp= strTemp & "<td width=50% height=12>"
                strTemp= strTemp & "<div align=center><input name='Product_Id' type='checkbox'    id='Product_Id' value="&cstr(rsSearch("Product_Id"))&"> ѡȡ"
                strTemp= strTemp & "</div></td>"
				
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
                strTemp= strTemp & "</td>"
                strTemp= strTemp & "</tr><tr>" 
                strTemp= strTemp & "<td height=1 colspan=3 bgcolor=#CCCCCC></td>"
                strTemp= strTemp & "</tr>"
                strTemp= strTemp & "</table>"
		response.write strTemp
		rsSearch.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop
end sub 


'=================================================
'��������ShowSearch
'��  �ã���ʾ������������
'��  ����ShowType ----��ʾ��ʽ��1Ϊ����2Ϊ����
'=================================================
sub ShowSearch(ShowType)
	dim count
	if ShowType<>1 and ShowType<>2 then
		ShowType=1
	end if
	set rs=server.createobject("adodb.recordset")
	sql = "select * from SmallClass order by SmallClassID asc"
	rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>
<table border="0" cellpadding="2" cellspacing="0" align="center">
  <form method="Get" name="myform" action="search.asp">
    <tr>
      <td height="28"> <select name="Field" size="1">
          <option value="Title" selected>��Ʒ����</option>
          <option value="Content">��Ʒ˵��</option>
        </select> 
        <%if ShowType=1 then%>
      </td>
    </tr>
    <tr>
      <td height="28"> 
        <%end if%>
        <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
          <option selected value="">���д���</option>
          <%
if not (rsBigClass.bof and rsBigClass.eof) then
	rsBigClass.movefirst
	do while not rsBigClass.eof
        response.Write "<option value='" & trim(rsBigClass("BigClassName")) & "'>" & trim(rsBigClass("BigClassName")) & "</option>"
	   	rsBigClass.movenext
	loop
end if
%>
        </select> 
        <%if ShowType=1 then%>
      </td>
    </tr>
    <tr>
      <td height="28"> 
        <%end if%>
        <select name="SmallClassName">
          <option selected value="">����С��</option>
        </select> 
        <%if ShowType=1 then%>
      </td>
    </tr>
    <tr>
      <td height="28"> 
        <%end if%>
        <input type="text" name="keyword"  size=12 value="�ؼ���" maxlength="50" onFocus="this.select();"> 
        <input type="submit" name="Submit"  value="����"> </td>
    </tr>
  </form>
</table>
<%
end sub

'=================================================
'��������ShowAllClass
'��  �ã���ʾ������Ŀ����Ŀ������
'��  ������
'=================================================
sub ShowAllClass()
	if rsBigClass.bof and rsBigClass.eof then 
		response.Write "&nbsp;û���κ���Ŀ"
	else
		dim sqlClass,rsClass,strClassName
		rsBigClass.movefirst
		do while not rsBigClass.eof
			strClassName= "��<a href='Product.asp?BigClassName=" & rsBigClass("BigClassName") & "'><b>" & rsBigClass("BigClassName") & "</b></a>��<br><br>"
			sqlClass="select * from SmallClass where BigClassName='" & rsBigClass("BigClassName") & "' Order by SmallClassID"
			Set rsClass= Server.CreateObject("ADODB.Recordset")
			rsClass.open sqlClass,conn,1,1
			do while not rsClass.eof
				strClassName=strClassName & "&nbsp;<a href='Product.asp?BigClassName=" & rsClass("BigClassName") & "&SmallClassName=" & rsClass("SmallClassName") & "'>" & rsClass("SmallClassName") & "</a>&nbsp;"
				rsClass.movenext
			loop
			response.write strClassName & "<br><br>"
			rsBigClass.movenext
		loop
		rsClass.close
		set rsClass=nothing
	end if
end sub


'=================================================
'��������ShowProductContent
'��  �ã���ʾ���¾�������ݣ����Է�ҳ��ʾ
'��  ������
'=================================================
sub ShowProductContent()
	dim ID,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ID=rs("ID")
	strContent=rs("Content")
	response.write strContent
end sub
'=================================================
'��������ShowUserLogin
'��  �ã���ʾ�û���¼����
'��  ������
'=================================================
sub ShowUserLogin()
	dim strLogin
	If Session("UserName")="" Then
    	strLogin= "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		strLogin=strLogin & "<form action='UserLogin.asp' method='post' name='UserLogin' onSubmit='return CheckForm();'>"
        strLogin=strLogin & "<tr><td height='25' align='right'>�û�����</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr><td height='25' align='right'>��&nbsp;&nbsp;�룺</td><td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td></tr>"
        strLogin=strLogin & "<tr align='center'><td height='25' colspan='2'><input name='Login' type='submit' id='Login' value=' ��¼ '> <input name='Reset' type='reset' id='Reset' value=' ��� '>"
        strLogin=strLogin & "</td></tr>"
        strLogin=strLogin & "<tr><td height='20' align='center' colspan='2'><a href='UserReg.asp' target='_blank'>���û�ע��</a>&nbsp;&nbsp;<a href='GetPassword.asp' target='_blank'>�������룿</a></td></tr>"      
        strLogin=strLogin & "</form></table>"
		response.write strLogin
%>
<script language=javascript>
	function CheckForm()
	{
		if(document.UserLogin.UserName.value=="")
		{
			alert("�������û�����");
			document.UserLogin.UserName.focus();
			return false;
		}
		if(document.UserLogin.Password.value == "")
		{
			alert("���������룡");
			document.UserLogin.Password.focus();
			return false;
		}
	}
</script>
<%
	Else 
		response.write "��ӭ����" & Session("UserName") & "<br><br>"
		response.write "�û�������壺<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='UserServer.asp'><b>�����Ա����</b></a><br><br>"
	end if
end sub
%>