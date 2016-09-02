<%
dim conn,db
dim connstr
db="Databases/qianxihunqing.mdb" '数据库文件位置
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
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
end if
sub CloseConn()
	conn.close
	set conn=nothing
end sub
dim currentPage
QUERY_STRING = request.ServerVariables("QUERY_STRING")
PageName = request.ServerVariables("SCRIPT_NAME")
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = replace(PageName,".asp",".html")
PageName = replace(PageName,"Show","")
%>