<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="inc/conn.asp" -->
<!--#include file="inc/config.asp" -->
<!--#include file="inc/page.asp" -->
<!--#include file="inc/check_sql.asp" -->
<%
	classtype=request("type")
	title="服务项目"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%>-<%=title%></title>
<meta http-equiv="keywords" content="<%=SiteKeyword%>" />
<meta http-equiv="description" content="<%=SiteDescription%>" />
<link rel="stylesheet" href="/css/common.css" />
<style>
.zuo_mul01{ width:225px; float:left;height:auto;}
</style>
</head>

<body>
<!--#include file="head.asp" -->
<!--#include file="bottom.asp" -->
</body>
</html>