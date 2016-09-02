<%
if session("AdminName") = "" then
    response.Redirect "Login.asp"
end if
%>

<html>
<head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<title>��ҵ��վ����ϵͳ</title>
<style type="text/css">
.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
.a2{BACKGROUND-COLOR: A4B6D7;}
</style>
</head>
<BODY leftMargin=0 topMargin=0 rightMargin=0> 
<CENTER>
<div style="width:100%;">
	<div style="width:7%;float:left;">
    	<div style="background: url('images/top.jpg') repeat-x; height:30px;"></div>
        <div style="background: url('images/logo2.gif') repeat-x;height:55px;"></div>
    </div>
    <div style="width:93%;float:left;background: url('images/top2.gif') repeat-x;height:85px;position:relative;">
   	  <div style="position:absolute;top:45px;left:200px;color:#275D8B;font-weight:bold;">ǧ�����ǻ�����վ����ƽ̨</div>
    <div style="position:absolute; left: 763px; top: 30px;"><a target="_top" href="../index.asp"><img src="images/index.jpg" border="0"></a></div>
    <div style="position:absolute; left: 851px; top: 30px;"><a target="_top" href="Loginout.asp"><img src="images/layout.jpg"" border="0"></a></div>
    </div>	
    
</div>
<%
select case Request("menu")
case ""
main
end select
%>
<% sub main %>
<body style="MARGIN: 0px" scroll=no onResize=javascript:parent.carnoc.location.reload()>
<script>
if(self!=top){top.location=self.location;}
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>
<table border="0" cellPadding="0" cellSpacing="0" width="100%">
	<tr>
    	<td height="5"></td>
    </tr>
</table>
<table border="0" cellPadding="0" cellSpacing="0" width="100%" HEIGHT="85%">
  <tr>
    <td align="middle" noWrap vAlign="center" id="frmTitle">
    <iframe frameBorder="0" id="carnoc" name="carnoc" scrolling=no src="LeftTree.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
    </iframe>
    </td>
    <td bgcolor="#275D8B" style="WIDTH: 9pt">
    <table border="0" cellPadding="0" cellSpacing="0" height="100%">
      <tr>
        <td style="HEIGHT: 100%" onClick="switchSysBar()">
        <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <span class="navPoint" id="switchPoint" title="�ر�/������">3</span><br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
        ��Ļ�л� </font></td>
      </tr>
    </table>
    </td>
		<td style="WIDTH: 100%; HEIGHT: 85%;">
		<iframe frameborder="0" id="main" name="main" scrolling="yes" src="sysadmin_view.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></iframe></td>
  </tr>
</table>
<script>
  if(window.screen.width<'1024'){switchSysBar()}
</script>
</body>
<%
end sub

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
%>
</CENTER>
</body>
</html>