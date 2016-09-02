<%
'=================================
'
'    SOUTHIDC.NET(南方数据)
'    QQ:635495 MSN:jxhwq@hotmail.com
'    Email:hdz2008@163.com  
'    http://www.southidc.net
'    copyright(c)2004 科创网络
'
'=================================
%>
<html><head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<link href="Inc/southidc.css" rel="stylesheet" type="text/css">
<style type="text/css">
BODY {
	BACKGROUND:799ae1; MARGIN: 0px;
}

.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</style>
</head>

<BODY>
<table cellspacing="0" cellpadding="0" width="158" align="center">

	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="image/title_bg_quit.gif" height="25"> 
    <img src="Image/daohang.jpg"><span style="FONT-WEIGHT: bold; LEFT: 10px; COLOR: #215dc6;">网站管理平台</span></td>
	</tr>
	<tr>
    	<td align="center" onMouseOver="aa('up')" onMouseOut="StopScroll()">&nbsp; </td>
	</tr>
</table>
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td id="imgmenu1" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(1)" onMouseOut="this.className='menu_title';" style="cursor:hand" background=image/menudown.gif height="25"> 
      <span>系统管理 </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="Admin_Manage.asp">管理员管理</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="SiteConfig.asp">网站配置</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Bs.asp"></a><a target="main" href="Admin_DataBackup.asp">数据库备份</a></td>
          </tr>
        </table>
      </div>
		<br>
		</td>
	</tr>
	<tr>
    <td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>关于我们 </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
               <tr>
                <td><a target="main" href="AdminAboutus.asp"><font color="000000">关于我们管理</font></a></td>
              </tr>
              <tr>
                <td><a target="main" href="AdminAboutusAdd.asp"><font color="000000">添加关于我们</font></a></td>
              </tr>
               <tr>
                <td><a target="main" href="Admin_CompVisualize.asp"><font color="000000">广告管理</font></a></td>
              </tr>
              <tr>
                <td><a target="main" href="Admin_CompVisualizeAdd.asp"><font color="000000">添加广告</font></a></td>
              </tr>
              <tr>
                <td><a target="main" href="Admin_MovieAdd.asp"></a><a target="main" href="Admin_Movie.asp"><font color="000000">视频管理</font></a></td>
              </tr>
              <tr> 
                <td><a target="main" href="Admin_MovieAdd.asp"><font color="000000">添加视频</font></a></td>
              </tr>
			  
			</table>
		</div>
		<br>
		</td>
	</tr>
	<tr>
		
    <td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>信息管理 </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td> <a href="News_add.asp" target="main">添加信息内容</a> </td>
          </tr>
          <tr> 
            <td><a href="News_Manage.asp" target="main">管理全部信息</a></td>
          </tr>
          <tr> 
            <td><a href="News_ClassManage.asp" target="main">管理信息类别</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>	
	<tr>
    <td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>友情链接管理 </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_FriendLinks.asp"><font color="000000">友情链接管理</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>		
    <td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>菜单管理 </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr>
            <td><a href="Admin_Class_Menu.asp" target="main">中文菜单管理</a></td>
          </tr>
          <!--
          <tr> 
            <td><a href="Admin_Class_MenuEn.asp" target="main">英文菜单管理</a></td>
          </tr>
          -->
        </table>
		</div>
		<br>
		</td>
	</tr>

</table>
&nbsp;
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="image/title_bg_quit.gif" height="25"> 
      <span>Web Information</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					
              <td width="141" height="20" style="line-height: 150%;">  朝阳网络<br>
                Copyright： 朝阳网络<br>
                Design By： <a href="mailto:371029951@qq.com"><font color="000000">hyf</font></a><br>
					</td>
				</tr>
			</table>
			</div>
		</div>
		</td>
	</tr>
</table>
</div>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td align="center" onMouseOver="aa('Down')" onMouseOut="StopScroll()" valign="bottom">&nbsp; 
    </td>
	</tr>
</table>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="image/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="image/menudown.gif";
}
}

function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title="企业网站管理系统"; 
</script>
</BODY></HTML>
