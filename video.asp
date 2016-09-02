<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/conn.asp" -->
<!--#include file="inc/config.asp" -->
<!--#include file="inc/function.asp" -->

<%
	title="视频中心"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=SiteName%></title>
<meta http-equiv="keywords" content="<%=SiteKeyword%>" />
<meta http-equiv="description" content="<%=SiteDescription%>" />
<script type="text/javascript" src="scripts/swfobject.js"></script>
    <script type="text/javascript" src="Scripts/jssor.slider-21.1.5.min.js"></script>
<link rel="stylesheet" href="/css/index.css" />
<link rel="stylesheet" href="/css/common.css" />
<style>
.gsjj {
	float: left;
	height: 748px;
	width: 458px;
}.flash {
	height:748px;
	width:458px;
	overflow:hidden;
	float: left;
}
.flash2 {
	height:750px;
	width:460px;
	margin-top:-1px;
}
.flasht {
	padding:4px;
 border:1px solid;
}

</style>
</head>
    <script>
        jssor_1_slider_init = function() {

            var jssor_1_SlideshowTransitions = [
              {$Duration:1200,$Opacity:2}
            ];

            var jssor_1_options = {
              $AutoPlay: true,
              $SlideshowOptions: {
                $Class: $JssorSlideshowRunner$,
                $Transitions: jssor_1_SlideshowTransitions,
                $TransitionsOrder: 1
              },
              $ArrowNavigatorOptions: {
                $Class: $JssorArrowNavigator$
              },
              $BulletNavigatorOptions: {
                $Class: $JssorBulletNavigator$
              }
            };

            var jssor_1_slider = new $JssorSlider$("jssor_1", jssor_1_options);

            //responsive code begin
            //you can remove responsive code if you don't want the slider scales while window resizing
            function ScaleSlider() {
                var refSize = jssor_1_slider.$Elmt.parentNode.clientWidth;
                if (refSize) {
                    refSize = Math.min(refSize, 1000);
                    jssor_1_slider.$ScaleWidth(refSize);
                }
                else {
                    window.setTimeout(ScaleSlider, 30);
                }
            }
            ScaleSlider();
            $Jssor$.$AddEvent(window, "load", ScaleSlider);
            $Jssor$.$AddEvent(window, "resize", ScaleSlider);
            $Jssor$.$AddEvent(window, "orientationchange", ScaleSlider);
            //responsive code end
        };
        </script>
<body>
<!--#include file="head.asp" -->
<div class="content h275">

        <div id="jssor_1" style="position: relative;cursor:pointer; margin: 0 auto; top: 0px; left: 0px; width: 1000px; height: 320px; overflow: hidden; visibility: hidden;">
            <!-- Loading Screen -->
            <div data-u="loading" style="position: absolute; top: 0px; left: 0px;">
                <div style="filter: alpha(opacity=70); opacity: 0.7; position: absolute; display: block; top: 0px; left: 0px; width: 100%; height: 100%;"></div>
                <div style="position:absolute;display:block;background:url('imgs/loading.gif') no-repeat center center;top:0px;left:0px;width:100%;height:100%;"></div>
            </div>
            <div data-u="slides" style="cursor: default; position: relative; top: 0px; left: 0px; width: 1000px; height: 320px; overflow: hidden;">
                <div data-p="112.50" style="display: none;">
                    <img data-u="image" src="imgs/1.png" />
                </div>
                <div data-p="112.50" style="display: none;">
                    <img data-u="image" src="imgs/2.png" />
                </div>
                <div data-p="112.50" style="display: none;">
                    <img data-u="image" src="imgs/3.png" />
                </div>
                <div data-p="112.50" style="display: none;">
                    <img data-u="image" src="imgs/4.png" />
                </div>
                <a data-u="add" href="http://www.jssor.com/demos/simple-fade-slideshow.slider" style="display:none">Simple Fade Slideshow</a>

            </div>
            <!-- Bullet Navigator -->
            <div data-u="navigator" class="jssorb05" style="bottom:16px;right:16px;" data-autocenter="1">
                <!-- bullet navigator item prototype -->
                <div data-u="prototype" style="width:16px;height:16px;"></div>
            </div>
            <!-- Arrow Navigator -->
            <span data-u="arrowleft" class="jssora12l" style="top:0px;left:0px;width:30px;height:46px;" data-autocenter="2"></span>
            <span data-u="arrowright" class="jssora12r" style="top:0px;right:0px;width:30px;height:46px;" data-autocenter="2"></span>
        </div>

        <script>
                jssor_1_slider_init();
            </script>
</div>
<div class="message container">
        <span style="font-size: 12px;color: #ffffff;line-height: 40px;"><img style="float: left;margin: 4px 7px 2px 8px;" src="imgs/arrows.png" alt="">本站公告通知：</span>
       <div class="area">
                   <marquee scrollAmount=2 width=840>我钟意网页树树我钟意网页树树我钟意网页树树我钟意网页树树我钟意网页树树我钟意网页树树</marquee>

</div>
<div class="content">
<!--#include file="left.asp" -->
<div  class="clearfloat area-container">

<div class="cont_you">
<div class="toub"><img src="/images/huab.jpg" /></div>
<div class="you_nr">
    	<table width="740" border="0" cellpadding="0" cellspacing="0">
        	<tr>
            	<td valign="top" width="45">
                      <img src="/images/huab_shu.jpg" border="0"/>
                </td>
                <td valign="top">

<%
	strFileName = ListFileName()

	set rs = server.CreateObject("adodb.recordset")
	sql = " Select DefaultPicUrl,id,readme From movie order by id desc "
	rs.open sql,conn,1,1
	
	totalPut=rs.recordcount
	if rs.eof then response.Write("暂无信息")
	if not rs.eof Then
		if currentpage<1 then
			currentpage=1
		end if
		if (currentpage-1)*MaxPerPage>totalput then
			If (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			Else
				currentpage= totalPut \ MaxPerPage + 1
			End If
		End If
		if currentPage=1 then
			List
			'showpage strFileName,totalput,MaxPerPage,true,false
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rs.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rs.bookmark
				List
				'showpage strFileName,totalput,MaxPerPage,true,false
			else
				currentPage=1
				List
				'showpage strFileName,totalput,MaxPerPage,true,false
			end if
		end if
	End If
	'***********************************************
	'过程名：List
	'作  用：列表内容
	'参  数：/
	'***********************************************
	Sub List()
		dim i
		i=0
		Do While Not rs.Eof
%>
<dl class="ship"><dt><a href="/videoshow<%=rs("id")%>.html"><img src="/<%=rs("DefaultPicUrl")%>" /></a></dt><dd><a href="/videoshow/<%=rs("id")%>.html"><%=rs("readme")%></a></dd></dl>
<%
		i=i+1
	      if i>=MaxPerPage then exit do
	      rs.movenext
			loop
		  rs.close
		  set rs=nothing
	End Sub
%>

    <%
showpage strFileName,totalput,MaxPerPage,true,false
%>
                </td>
            </tr>
        </table>
</div>
</div>
</div>
</div>
</div>

<div class="clr"></div>
<!--#include file="bottom.asp" -->
</body>
</html>