<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/conn.asp" -->
<!--#include file="inc/config.asp" -->
<!--#include file="inc/page.asp" -->
<!--#include file="inc/check_sql.asp" -->
<%
	title="信息详情"
	id=request("id")

	if isnumeric(id) and id<>"" then
		set rs = server.CreateObject("adodb.recordset")
		sql = "select title,content from news where id="&id
		rs.open sql,conn,1,1
		if not rs.eof then
			content = rs("content")
			title = rs("title")
		end if
		rs.close
		set rs=nothing
	end if
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
      <div class="you_nr_le"><img src="/images/huab_shu.jpg" /></div>
      <dl class="xinxi">
        <dd><b><%=title%></b></dd>
      </dl>
      <div class="xixi_nr"><%=content%></div>
    </div>
  </div>
  </div>
  </div>
</div>
<div class="clr"></div>
<!--#include file="bottom.asp" -->
</body>
</html>
