<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="inc/conn.asp" -->
<!--#include file="inc/config.asp" -->
<!--#include file="inc/function.asp" -->
<%title="首页"%>
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
         jssor_2_slider_init = function() {

                    var jssor_2_SlideshowTransitions = [
                      {$Duration:1200,$Opacity:2}
                    ];

                    var jssor_2_options = {
                      $AutoPlay: true,
                      $SlideshowOptions: {
                        $Class: $JssorSlideshowRunner$,
                        $Transitions: jssor_2_SlideshowTransitions,
                        $TransitionsOrder: 1
                      },
                      $ArrowNavigatorOptions: {
                        $Class: $JssorArrowNavigator$
                      },
                      $BulletNavigatorOptions: {
                        $Class: $JssorBulletNavigator$
                      }
                    };

                    var jssor_2_slider = new $JssorSlider$("jssor_2", jssor_2_options);

                    //responsive code begin
                    //you can remove responsive code if you don't want the slider scales while window resizing
                    function ScaleSlider() {
                        var refSize = jssor_2_slider.$Elmt.parentNode.clientWidth;
                        if (refSize) {
                            refSize = Math.min(refSize, 375);
                            jssor_2_slider.$ScaleWidth(refSize);
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
<!--第一部分-->
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
</div>
<div class="changp01 mt10">
  <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">&nbsp;&nbsp;乡村文化礼仪 <span style="font-size: 9px">Rural cultural etiquette</span></span>
  <span class="gd"><a href="/sets.html">more<img src="imgs/right.png" alt="" class="more"></a></span></div>
<div class="lunh">
  <div id="lunhuan3">
    <div id="indemo">
      <div id="lunhuan4">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 12 id,title,FirstImageName from news where bigclassname='乡村文化' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <div><a href="/show<%=rs("id")%>.html"><img src="<%=rs("FirstImageName")%>"/></a><a href="/show<%=rs("id")%>.html"><span><%=rs("title")%></span></a></div>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
      </div>
      <div id="lunhuan5"></div>
    </div>
  </div>
  <script>
<!--
var speed=50;
var tab3=document.getElementById("lunhuan3");
var tab4=document.getElementById("lunhuan4");
var tab5=document.getElementById("lunhuan5");
tab5.innerHTML=tab4.innerHTML;
function Marquee(){
if(tab5.offsetWidth-tab3.scrollLeft<=0)
tab3.scrollLeft-=tab4.offsetWidth
else{
tab3.scrollLeft++;
}
}
var MyMar=setInterval(Marquee,speed);
tab3.onmouseover=function() {clearInterval(MyMar)};
tab3.onmouseout=function() {MyMar=setInterval(Marquee,speed)};
-->
</script>
</div>
<div style=" clear:both"></div>
<div class="container">
<div class="panner">
<div class="changp01 mt10" style="width: 248px;margin: 0px;!important;">
  <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">新闻动态 <span style="font-size: 9px">News information</span></span>
  <span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></div>
  <div class="con_lie ml10">
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 6 id,title from news where bigclassname='新闻动态' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
  </div>
</div>

<div class="panner">
<!--婚庆用品改的-->
<div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/shangyehuodong.jpg" style="height: 100%;width: 100%!important;"/></div>
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='服务项目' and smallclassname='商业活动' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
    <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>
  </div>
</div>
<div class="panner">
<div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/liyi.png" style="height: 100%;width: 100%!important;"/></div>
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='服务项目' and smallclassname='礼仪庆典' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>
<div class="panner">
<div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/yanchu.png" style="height: 100%;width: 100%!important;"/></div>
   <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='服务项目' and smallclassname='文艺演出' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>
</div>
<div class="container">
<div class="panner">
<div class="changp01 mt10" style="width: 248px;margin: 0px;!important;">
  <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">商务信息 <span style="font-size: 9px">News information</span></span>
  <span class="gd"><a href="/shangwuqingdian.html">more<img src="imgs/right.png" alt="" class="more"></a></span></div>
  <div class="con_lie ml10">
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 6 id,title from news where bigclassname='商务信息' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>
<div class="panner">
<div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/hunqing.png" style="height: 100%;width: 100%!important;"/></div>
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='服务项目' and smallclassname='婚庆服务' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>
<div class="panner">
<div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/shouqing.png" style="height: 100%;width: 100%!important;"/></div>
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='服务项目' and smallclassname='寿庆服务' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>


  <div class="con_lie ml10 bianx">
    <div class="tu"><img src="imgs/qiche.png" /></div>
    <ul class="lie01">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 5 id,title from news where bigclassname='商务庆典'  and smallclassname='汽车租凭' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <li><img src="images/heart.jpg" /><a href="/show<%=rs("id")%>.html" ><%=rs("title")%></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>   
    </ul>
        <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

  </div>
</div>
 <div class="container">
 <div class="panner">
<div class="changp01 mt10" style="width: 248px;margin: 0px;!important;">
 <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">联系我们 <span style="font-size: 9px">News information</span>
          <span class="gd"><a href="/contact.html">more<img src="imgs/right.png" alt="" class="more"></a></span></span></div>
     <div class="lianxi">
           <p>
             联系电话：<br/><%=mobile%><br/>
              座机电话：<br/><%=tel%><br/>
               联系人： <%=linkman%><br/>
             座机电话：<br/><%=tel%><br/>
             Q Q：442645168<br/>
             邮箱：442645168@qq.com<br/>
             网址：www.jxxcwh.com<br/>
             公司地址：<br/>
             萍乡市安源区火车站站前东路<br/>
           </p>
</div>


</div>
    <div class="panner" style="width: 375px">
    <div class="changp01 mt10" style="width: 373px;margin: 0px;!important;">
     <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">经典案例 <span style="font-size: 9px">News information</span></span>
              <span class="gd"><a href="/case.html">more<img src="imgs/right.png" alt="" class="more"></a></span></div>
            <div id="jssor_2" style="position: relative;cursor:pointer; margin: 0 auto; top: 0px; left: 0px; width: 375px; height: 215px; overflow: hidden; visibility: hidden;">
                <!-- Loading Screen -->
                <div data-u="loading" style="position: absolute; top: 0px; left: 0px;">
                    <div style="filter: alpha(opacity=70); opacity: 0.7; position: absolute; display: block; top: 0px; left: 0px; width: 100%; height: 100%;"></div>
                    <div style="position:absolute;display:block;background:url('imgs/loading.gif') no-repeat center center;top:0px;left:0px;width:100%;height:100%;"></div>
                </div>
                <div data-u="slides" style="cursor: default; position: relative; top: 0px; left: 0px; width: 375px; height: 215px; overflow: hidden;">
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
                    jssor_2_slider_init();
                </script>
</div>
   <div class="panner" style="width: 375px">
<div class="changp01 mt10" style="width: 373px;margin: 0px;!important;">
          <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">视频中心 <span style="font-size: 9px">News information</span></span>
          <span class="gd"><a href="/case.html">more<img src="imgs/right.png" alt="" class="more"></a></span></div>
        <embed src=" http://player.youku.com/player.php/sid/XMTMzMjc5NTEwNA==/v.swf" quality="high" width="375" height="215" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash"></embed>
        </div>
</div>
</div>
<div style=" clear:both"></div>
<div class="changp01 mt10">
 <span style="font-size: 11px"><img src="imgs/hudie.png" alt="" style="float: left;">户外俱乐部 <span style="font-size: 9px">News information</span></span>
  <span class="gd"><a href="/case.html">更多&gt;&gt;</a></span></div>
<div class="lunh">
  <div id="lunhuan">
    <div id="indemo">
      <div id="lunhuan1">
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 12 id,title,FirstImageName from news where bigclassname='户外俱乐部' order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
      <div><a href="/show<%=rs("id")%>.html"><img src="<%=rs("FirstImageName")%>"/></a><a href="/show/<%=rs("id")%>.html"><span><%=rs("title")%></span></a></div>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %> 
      </div>
      <div id="lunhuan2"></div>
    </div>
  </div>
  <script>
<!--
var speed=50;
var tab=document.getElementById("lunhuan");
var tab1=document.getElementById("lunhuan1");
var tab2=document.getElementById("lunhuan2");
tab2.innerHTML=tab1.innerHTML;
function Marquee(){
if(tab2.offsetWidth-tab.scrollLeft<=0)
tab.scrollLeft-=tab1.offsetWidth
else{
tab.scrollLeft++;
}
}
var MyMar=setInterval(Marquee,speed);
tab.onmouseover=function() {clearInterval(MyMar)};
tab.onmouseout=function() {MyMar=setInterval(Marquee,speed)};
-->
</script> 
</div>
<div style=" clear:both"></div>
<div class="changp01 mt10" style="display: none">
  <h2>&nbsp;&nbsp;友情链接</h2>
  <span class="gd"></span></div>
<div class="lianjie" style="display: none">
  <ul>
    <%
		set rs = server.CreateObject("adodb.recordset")
		sql = "select top 12 LogoUrl,SiteUrl,SiteName from FriendLinks where IsOK=true order by id desc"
		rs.open sql,conn,1,1
		if rs.eof then
		response.Write("<li>暂无信息</li>")
		else
		do while not rs.eof
	%>
    <li><a href="<%="http://"&replace(rs("SiteUrl"),"http://","")%>"><img src="<%=rs("LogoUrl")%>" /></a></li>
   <%
   		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
   %>
  </ul>
      <li class="li-more"><span class="gd"><a href="/newclass.html">more<img src="imgs/right.png" alt="" class="more"></a></span></li>

</div>
<!--#include file="bottom.asp" -->
<EMBED src="m.mp3"autostart="true" loop="ture"  width="0" height="0" style=" width:0px; height:0px;"></embed>

</body>
</html>
