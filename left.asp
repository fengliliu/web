
  <div class="con_zuo">
 <%if title="服务项目" then %>
    <div><img src="./imgs/fuwu.png" alt="" style="width: 100%; "></div>
    <div class="zuo_mul01 na">
    <ul>
    <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/商业活动service.html">商业活动</a></li>
    <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/礼仪庆典service.html">礼仪庆典</a></li>
    <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/文艺演出service.html">文艺演出</a></li>
        <li class="gud" id="hunqin">&nbsp;&nbsp;<img class="left-img" id="arrow" src="../imgs/arrow1.png" alt="">婚庆服务
        </li>
        <div id="fuwu" style="display: none">
            <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/场景布置service.html">场景布置</a></li>
          <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/婚车鲜花service.html">婚车鲜花</a></li>
          <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/婚礼道具service.html">婚礼道具</a></li>
          <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/婚礼外景service.html">婚礼外景</a></li>
          <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/婚纱化妆service.html">婚纱化妆</a></li>
          <li class="gud">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/摄影摄像service.html">摄影摄像</a></li>
</div>

  <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/寿庆服务service.html">寿庆服务</a></li>
    <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/汽车租凭service.html">汽车租凭</a></li>
</ul>
<script >
    var hunqin = document.getElementById("hunqin")
    var fuwu = document.getElementById("fuwu")
    var arrow = document.getElementById("arrow")
 var num = 1;
    hunqin.onclick = function() {
             if(num==0){
                fuwu.style.display = "none";
                arrow.src = "../imgs/arrow1.png";
                num=1;
              }else{
                fuwu.style.display = "block";
                arrow.src = "../imgs/arrow-down.png";
                num=0;
              }
    }
</script>
      <ul style="display: none">
      <%
	  	 set rs =server.CreateObject("adodb.recordset")
		 sql = "select * from SmallClass_New where bigclassname='服务项目'"
		 rs.open sql,conn,1,1
		 do while not rs.eof
		 smallclassname = rs("smallclassname")
		 if classtype=rs("smallclassname") then
	  %>
        <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/<%=server.URLEncode(smallclassname)%>service.html"><%=smallclassname%></a></li>
       <%else%>
        <li>&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt="">&nbsp;<a href="/<%=server.URLEncode(smallclassname)%>service.html"><%=smallclassname%></a></li>
      <%
	  	end if
	  	rs.movenext
		loop
		rs.close
		set rs=nothing
	  %>
      </ul>
    </div>
 <%else%>
    <div><img src="./imgs/fuwu.png" alt="" style="width: 100%;"></div>
    <div class="zuo_mul01">
          <ul>
    <%
      set rs = server.createobject("adodb.recordset")
      sql="select ClassID,ClassName,Depth,NextID,LinkUrl,Child,readme From MenuClass Where Depth=0 and ShowOnLeft=True order by RootID"
      rs.open sql,conn,1,1
      do while not rs.eof
      if rs("classname")=title then
    %>
        <li class="gud">&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/<%=rs("LinkUrl")%>"><%=rs("ClassName")%>&nbsp;<%=rs("readme")%></a></li>
        <%else%>
        <li>&nbsp;&nbsp;<img class="left-img" src="../imgs/arrow1.png" alt=""><a href="/<%=rs("LinkUrl")%>"><%=rs("ClassName")%>&nbsp;<%=rs("readme")%></a></li>
    <%
        end if
        rs.movenext
        loop
        rs.close
        set rs=nothing
    %>
          </ul>
    </div>
    <%end if%>
	    <div><img src="./imgs/contact.png" alt="" style="width: 100%;"></div>
    <div class="lianxi">
      <p>
        联系电话：<%=mobile%><br/>
         座机电话：<%=tel%><br/>
          联系人： <%=linkman%><br/>
        座机电话：<%=tel%><br/>
        Q Q：442645168<br/>
        邮箱：442645168@qq.com<br/>
        网址：www.jxxcwh.com<br/>
        公司地址：<br/>
        萍乡市安源区火车站站前东路<br/>
      </p>
    </div>
  </div>