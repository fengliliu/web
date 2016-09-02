<div class="bottom">
<div class="touxian"></div>
<p> <a href="index.asp">网站首页</a>|<a href="company.asp">公司简介</a>|<a href="news.asp">新闻中心</a>|<a href="chanp.asp">产品介绍</a>|<a href="huanjin.asp">环境展示</a>|<a href="zhisih.asp">相关知识</a>|<a href="liuyan.asp">在线留言</a>|<a href="contacts.asp">联系我们</a><br />
<%=Copyright%>
</p></div>
</body>
</html>
<%
	 if not isnull(conn) then
	 		conn.close
	 		set conn = nothing
	 end If
%>