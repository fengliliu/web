<div class="bottom">
<div class="touxian"></div>
<p> <a href="index.asp">��վ��ҳ</a>|<a href="company.asp">��˾���</a>|<a href="news.asp">��������</a>|<a href="chanp.asp">��Ʒ����</a>|<a href="huanjin.asp">����չʾ</a>|<a href="zhisih.asp">���֪ʶ</a>|<a href="liuyan.asp">��������</a>|<a href="contacts.asp">��ϵ����</a><br />
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