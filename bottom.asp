<div class="bottom">
<p>
<%
  set rs = server.createobject("adodb.recordset")
  sql="select ClassID,ClassName,Depth,NextID,LinkUrl,Child From MenuClass Where Depth=0 and ShowOnTop=True order by RootID"
  rs.open sql,conn,1,1
  if not rs.eof then
  	response.Write("<a href='"&rs("LinkUrl")&"'>"&rs("ClassName")&"</a>")
	rs.movenext
  end if
  do while not rs.eof
  response.Write("&nbsp;|&nbsp;<a href='"&rs("LinkUrl")&"'>"&rs("ClassName")&"</a>")
  rs.movenext
  loop
  rs.close
  set rs=nothing
%>
</br>版权所有：萍乡金玉满堂乡村礼仪文化传播有限公司
</br>公司地址：萍乡市安源区火车站站前东路  联系电话：0799-3455666  13807993155
</br>技术支持：萍乡朝阳网络  备案号：赣ICP备123456789号
 </p></div>