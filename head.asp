
<div class="head">

  <div class="screen">
      <div class="logo fleft">
        <a href="">
            <img src="/imgs/logo.png" style="padding-top: 10px;padding-left: 25px" alt="">
        </a>
      </div>
      <div class="right-head fright">
      <div class="top-menu">
        <a href="">设为首页</a>
        <a href="">加入收藏</a>
        <a href="">联系我们</a>
</div>
        <div class="name">
        <p style="font-size: 30px;font-weight: 600;">萍乡市金玉满堂乡村礼仪文化传播有限公司</p>
        <p style="font-size: 19px;">pingxiangshijinyumantangxiangcunliyiwenhuachuanboyouxiangongsi</p>

</div>
          <ul class="daoh">
            <%call MainMenu()%>
          </ul>
</div>

  </div>
</div>
<%if title<>"首页" then%>
<div class="banner" style="display: none;"><img src="<%=BannerUrl%>" /></div>
<%end if%>