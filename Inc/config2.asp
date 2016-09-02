<%
Const EnterpriseMail="http://mail.163.com"        '企业邮局
Const LogoUrl="images/logo.gif"        'Logo地址
Const BannerUrl="Images/0791idc.swf"        'Banner地址
Const IsFlash="Yes"        '是否为FLASH
Const width=""        '宽度
Const height=""        '高度
Const WebmasterName="weiqun"        '站长姓名
Const WebmasterEmail="southidc2008@163.com"        '站长信箱
Const ComPic_count="20"        '首页形象图数
Const New_count=20        '首面新闻资讯条数
Const MaxPerPage=20        '文章搜索页每页文章数
Const EnableProductCheck=""        '是否启用文章审核功能
Const EnableUploadFile="Yes"        '是否开放文件上传
Const MaxFileSize=200        '上传文件大小限制
Const SaveUpFilesPath="UploadFiles"        '存放上传文件的目录
Const UpFileType="gif|jpg|bmp|png|swf|doc|rar"        '允许的上传文件类型
Const DelUpFiles="Yes"        '删除文章时是否同时删除文章中的上传文件
Const SessionTimeout=30        'Session会话的保持时间
Const MailObject="Jmail"        '邮件发送组件
Const MailServer="smtp.163.com"        '用来发送邮件的SMTP服务器
Const MailServerUserName="feng"        '登录用户名
Const MailServerPassWord="888888"        '登录密码
Const MailDomain="163.com"        '域名
dim PageName,PageID,Readme
PageName = request.ServerVariables("SCRIPT_NAME")
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = replace(PageName,".asp",".html")
PageName = replace(PageName,"Show","")
%>