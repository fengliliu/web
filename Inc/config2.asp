<%
Const EnterpriseMail="http://mail.163.com"        '��ҵ�ʾ�
Const LogoUrl="images/logo.gif"        'Logo��ַ
Const BannerUrl="Images/0791idc.swf"        'Banner��ַ
Const IsFlash="Yes"        '�Ƿ�ΪFLASH
Const width=""        '���
Const height=""        '�߶�
Const WebmasterName="weiqun"        'վ������
Const WebmasterEmail="southidc2008@163.com"        'վ������
Const ComPic_count="20"        '��ҳ����ͼ��
Const New_count=20        '����������Ѷ����
Const MaxPerPage=20        '��������ҳÿҳ������
Const EnableProductCheck=""        '�Ƿ�����������˹���
Const EnableUploadFile="Yes"        '�Ƿ񿪷��ļ��ϴ�
Const MaxFileSize=200        '�ϴ��ļ���С����
Const SaveUpFilesPath="UploadFiles"        '����ϴ��ļ���Ŀ¼
Const UpFileType="gif|jpg|bmp|png|swf|doc|rar"        '������ϴ��ļ�����
Const DelUpFiles="Yes"        'ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ�
Const SessionTimeout=30        'Session�Ự�ı���ʱ��
Const MailObject="Jmail"        '�ʼ��������
Const MailServer="smtp.163.com"        '���������ʼ���SMTP������
Const MailServerUserName="feng"        '��¼�û���
Const MailServerPassWord="888888"        '��¼����
Const MailDomain="163.com"        '����
dim PageName,PageID,Readme
PageName = request.ServerVariables("SCRIPT_NAME")
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = right(PageName,len(PageName)-instrrev(PageName,"/"))
PageName = replace(PageName,".asp",".html")
PageName = replace(PageName,"Show","")
%>