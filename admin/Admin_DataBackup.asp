<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>�������ݿ�</strong></div>
            <%
if request("action")="Backup" then
call backupdata()
else
%></td>
        </tr>
        <tr class="tr_southidc"> 
          <form method="post" action="Admin_DataBackup.asp?action=Backup">
            <td><br> 
              <table width="91%" border="0" align="center" cellpadding="0" cellspacing="2">
                <tr> 
                  <td height="25"><strong>������ҵ��վ����ϵͳ�����ļ�</strong>[��ҪFSOȨ��]</td>
                </tr>
                <tr> 
                  <td height="22"> ��ǰ���ݿ�·��</td>
                </tr>
                <tr> 
                  <td height="22"><input type=text size=50 name=DBpath value="<%=db%>"></td>
                </tr>
                <tr> 
                  <td height="22"><input type="hidden" size=50 name=bkfolder value=Databackup ></td>
                </tr>
                <tr> 
                  <td height="22">�������ݿ�����[�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����]</td>
                </tr>
                <tr> 
                  <td height="22"><input type=text size=30 name=bkDBname value="<%=date()%>"></td>
                </tr>
                <tr> 
                  <td height="22"><div align="center"> 
                      <input type=submit value="ȷ��">
                    </div></td>
                </tr>
                <tr> 
                  <td height="22"><br> <br>
                    �������Ĭ�����ݿ��ļ�Ϊ<%=db%><br>
                    ������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
                    ע�⣺����·��������������ռ��Ŀ¼�����·��</td>
                </tr>
                <tr> 
                  <td height="22">&nbsp;</td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <%end if%>
      <% 
sub backupdata() 
Dbpath=request.form("Dbpath") 
Dbpath=server.mappath(Dbpath) 
bkfolder=request.form("bkfolder") 
bkdbname=request.form("bkdbname") 
Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = True Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname & ".asa"
end if 
response.write "<center>�������ݿ�ɹ������ݵ����ݿ�Ϊ " & bkfolder & "\" & bkdbname & ".asa</center>" 
Else 
response.write "�Ҳ���������Ҫ���ݵ��ļ���" 
End if 
end sub 
'------------------���ĳһĿ¼�Ƿ����------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'���� 
CheckDir = True 
Else 
'������ 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------����ָ����������Ŀ¼--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = True 
Set fso1 = nothing 
End Function 
%></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->