<!--#include file="../xwInc/conn.asp"-->
<!--#include file="admin_check.asp"-->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>系统管理</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
</head>

<body topmargin="0" leftmargin="0">

<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">　</td>
		<td valign="top"><br>
		
		
		
<%
Select Case Request("Sub")
Case "Logout"
  session("xiaoweiAdmin")  =""
  Response.Cookies("xiaoweimanage")("UserName") = ""
  response.Redirect "Admin_Login.asp"
Case "delmdb"
  call BackupData()
  set rs=conn.execute("delete from xiaowei_Class")
  set rs=conn.execute("delete from xiaowei_Article")
  set rs=conn.execute("delete from xiaowei_2weima")
  set rs=conn.execute("delete from xiaowei_GuestBook")
  set rs=conn.execute("delete from xiaowei_link")
  set rs=conn.execute("delete from xiaowei_User")
  set rs=conn.execute("delete from xiaowei_Bots")
	Call Alert ("初始化完成!",-1)
Case ""
%>
<table width="98%" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td bgcolor="#ffffff">
          <table border="0" align="center" cellpadding="3" cellspacing="1" class="admintable1">
            <tr>
              <td align="left" class="admintitle" colspan="4"><img src="images/311.gif" width="16" height="16" /> 系统信息</td>
            </tr>
            <tr>
              <td width="33%" align="left" bgcolor="#FFFFFF" style="height:30px;">文章总数：<%If Mydb("Select Count([ID]) From [xiaowei_Article] Where yn=1",1)(0)>0 then%><font color="blue"><b><%=Mydb("Select Count([ID]) From [xiaowei_Article] Where yn=1",1)(0)%></b></font>/<%end if%><font color="red"><%=Mydb("Select Count([ID]) From [xiaowei_Article]",1)(0)%></font> <font color="blue">[<a href="Admin_Article.asp">管理</a>]</font></td>
            </tr>


            </tr>

          </table>




 <table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable1" style="margin-top:5px;">
            <tr>
              <td align="left" class="admintitle">网站初始化</td>
            </tr>
            <tr>
              <td height="50" bgcolor="#FFFFFF" style="text-align:left;line-height:40px;">
                <form id="form1" name="form1" method="post" action="index.asp?Sub=delmdb">
                <font color="red"><b>警告：此功能会清空网站栏目、文章、评论、留言、链接，你确定这么做吗？</b></font>
                    <input type="submit" name="button" id="button" value="确定初始化数据库" onclick="JavaScript:return   confirm('真的要清？不可恢复的哦!')" style="background:#ffffff;"/>
                </form></td>
            </tr>
          </table>


     </td>
      </tr>
    </table></td>
  </tr>
</table>
<%
sub BackupData()
	dim Dbpath,bkfolder,bkdbname,fso
	Dbpath=SitePath&"data/"&DataName
	Dbpath=server.mappath(Dbpath)
	bkfolder="../Data/bak/"
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& ""&FormatDate(Now,12)&""&".mdb"
		else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& ""&FormatDate(Now,12)&""&".mdb"
		end if
	'Else
		
	End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
    dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f,fso1
    Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>

<%End Select%>
		
		
		</td>
	</tr>
</table>

</body>

</html>