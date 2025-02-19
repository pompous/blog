<!--#include file="../xwInc/conn.asp"-->
<!--#include file="admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<script src="js/admin.js"></script>
<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">　</td>
		<td valign="top"><br>	
			
<%
select case request("action")
    case "SpaceSize"
	    SpaceSize()
	case "CompressData"
		if IsSqlDataBase = 1 then
			SQLUserReadme()
		else
			CompressData()
		end if
	case "BackupData"
	    if request("act")="Backup" then
			call updata()
		else
			call BackupData()
		end if
	case "RestoreData"
			if request("act")="Restore" then
				dim Dbpath,backpath,fso
				Dbpath=request.form("Dbpath")
				backpath=request.form("backpath")
				if dbpath="" then
				Call Alert("请输入您要恢复的数据库路径及名称!","-1")	
				else
				Dbpath=server.mappath(Dbpath)
				end if
				backpath=server.mappath(backpath)
			
				Set Fso=Server.CreateObject("Scripting.FileSystemObject")
				if fso.fileexists(dbpath) then  					
					fso.copyfile Dbpath,Backpath
					Call Alert("成功恢复数据!","Admin_data.asp?action=SpaceSize")	
				else
					Call Alert("备份目录下并无您的备份文件!","-1")	
				end if
			else
				call RestoreData()
			end if
end select
%>
<%
'====================压缩数据库 =========================
sub CompressData()
%><center>
<table border="0"  cellspacing="1" cellpadding="3" height="1" class="admintable1">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="admintitle">压缩数据库</td>
</tr><tr>
<td height=30 bgcolor="#FFFFFF" class="td"><b>注意：</b>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库可能会压缩失败，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td height="30" bgcolor="#FFFFFF" class="td">压缩数据库：<input name="dbpath" type="text" value="../xwData/<%=DataName%>" size="50">
&nbsp;
<input type="submit" value="开始压缩"></td>
</tr>
<tr>
<td height="30" bgcolor="#FFFFFF" class="td"><input name="boolIs97" type="checkbox" class="noborder" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">注：请尽量用ftp下载回数据库后压缩，以免出错！如果你非要使用此功能，请备份后再压缩！<strong>数据库出错本程序作者概不负责!</strong></td>
</tr>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
	Dim fso, Engine, strDBPath,JET_3X
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	If fso.FileExists(dbPath) Then
		fso.CopyFile dbpath,strDBPath & "temp.mdb"
		Set Engine = CreateObject("JRO.JetEngine")
	
		If boolIs97 = "True" Then
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
			& "Jet OLEDB:Engine Type=" & JET_3X
		Else
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		End If
	
		fso.CopyFile strDBPath & "temp1.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		fso.DeleteFile(strDBPath & "temp1.mdb")
		Set fso = nothing
		Set Engine = nothing
		
		Call Alert("压缩成功!","Admin_data.asp?action=SpaceSize")
	Else
		Call Alert("数据库名称或路径不正确. 请重试!","-1")
	End If
End Function
%>
<%
'====================备份数据库=========================
sub BackupData()
%><center>
	<table border="0"  cellspacing="1" cellpadding="3" class="admintable1">
	  <tr>
		  <td colspan="2" class="admintitle" >备份网站系统数据( 需要FSO支持，FSO相关帮助请看微软网站 )</td>
	  	</tr>
  				<form method="post" action="Admin_data.asp?action=BackupData&act=Backup">
  				<tr>
  				  <td width="19%" height="30" align="center" bgcolor="#FFFFFF" class="td">当前数据库路径(相对)：</td>
				  <td width="81%" bgcolor="#FFFFFF" class="td" align="left"><input name=DBpath type=text id="DBpath" value="../xwData/<%=DataName%>" size="40" /></td>
  				</tr>
  				<tr>
  				  <td height="30" align="center" bgcolor="#FFFFFF" class="td">备份数据库目录(相对)：</td>
				  <td bgcolor="#FFFFFF" class="td" align="left"><input name=bkfolder type=text value="../xwdata/bak/" size="40" Readonly="true"/>
&nbsp;如目录不存在，程序将自动创建</td>
  				</tr>
  				<tr>
  				  <td height="30" align="center" bgcolor="#FFFFFF" class="td">备份数据库名称(名称)：</td>
				  <td bgcolor="#FFFFFF" class="td" align="left"><input name=bkDBname type=text value="<%=FormatDate(Now(),12)%>.mdb" size="40" />
&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建</td>
  				</tr>
  				<tr>
  				  <td height="30" bgcolor="#FFFFFF" class="td">　</td>
				  <td bgcolor="#FFFFFF" class="td" align="left"><input type=submit value="确定" /></td>
  				</tr>
  				<tr>
  				  <td height="30" bgcolor="#FFFFFF" class="td">　</td>
				  <td bgcolor="#FFFFFF" class="td" align="left">注意：所有路径都是相对与程序空间根目录的相对路径 </td>
  				</tr>	
  				</form>
  </table>
<%
end sub

sub updata()
	dim Dbpath,bkfolder,bkdbname,fso
	Dbpath=request.form("Dbpath")
	Dbpath=server.mappath(Dbpath)
	bkfolder=request.form("bkfolder")
	bkdbname=request.form("bkdbname")
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		end if		
		Call Alert ("备份数据库成功!","Admin_data.asp?action=SpaceSize")
	Else
		Call Alert ("数据库路径错误","-1")
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
<%
'====================恢复数据库=========================
sub RestoreData()
%><center>
<table border="0"  cellspacing="1" cellpadding="3" class="admintable1">
	<tr>
		<td colspan="2" class="admintitle">恢复网站系统数据( 需要FSO支持，FSO相关帮助请看微软网站 )</td>
    </tr>
<form method="post" action="Admin_data.asp?action=RestoreData&act=Restore">
<tr>
  <td width="19%" height="30" align="center" bgcolor="#FFFFFF" class="td">备份数据库路径(相对)：</td>
  <td width="81%" bgcolor="#FFFFFF" class="td" align="left"><input name=DBpath type=text id="DBpath" value="../xwdata/bak/<%=year(Now)&"-"&month(Now)&"-"&day(Now)%>.MDB" size=40 /> 
  找到的备份数据库：<%=FileList("../xwdata/bak","mdb")%></td>
</tr>
<tr>
  <td height="30" align="center" bgcolor="#FFFFFF" class="td">目标数据库路径(相对)：</td>
  <td bgcolor="#FFFFFF" class="td" align="left"><input type=text size=40 name=backpath value="../xwdata/<%=DataName%>" /></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">　</td>
  <td bgcolor="#FFFFFF" class="td" align="left"><input type=submit value="恢复数据" /></td>
</tr>
<tr>
  <td height="30" bgcolor="#FFFFFF" class="td">　</td>
  <td bgcolor="#FFFFFF" class="td" align="left">请谨慎操作！</td>
</tr>	
</form>
</table>
<%
end sub
sub SpaceSize()
%><center>
<table height="1" border="0" cellpadding="3"  cellspacing="1" bgcolor="#F2F9E8" class="admintable1">
  <tr>
    <td colspan="2" class="admintitle">程序占用空间情况 </td>
  </tr>
  <tr>
    <td width="19%" height="30" align="center" bgcolor="#FFFFFF">系统占用空间总计：</td>
    <td width="81%" bgcolor="#FFFFFF" align="left">&nbsp;<%allsize()%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">数据库总占用空间：</td>
    <td bgcolor="#FFFFFF" align="left">&nbsp;<%othersize("xwData")%></td>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">系统后台占用空间：</td>
    <td bgcolor="#FFFFFF" align="left">&nbsp;<%othersize("admin")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">备份数据占用空间：</td>
    <td bgcolor="#FFFFFF" align="left">&nbsp;<%othersize("xwdata/bak")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">系统图片占用空间：</td>
    <td bgcolor="#FFFFFF" align="left">&nbsp;<%othersize("images")%></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#FFFFFF">上传目录占用空间：</td>
    <td bgcolor="#FFFFFF" align="left">
	&nbsp;<%othersize("kindeditor/attached/image")%></td>
  </tr>
</table>
<%end sub%>
<%
sub othersize(names)
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("..\Images")
	ml=left(path,(instrrev(path,"\")-1))&"\"&names
	
	On Error Resume Next
	set d=fso.getfolder(ml) 
	If Err Then
		err.Clear
		Response.Write "<font color=red>提示：没有“"&names&"”目录</font>"					
		'Response.End()
	End If
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=formatnumber(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

sub allsize()
	dim fso,path,ml,mlsize,dx,d,size
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	path=server.mappath("../index.asp")
	ml=left(path,(instrrev(path,"\")-1))
	set d=fso.getfolder(ml) 
	mlsize=d.size
	size=mlsize
	dx=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   dx=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   dx=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write dx
end sub

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	drvpathroot=server.mappath("../Images")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	On Error Resume Next
	drvpath=server.mappath("../"&drvpath)
	If Err Then
		err.Clear
		Response.Write "没有名为“"&drvpath&"”的目录，您可以修改文件以正确显示该目录的使用量。"			
		Response.End()
	End If
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 

Function FileList(FolderUrl,FileExName)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileList=""
For Each FileName in file
If Trim(FileExName)<>"" Then
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&""&FileName.Name&"|"
	End If
Else
     FileList=FileList&"<a href='#'>"&FileName.Name&"</a><br>"
End If
Next
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function
%>

				
		</td>
	</tr>
</table>


</body>

</html>