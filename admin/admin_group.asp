<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwInc/md5.asp"-->
<!--#include file="admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS 系统管理</title>
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

		if request("action")="save" then
		call savegrade()
		elseif request("action")="add" then
		call add()
		elseif request("action")="savenew" then
		call savenew()
		elseif request("action")="del" then
		call del()
		elseif request("action")="per" then
		call per()
		else
		call gradeinfo()
		end if
sub gradeinfo()
%>
<form method="POST" action=admin_group.asp?action=save>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
<th height="23" colspan="6" class="admintitle" >用户等级设定[<a href="?action=add">添加</a>]</th>
</tr>
<tr> 
<td width="22%" height="30" align="center" bgcolor="#f7f7f7" class="ButtonList"><B>等级名称</B></td>
<td width="15%" bgcolor="#f7f7f7" class="ButtonList"><B>最少积分</B></td>
<td width="21%" bgcolor="#f7f7f7" class="ButtonList"><B>前台显示颜色</B></td>
<td width="30%" bgcolor="#f7f7f7" class="ButtonList"><B>图片</B></td>
<td width="12%" bgcolor="#f7f7f7" class="ButtonList"><B>操作</B></td>
</tr>
<%
Dim TempArray,DefaultLock
Set Rs=conn.Execute("select UserGroupID,GroupName from xiaowei_UserGroup where issetting=1 And ParentGID=0 order by UserGroupID")
TempArray = Rs.GetRows(-1)
set rs=conn.Execute("select * from xiaowei_UserGroup order by Usermoney asc,UserGroupID asc")
do while not rs.eof
	If Rs("ParentGID")=0 Then 
		DefaultLock="1"
	Else
	DefaultLock=""
	End If
	%>
	<input type=hidden value="<%=rs("UserGroupID")%>" name="GroupNameid">
	<tr> 
	<td align="center" bgcolor="#f7f7f7" class=Forumrow><input size=15 value="<%=rs("GroupName")%>" name="GroupName" type="text"<%If rs("GroupName")="管理员" then Response.Write( " readOnly") end if%>>
	  <input name="groupid" type="hidden" id="groupid" value="<%=rs("ParentGID")%>"></td>
	<td align="center" bgcolor="#f7f7f7" class=Forumrow>
	<%If DefaultLock <>"1" Then %>
		<input size=5 value="<%=rs("Usermoney")%>" name="Usermoney" type=text >
	<%Else%>
		<input type=hidden   value="<%=rs("Usermoney")%>" name="Usermoney"  >
		<%=rs("Usermoney")%>
	<%End If%>	</td>
	<td align="center" bgcolor="#f7f7f7" class=Forumrow><input name="color" type="text" id="color" value="<%=rs("color")%>" size=15></td>
	<td bgcolor="#f7f7f7" class=Forumrow><input size=15 value="<%=rs("grouppic")%>" name="dengjipic" type=text><img src="../xwimg/level/<%=rs("GroupPic")%>"></td>
	<td align="center" bgcolor="#f7f7f7" class=Forumrow><%If Rs("UserGroupID")>4 Then%><a href="?action=del&id=<%=rs("UserGroupID")%>">删除</a><%End If%></td>
	</tr>
	<%
	rs.movenext
	loop
rs.close
set rs=nothing
%>
<tr> 
<td colspan=6 align="center" bgcolor="#f7f7f7" class=Forumrow> 
<input type="submit" name="Submit" value=" 编 辑 "></td>
</tr>
</table>
</form>
<%
end sub

Sub savegrade()
	Server.ScriptTimeout=99999999
	Dim GroupNameid,iuserclass,GroupName,Usermoney,dengjipic,groupid
	For i=1 to request.form("GroupNameid").count
		GroupNameid=replace(request.form("GroupNameid")(i),"'","")
		GroupName=replace(request.form("GroupName")(i),"'","")
		Usermoney=replace(request.form("Usermoney")(i),"'","")
		dengjipic=replace(request.form("dengjipic")(i),"'","")
		groupid=replace(request.form("groupid")(i),"'","")
		color=replace(request.form("color")(i),"'","")
		if isnumeric(GroupNameid) and isnumeric(iuserclass) and GroupName<>"" and isnumeric(Usermoney) and dengjipic<>"" and isnumeric(groupID) then
		set rs=conn.Execute("select * from xiaowei_UserGroup where UserGroupID="&GroupNameID)
		if rs("GroupName")<>trim(GroupName) or rs("grouppic")<>trim(dengjipic) or (rs("parentgid")<>cint(groupid) and rs("parentgid")>0) then
			conn.execute("update xiaowei_User set dengji='"&GroupName&"',dengjipic='"&dengjipic&"' Where dengji='"&rs("GroupName")&"'")
		end if
		if rs("parentgid")=0 then groupid=0
		conn.Execute("update xiaowei_UserGroup set GroupName='"&GroupName&"',color='"&color&"',Usermoney="&Usermoney&",grouppic='"&dengjipic&"',parentgid="&groupid&" where usergroupid="&GroupNameID)
		end if
	next
	Call Alert ("设置成功!","admin_group.asp")
	set rs=nothing
End Sub

sub add()
%>
<form method="POST" action=admin_group.asp?action=savenew>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
<th colspan="2" class="admintitle">添加新的用户等级</th>
</tr>
<tr>
<td width="37%" class=forumrow><B>等级名称</B></td>
<td width="63%" class=forumrow><input size=30 name="GroupName" type=text></td>
</tr>
<tr>
<td width="37%" class=forumrow><B>最少积分</B><BR>
  如果该等级是荣誉称号或者管理身份，这里可以填写-1，表示不跟随积分增长而升级</td>
<td width="63%" class=forumrow><input size=30 name="Usermoney" type=text></td>
</tr>
<tr>
<td width="37%" class=forumrow><B>等级图片</B></td>
<td width="63%" class=forumrow><input name="dengjipic" type=text value="level1.gif" size=30>&nbsp;</td>
</tr>
<tr> 
<td colspan=2 class=forumrow> 
  <input type="submit" name="Submit" value="提 交"></td>
</tr>
</table>
</form>
<%
end sub
sub savenew()
Dim GroupTitle,GroupPic
Set rs=Conn.execute("select * from xiaowei_UserGroup")
GroupTitle=rs("GroupName")
GroupPic=rs("GroupPic")
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from xiaowei_UserGroup where GroupName='"&request.form("GroupName")&"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
rs.addnew
rs("GroupName")=request.form("GroupName")
rs("Usermoney")=request.form("Usermoney")
rs("grouppic")=request.form("dengjipic")
rs("parentgid")=3
rs("isdisp")=0
rs("IsSetting")=0
rs.update
else
	Call Alert ("该等级名称已经存在!","-1")
	exit sub
end if
rs.close
set rs=nothing
Call Alert ("添加成功!","Admin_Group.Asp")
end sub

sub del()
Server.ScriptTimeout=99999999
dim Usermoney,minuserclass
if isnumeric(request("id")) then
if CLng(request("id"))<4 then
	Call Alert ("系统默认等级不能删除!","-1")
	exit sub
end if
set rs=Conn.Execute("select * from xiaowei_UserGroup where usergroupid="&request("id"))
Usermoney=rs("Usermoney")
mindengji=rs("GroupName")
if Usermoney=-1 then
set rs=Conn.Execute("select top 1 * from xiaowei_UserGroup order by Usermoney")
else
set rs=Conn.Execute("select top 1 * from xiaowei_UserGroup where Usermoney<"&Usermoney&" and usergroupid="&request("id")&" order by Usermoney desc")
end if
if not (rs.eof and rs.bof) then
	Conn.Execute("update [xiaowei_User] set dengji='"&rs("GroupName")&"',dengjipic='"&rs("grouppic")&"' where dengji='"&minuserclass&"'")
else
	set rs=nothing
	set rs=Conn.Execute("select top 1 * from xiaowei_UserGroup where Usermoney<"&Usermoney&" order by Usermoney desc")
	if not (rs.eof and rs.bof) then
	Conn.Execute("update [xiaowei_User] set dengji='"&rs("GroupName")&"',dengjipic='"&rs("grouppic")&"' where dengji='"&minuserclass&"'")
	end if
end if
Conn.Execute("delete from xiaowei_UserGroup where usergroupid="&request("id"))
Call Alert ("删除成功!","admin_group.asp")
set rs=nothing
End If
End Sub

sub per()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim groupid
	if request("isdefault")=1 then
		set rs=Conn.execute("select * from xiaowei_UserGroup where usergroupid="&request("groupid"))
		If Rs("ParentGID")=0 Then
			Dv_suc("您没有选择自定义等级选项，所做修改将无效")
			Exit Sub
		End If
		if rs("issetting")=1 then
		groupid=rs("parentgid")
		set rs=nothing
		set rs=Conn.execute("select * from xiaowei_UserGroup where usergroupid="&groupid)
		Set Rs=Nothing
		Conn.execute("update xiaowei_UserGroup set issetting=0 where usergroupid="&request("regroupid"))
		'取消自定义设置，更新用户数据，更新为用户组ID
		Conn.execute("update [xiaowei_User] set usergroupid="&groupid&" where dengji='"&request("GroupName")&"'")
		end if
		
	else
		Conn.execute("update xiaowei_UserGroup set issetting=1 where usergroupid="&request("regroupid"))
		'更新用户数据
		Conn.execute("update [xiaowei_User] set usergroupid="&request("regroupid")&" where dengji='"&request("GroupName")&"'")
	End If

	ReloadGroup(request("regroupid"))
	Dv_suc("修改等级自定义权限成功")
else
Dim founduserper,usergrade
If IsNumerIc(request("regroupid")) and request("regroupid")<>"" Then
	Set Rs=Conn.Execute("select * from xiaowei_UserGroup where usergroupid="&request("regroupid"))
	usergrade=rs("GroupName")
End If
founduserper=false
set rs=Conn.Execute("select * from xiaowei_UserGroup where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到用户等级"
exit sub
end if
If Rs("UserGroupID")<9 Then
	founduserper=false
Else
	If Rs("IsSetting")=1 Then
		founduserper=true
	Else
		founduserper=false
	End If
End If

set rs=nothing
end if
end sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>