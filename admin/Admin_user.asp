<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->
<!--#include file="../xwInc/md5.asp"-->
<!--#include file="../xwInc/Function_Page.asp"-->

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
			
<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_User.asp">
      <td height="25" bgcolor="f7f7f7">快速查找：
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value=""<%If request("s")="" then Response.Write(" selected") end if%>>-=请选择=-</OPTION>
        <OPTION value="?s=all"<%If request("s")="all" then Response.Write(" selected") end if%>>所有用户</OPTION>
        <OPTION value="?s=yn1"<%If request("s")="yn1" then Response.Write(" selected") end if%>>已审的用户</OPTION>
        <OPTION value="?s=yn0"<%If request("s")="yn0" then Response.Write(" selected") end if%>>未审的用户</OPTION>
        <OPTION value="?s=2"<%If request("s")="2" then Response.Write(" selected") end if%>>24小时登录用户</OPTION>
        <OPTION value="?s=1"<%If request("s")="1" then Response.Write(" selected") end if%>>24小时注册用户</OPTION>
      </SELECT>      </td>
      <td bgcolor="f7f7f7">
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>">
        <input type="submit" name="Submit2" value="搜索"></td>
      </form>
    </tr>
</table>
<%
	if request("action") = "add" then 
		call add()
	elseif request("action")="edit" then
		call edit()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="del" then
		call del()
	elseif request("action")="updateusergroup" then
		call updateusergroup()
	elseif request("action")="deljf" then
		call deljf()
	elseif request("action")="isyn" then
		call isyn()
	else
		call List()
	end if
 
sub List()
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
<tr> 
  <td colspan="9" align=left class="admintitle">用户列表</td>
</tr>
  <tr align="center"> 
    <td width="15%" class="ButtonList">用户名</td>
    <td width="5%" class="ButtonList">积分</td>
    <td width="12%" class="ButtonList">等级</td>
    <td width="4%" class="ButtonList">性别</td>
    <td width="10%" class="ButtonList">籍贯</td>
    <td width="14%" class="ButtonList">注册时间</td>
    <td width="13%" class="ButtonList">注册ＩＰ</td>
    <td width="14%" class="ButtonList">操 作</td>
  </tr>
<%
page=request("page")
Hits=request("hits")
s=Request("s")
Articleclass=request("ClassID")
keyword=request("keyword")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from xiaowei_User"
	if s="yn0" then
	mysql=mysql&" Where yn=0"
	elseif s="yn1" then
	mysql=mysql&" Where yn=1"
	elseif s="1" then
	mysql=mysql&" Where datediff('h',RegTime,Now()) <= 24"
	elseif s="2" then
	mysql=mysql&" Where datediff('h',LastTime,Now()) <= 24"
	elseif s="vip" then
	mysql=mysql&" Where usergroupid=30"
	elseif keyword<>"" then
	mysql=mysql&" Where UserName like '%"&keyword&"%'"
	End if
mysql=mysql&" order by "
mysql=mysql&"ID desc"
mypage.getsql=mysql
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=rs("UserName")%> (<font color=red> <%=Mydb("Select Count([ID]) From [xiaowei_Article] Where UserName='"&rs("UserName")&"'",1)(0)%> </font>)</td>
    <td height="25" align="center" class="tdleft"><%=rs("UserMoney")%></td>
    <td height="25" align="center" class="tdleft"><%=rs("dengji")%></td>
    <td height="25" align="center"><%If rs("Sex")=1 then Response.Write("男") else Response.Write("女") end if%></td>
	<td align="center"><%=rs("province")%><%=rs("city")%></td>
    <td align="center"><%=rs("regtime")%></td>
    <td align="center"><u><%=rs("IP")%></u></td>
    <td width="11%" align="center">
			<%
			Response.Write "<a href='?action=isyn&yn="&rs("yn")&"&ID=" & rs("ID") & "&page="&request("page")&"'>"
            If rs("yn")=0 Then
               Response.Write "<font color=red>未审</font>"
            Else
               Response.Write "已审"
            End If
            Response.Write "</a>"
           %>|<a href="?action=edit&id=<%=rs("ID")%>">编辑</a>|<a href="?action=del&id=<%=rs("ID")%>&UserName=<%=rs("UserName")%>" onClick="JavaScript:return confirm('确认删除吗？这将会连同该用户发表的文章一起删除！')">删除</a></td>
  </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
<tr><td colspan=8 class=td>
<div class="movies end">
<table border="0" cellspacing="5" cellpadding="2" align="center"><tr><%=mypage.showpage()%></tr></table>
</div>
</td>
</tr>
</table>
<%
	rs.close
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
  <tr>
    <td colspan="2" align=left class="admintitle">清空用户积分</td>
  </tr>
  <tr>
    <td height="50">
      <form name="form1" method="post" action="?action=deljf">
        <input type="submit" name="button" id="button" value="用户积分归零"  onClick="JavaScript:return confirm('确认清空吗？不可恢复！')">
      请慎重操作：此操作将导致所有用户积分归零并不可恢复!
      </form>
    </td>
  </tr>
</table>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
  <tr>
    <td colspan="2" align=left class="admintitle">更新用户等级</td>
  </tr>
  <tr>
    <td height="50">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <FORM METHOD=POST ACTION="?action=updateusergroup">
<tr>
<td width="20%" height="30" class="forumrow">更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照当前论坛数据库用户积分和等级设置重新计算用户等级。</td>
</tr>
<tr>
<td width="20%" height="30" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" height="30" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
</form>
    </table></td>
  </tr>
</table>
<%
end sub

sub add()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="5" class="admintitle">添加栏目</th></tr>
<form action="?action=savenew" method=post>
<tr>
<td width="20%" class=b1_1>栏目名称</td>
<td class=b1_1 colspan=4><input type="text" name="ClassName" size="30"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>排　　序</td>
<td colspan=4 class=b1_1><input name="num" type="text" value="10" size="4" maxlength="2"></td>
</tr>
<tr>
  <td class=b1_1>栏目介绍</td>
  <td colspan=4 class=b1_1><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"></textarea></td>
</tr>
<tr>
  <td class=b1_1>导航栏是否显示</td>
  <td colspan=4 class=b1_1><input name="IsMenu" type="radio" class="noborder" value="1" checked>
    是
      <input name="IsMenu" type="radio" class="noborder" value="0">
      否</td>
</tr>
<tr>
  <td class=b1_1>首页是否显示</td>
  <td colspan=4 class=b1_1><input name="IsIndex" type="radio" class="noborder" value="1" checked>
是
  <input name="IsIndex" type="radio" class="noborder" value="0">
否</td>
</tr>
<tr>
  <td class=b1_1>首页显示数量</td>
  <td colspan=4 class=b1_1><input name="IndexNum" type="text" id="IndexNum" value="10" size="4" maxlength="2"></td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td class=b1_1 colspan=4><input type="submit" name="Submit" value="添 加"></td>
</tr></form>
</table>
<%
end sub

sub del()
	id=request("id")
	set rs=conn.execute("delete from xiaowei_User where id="&id)
	set rs=conn.execute("delete from xiaowei_Article where UserID="&id)
	Call Alert ("删除成功","Admin_User.asp")
end sub

sub isyn()
	id=request("id")
	yn=request("yn")
	page=request("page")
	If yn=1 then
	yn=0
	else
	yn=1
	End if
	set rs=conn.execute("Update [xiaowei_User] set yn="&yn&" Where ID="&id)
	Response.Redirect "Admin_User.asp?page="&page&""
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_User where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" method=post>
<tr> 
    <td colspan="3" class="admintitle">修改会员</td>
</tr>
<tr> 
<td width="19%" class="b1_1">会员登录名</td>
<td colspan=2 class=b1_1><%=rs("UserName")%></td>
</tr>
<tr>
  <td class="b1_1">用户组</td>
  <td colspan=2 class=b1_1><select size=1 name="usergroupid">
<%
set trs=conn.Execute("select UserGroupID,GroupName from xiaowei_UserGroup where usermoney=-1 order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&trs(1)
response.write "</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select></td>
</tr>
<tr>
  <td class="b1_1">密码</td>
  <td colspan=2 class=b1_1><input name="PassWord" type="text" id="PassWord" size="30">
    *不修改请留空,原密码:<%=rs("PassWord")%></td>
</tr>
<tr>
  <td class="b1_1">积分</td>
  <td colspan=2 class=b1_1><input name="UserMoney" type="text" id="UserMoney" value="<%=rs("UserMoney")%>" size="30"></td>
</tr>
<tr>
  <td class="b1_1">用户等级</td>
  <td colspan=2 class=b1_1><select name="dengji" size=1 id="dengji">
<%
set rsw=conn.Execute("select GroupName,GroupPic from xiaowei_UserGroup order by Usermoney asc")
do while not rsw.eof
response.write "<option value="&rsw(0)&""
If rs("dengji")=rsw(0) then
response.write " selected"
End if
response.write ">"&rsw(0)&"</option>"
rsw.movenext
loop
rsw.close
set rsw=nothing
%>
</select></td>
</tr>
<tr>
  <td class="b1_1">等级图片</td>
  <td class=b1_1 width="9%"><input name="dengjipic" type="text" id="dengjipic" value="<%=rs("dengjipic")%>"> </td>
  <td class=b1_1 width="70%"><img src="../xwimg/level/<%=rs("dengjipic")%>"></td>
</tr>
<tr>
  <td class="b1_1">用户邮箱</td>
  <td colspan=2 class=b1_1><input name="Email" type="text" id="Email" value="<%=rs("Email")%>" size="30"></td>
</tr>
<tr>
  <td class=b1_1>个人主页</td>
  <td colspan=2 class=b1_1><input name="Birthday" type="text" id="Birthday" value="<%=rs("Birthday")%>" size="30">&nbsp;&nbsp;&nbsp;
	<input name="wangz" type="text" id="wangz" value="<%=rs("wangz")%>" size="30"></td>
</tr>
<tr>
  <td class=b1_1>QQ</td>
  <td colspan=2 class=b1_1><input name="UserQQ" type="text" id="UserQQ" value="<%=rs("UserQQ")%>" size="30"></td>
</tr>
<tr> 
<td width="19%" class="b1_1"></td>
<td class=b1_1 colspan=2><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="提 交"></td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	Dim id
	id=request.form("id")
	'UserName=trim(request.form("UserName"))
	PassWord=trim(request.form("PassWord"))
	Email=trim(request.form("Email"))
	Birthday=trim(request.form("Birthday"))
	wangz=trim(request.form("wangz"))
	UserQQ=trim(request.form("UserQQ"))
	UserMoney=trim(request.form("UserMoney"))
	dengji=trim(request.form("dengji"))
	dengjipic=trim(request.form("dengjipic"))
	usergroupid=trim(request.form("usergroupid"))
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_User where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		'rs("UserName")		= UserName
		rs("Email")			= Email
		If birthday<>"" then
		rs("Birthday")		=Birthday
		end if
		rs("UserQQ")		=UserQQ
		If PassWord<>"" then
		rs("PassWord")		=md5(PassWord,16)
		end if
		rs("UserMoney")		=UserMoney
		rs("dengji")		=dengji
		rs("dengjipic")		=dengjipic
		rs("wangz")		    =wangz
		rs("usergroupid")	=usergroupid
		
		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_User.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_User.asp"";</script>"
	end if
	rs.close
end sub

sub deljf()
	set rs=conn.execute("update xiaowei_User set UserMoney = 0")
	Call Alert ("积分已归零！","admin_user.asp")
end sub

Sub updateusergroup()
if not isnumeric(request.form("beginid")) then
	Call Alert ("开始ID错误","-1")
end if
if not isnumeric(request.form("endid")) then
	Call Alert ("结束ID错误","-1")
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	Call Alert ("开始ID不能比结束ID小","-1")
end if
dim oldMinArticle
oldMinArticle=0

dim rss
set rss=conn.Execute("select id from [xiaowei_User] where id>="&request.form("beginid"))
if rss.eof and rss.bof then
	Call Alert ("已经更新完了","admin_user.asp")
end if
rss.close
set rss=nothing

set rs=conn.Execute("select * from xiaowei_UserGroup Where IsSetting = 0 order by UserMoney desc")
do while not rs.eof
	conn.Execute("update [xiaowei_User] set dengji='"&rs("GroupName")&"',dengjipic='"&rs("grouppic")&"' where (id>="&request.form("beginid")&" and id<="&request.form("endid")&") and (usermoney<"&oldMinArticle&" and usermoney>="&rs("UserMoney")&") and usergroupid=3")
	oldMinArticle=rs("UserMoney")
rs.movenext
loop
rs.close
set rs=nothing

%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
  <tr>
    <td colspan="2" align=left class="admintitle">继续更新用户数据</td>
  </tr>
  <tr>
    <td height="50"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <FORM METHOD=POST ACTION="?action=updateusergroup">
<tr>
<td width="20%" height="30" class="forumrow">更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照当前论坛数据库用户积分和等级设置重新计算用户等级。</td>
</tr>
<tr>
<td width="20%" height="30" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" height="30" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
      </form>
    </table></td>
  </tr>
</table>
<%
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>