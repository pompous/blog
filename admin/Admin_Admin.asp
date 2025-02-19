<!--#include file="../xwInc/conn.asp"-->
<!--#include file="admin_check.asp"-->
<!--#include file="../xwInc/md5.asp"-->

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
Action=Request("Action")
If Action="Add" Then
    Call Add()
elseif Action="SAdd" Then
    Call SAdd()	
elseif Action="Del" Then
    Call Del()
elseif Action="Edit" Then
   Call Edit()
elseif Action="SEdit" Then
   Call Sedit()
else 
   Call Adm()
End if


Sub SEdit()
username=trim(request("username"))
Admin_Pass=trim(request("Admin_Pass"))
if len(username)<2 then
response.Write("<script>alert(""用户名不能少于2位"");history.back();</script>")
else
set rs=server.CreateObject("ADODB.RECORDSET")
sql="Select top 1 * from xiaowei_Admin"
rs.open sql,conn,1,3
rs("Admin_Name")=username
If Admin_Pass<>"" then
rs("Admin_Pass")=md5(Admin_Pass,16)
end if
rs.update
rs.close
set rs=nothing
Response.Write("<script language=javascript>alert('修改成功!');this.location.href='Admin_admin.asp';</script>")
end if
End Sub 
%>
<% Sub Adm() %>	
<table border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
  <tr>
    <td colspan="6" class="admintitle">管理员列表</td>
  </tr>
  <tr>
    <td width="12%" height="25" align="center" bgcolor="#FFFFFF" class="ButtonList">ID</td>
    <td width="20%" align="center" bgcolor="#FFFFFF" class="ButtonList">管理员名称</td>
    <td width="19%" align="center" bgcolor="#FFFFFF" class="ButtonList">最后登陆时间</td>
    <td width="18%" align="center" bgcolor="#FFFFFF" class="ButtonList">最后登陆IP</td>
    <td width="16%" align="center" bgcolor="#FFFFFF" class="ButtonList">管理选项</td>
  </tr>
<%
set rs=server.CreateObject("ADODB.RECORDSET")
sql="select * from xiaowei_Admin"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("<tr><td colspan=""5""><li>Sorry,当前没有管理员...</li></td></tr>")
else
do while not rs.eof 
%>
  <tr>
    <td height="25" align="center" bgcolor="f7f7f7"><%=rs("id")%></td>
    <td align="center" bgcolor="f7f7f7"><%=rs("Admin_Name")%></td>
    <td align="center" bgcolor="f7f7f7"><%if rs("Admin_Time")<>"" then response.Write(""&rs("Admin_Time")&"") else response.Write("尚未登陆") end if%></td>
    <td align="center" bgcolor="f7f7f7"><%if rs("Admin_IP")<>"" then response.Write(""&rs("Admin_IP")&"") else response.Write("尚未登陆") end if%></td>
	<td align="center" bgcolor="f7f7f7"><a href="?Action=Edit&id=<%=rs("id")%>">编辑</a></td>
  </tr>
  <%
  rs.movenext
  loop
  end if
  rs.close
  set rs=nothing
  %>
</table>
<% End Sub%>

<% Sub Edit()
set rs=server.CreateObject("ADODB.RECORDSET")
sql="select * from xiaowei_Admin"
rs.open sql,conn,1,1
 %>
<table border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
  <tr>
    <td colspan="2" class="admintitle"> 修改管理员资料</td>
  </tr>
  <form action="?Action=SEdit" method="post">
  <tr>
    <td width="20%" height="25" bgcolor="f7f7f7">&nbsp;用户名称：</td>
    <td height="25" bgcolor="f7f7f7"><input name="username" value="<%=rs("Admin_Name")%>" type="text" size="30"></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7">&nbsp;用户密码：</td>
    <td height="25" bgcolor="f7f7f7"><input name="Admin_Pass"type="text" size="30"></td>
  </tr>
  <tr>
    <td height="25" colspan="2" align="center" class="tabletd2"><input type="submit" name="Submit" value="确定修改"></td>
  </tr>
  </form>
</table>
<% 
rs.close
set rs=nothing
End Sub %>

				
		</td>
	</tr>
</table>


</body>

</html>