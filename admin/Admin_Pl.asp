<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwInc/Function_Page.asp"-->
<!--#include file="Admin_check.asp"-->

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
    <tr><form name="form1" method="get" action="Admin_pl.asp">
      <td height="25" bgcolor="f7f7f7">快速查找：
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value="" selected>-=请选择=-</OPTION>
        <OPTION value="?s=all">所有评论</OPTION>
        <OPTION value="?s=yn1">已审的评论</OPTION>
        <OPTION value="?s=yn0">未审的评论</OPTION>
      </SELECT>      </td>
      <td bgcolor="f7f7f7">
        <a href="?hits=1"></a>
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>">
        <input type="submit" name="Submit2" value="搜索">
        </td>
      
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
	elseif request("action")="delAll" then
		call delAll()
	else
		call List()
	end if

sub List()
%>
<form name="myform" method="POST" action="Admin_Pl.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="6" align=left class="admintitle">评论列表</td>
</tr>
    <tr bgcolor="#f1f3f5" style="font-weight:bold;">
    <td width="5%" height="30" align="center" class="ButtonList">　</td>
    <td width="47%" align="center" class="ButtonList">评论内容</td>
    <td width="17%" align="center" class="ButtonList">发布人</td>
    <td width="16%" height="25" align="center" class="ButtonList">发布时间</td>
    <td width="15%" height="25" align="center" class="ButtonList">管理</td>    
    </tr>
<%
page=request("page")
s=Request("s")
id=request("id")
keyword=request("keyword")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from xiaowei_Pl"
	if id<>"" then
	mysql=mysql&" Where ArticleID="&id&""
	elseif s="yn0" then
	mysql=mysql&" Where yn=0"
	elseif s="yn1" then
	mysql=mysql&" Where yn=1"
	elseif keyword<>"" then
	mysql=mysql&" Where memContent like '%"&keyword&"%'"
	End if
mysql=mysql&" order by id desc"
mypage.getsql=mysql
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
    <tr>
    <td height="25" align="center" bgcolor="f7f7f7"><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td height="25" bgcolor="f7f7f7"><a href="<%=SitePath%>xwArticle/?<%=rs("ArticleID")%>.html" target="_blank"><%=left(GlHtml(rs("memContent")),30)%>...</a></td>
    <td height="25" align="center" bgcolor="f7f7f7"><%=rs("memAuthor")%></td>
    <td height="25" align="center" bgcolor="f7f7f7"><span class="td"><%=rs("PostTime")%></span></td>
    <td align="center" bgcolor="f7f7f7"><%if rs("yn")=0 then Response.Write("<font color=red>未审</font>") else Response.Write("已审") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">回复</a>|<a href="?action=del&id=<%=rs("ID")%>">删除</a></td>
    </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
<tr><td align="center" bgcolor="f7f7f7"><input name="Action" type="hidden"  value="Del"><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="5" bgcolor="f7f7f7"><input type="submit" value="删除" name="Del" id="Del">
    <input type="submit" value="批量未审" name="Del" id="Del">
    <input type="submit" value="批量审核" name="Del" id="Del"></td>
  </tr><tr><td bgcolor="f7f7f7" colspan="6">
<div class="movies end">
<table border="0" cellspacing="5" cellpadding="2" align="center"><tr><%=mypage.showpage()%></tr></table>
</div>
</td></tr></table>
</form>
<%
	rs.close
end sub

sub del()
	id=request("id")
	set rs=conn.execute("delete from xiaowei_Pl where id="&id)
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Pl.asp"";</script>"
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_Pl where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">修改评论</td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">评论人</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><input name="memAuthor" type="text" class="inputbg" id="memAuthor" value="<%=rs("memAuthor")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">IP</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><%=rs("IP")%>　评论时间：<%=rs("PostTime")%></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">内容</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><textarea name="memContent" cols="80" rows="10" id="memContent"><%=rs("memContent")%></textarea></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">回复</td>
  <td colspan=4 bgcolor="#f7f7f7" class=td><textarea name="reContent" cols="80" rows="10" id="reContent"><%If rs("reContent")<>"" then%><%=rs("reContent")%><%else%><font color=red><b>管理员回复：</b></font><%End if%></textarea></td>
</tr>
<tr> 
<td width="20%"></td>
<td colspan=4 class=td><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="提 交">  </td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	id=request.form("id")
	memAuthor			=trim(request.form("memAuthor"))
	memContent				=request.form("memContent")
	reContent				=request.form("reContent")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Pl where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("memAuthor")			=memAuthor
		rs("memContent")		=memContent
		if reContent<>"<font color=red><b>管理员回复：</b></font>" then
		rs("reContent")			=reContent
		end if

		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_Pl.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_Pl.asp"";</script>"
	end if
	rs.close
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('请选择!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="批量未审" Then
   set rs=conn.execute("update xiaowei_Pl set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_Pl.asp"";</script>")
ElseIf Request("Del")="批量审核" Then
   set rs=conn.execute("update xiaowei_Pl set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_Pl.asp"";</script>")
ElseIf Request("Del")="删除" Then
	set rs=conn.execute("delete from xiaowei_Pl where ID In(" & ID & ")")
   	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Pl.asp"";</script>"
End If
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>