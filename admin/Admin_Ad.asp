
<!--#include file="../xwInc/conn.asp"-->
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
	else
		call List()
	end if
 
sub List()
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not IsNumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Ad order by ID desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		call add()
	else
%>
<table border="0"  align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="5" align=left class="admintitle">广告列表　[<a href="?action=add">添加</a>]</td>
</tr>
  <tr align="center"> 
    <td width="30%" class="ButtonList">广告名称</td>
    <td width="5%" class="ButtonList">ID号</td>
    <td width="30%" class="ButtonList">说明</td>
    <td width="20%" class="ButtonList">最后修改时间</td>
    <td width="15%" class="ButtonList">操 作</td>
  </tr>
<%
		rs.PageSize = 20
		NoI=0
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = 20)
		NoI=NoI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=NoI%> .<%=rs("Title")%></td>
    <td height="25" align="center"><%=rs("ID")%></td>
    <td height="25" align="center"><%=rs("Note")%></td>
    <td height="25" align="center"><%=rs("DateAndTime")%></td>
    <td align="center"><a href="?action=edit&id=<%=rs("ID")%>">编辑</a></td>
    </tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr>
  <td colspan=5 align=center class=b1_1>广告调用：在需要调用的地方插入 &lt;%Call ShowAD(广告ID号)%&gt; 即可。分页：
    <%Pcount=rs.PageCount
	if currentpage > 4 then
		response.write "<a href=""?page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
		endpage=currentpage+3
	else
		endpage=Pcount
	end if
	dim i
	for i=currentpage-3 to endpage
		if not i<1 then
			if i = clng(currentpage) then
        		response.write " <font color=red>["&i&"]</font>"
			else
        		response.write " <a href=""?page="&i&""">["&i&"]</a>"
			end if
		end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td>
</tr>
</table>
<%
	end if
	rs.close
end sub

sub add()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="5" class="admintitle">添加广告</th></tr>
<form action="?action=savenew" method=post>
<tr>
<td width="20%" class=b1_1>广告名称</td>
<td class=b1_1 colspan=4><input name="Title" type="text" id="Title" size="40" maxlength="20"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>广告代码</td>
<td colspan=4 class=b1_1><textarea name="Content" cols="80" rows="15" id="Content"></textarea></td>
</tr>
<tr>
  <td class=b1_1>广告说明</td>
  <td colspan=4 class=b1_1><input name="Note" type="text" id="Note" value="" size="40"></td>
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
	set rs=conn.execute("delete from xiaowei_Ad where id="&id)
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_AD.asp"";</script>"
end sub

sub savenew()
	if trim(request.form("Title"))="" or request.form("Content")="" then
		Response.write"<script>alert(""请填写广告名称及广告代码!"");location.href=""?action=add"";</script>"
		Response.End
	end if
	Title		=trim(request.form("Title"))
	Content		=trim(request.form("Content"))
	Note		=trim(request.form("Note"))
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Ad where Title='"& Title &"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("Title")			=Title
		rs("Content")		=Content
		rs("Note")			=Note
		rs.update
		Response.write"<script>alert(""恭喜,添加成功！"");location.href=""Admin_AD.asp"";</script>"
	else
		Response.write"<script>alert(""添加失败，该广告名称已经存在！"");location.href=""Admin_AD.asp"";</script>"
end if
	rs.close
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_Ad where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">修改广告</td>
</tr>
<tr>
  <td width="20%" class=b1_1>广告名称</td> 
<td colspan=4 class=b1_1><input name="Title" type="text" value="<%=rs("Title")%>" size="40" maxlength="20"></td>
</tr>
<tr>
  <td width="20%" class=b1_1>广告代码</td>
  <td colspan=4 class=b1_1><textarea name="Content" cols="80" rows="15" id="Content"><%=rs("Content")%></textarea></td>
</tr>
<tr>
  <td class=b1_1>广告说明</td>
  <td colspan=4 class=b1_1><input name="Note" type="text" id="Note" value="<%=rs("Note")%>" size="40"></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1 colspan=4><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="提 交"></td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	Dim Title
	ID			=trim(request.form("ID"))
	Title		=trim(request.form("Title"))
	Content		=trim(request.form("Content"))
	Note		=trim(request.form("Note"))
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Ad where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("Title")			=Title
		rs("Content")		=Content
		rs("Note")			=Note
		rs("DateAndTime")	=Now()
		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_AD.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_AD.asp"";</script>"
	end if
	rs.close
end sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>