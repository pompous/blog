<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS 系统管理</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
<style>

			form {
				margin: 0;
			}
			textarea {
				display: block;
			}
		</style>
		<link rel="stylesheet" href="../kindeditor/themes/default/default.css" />
		<script charset="utf-8" src="../kindeditor/kindeditor-min.js"></script>
		<script charset="utf-8" src="../kindeditor/lang/zh_CN.js"></script>
		<script>
			var editor;
			KindEditor.ready(function(K) {
				editor = K.create('textarea[name="content"]', {
					resizeType : 1,
					allowPreviewEmoticons : false,
					allowImageUpload : false,
					items : [
						'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
						'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
						'insertunorderedlist', '|', 'emoticons', 'image', 'link'] ,
 //经测试，下面这行代码可有可无，不影响获取textarea的值
 //afterCreate: function(){this.sync();}
 //下面这行代码就是关键的所在，当失去焦点时执行 this.sync();
 afterBlur: function(){this.sync();}
				});
			});
		</script>


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
	elseif request("action")="delAll" then
		call delAll()
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
	sql="select * from xiaowei_GuestBook where  kavkill_ReID=0 order by id desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Response.Write("没有贴子!")
	else
%>
<form name="myform" method="POST" action="Admin_Guestbook.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="4" align=left class="admintitle">贴子列表</td>
</tr>
<%
		rs.PageSize = 15
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		NoI=0
		while (not rs.eof) and (not page_count = 15)
		NoI=NoI+1
%>
    <tr>
    <td width="4%" height="35" align="LEFT" bgcolor="#E6E6E6">&nbsp;<%=NoI%>.&nbsp;&nbsp;<input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td width="46%" bgcolor="#E6E6E6"><%=rs("title")%>&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;<font color="#777777">共有</font><font color="#FF0000"><%=rs("kavkill_Replies")%></font><font color="#777777">条回复</font></td>
    <td width="35%" height="25" align="center" bgcolor="#E6E6E6"><%=rs("UserName")%>(<span class="td"><%=rs("AddTime")%></span>)</td>
    <td width="15%" align="center" bgcolor="#E6E6E6"><%if rs("yn")=0 then Response.Write("<font color=red>未审</font>") else Response.Write("已审") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">回复</a>|<a href="?action=del&id=<%=rs("ID")%>">删除</a></td>
    </tr>
<%
    Nid = RS("ID")
	set rs6 = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_GuestBook where  kavkill_ReID= "&Nid&" order by id desc"
	rs6.open sql,conn,1,1
	do while not (rs6.eof or err) 
	%>

    <tr>
    <td width="4%" height="35" align="RIGHT" bgcolor="f7f7f7"><input type="checkbox" value="<%=rs6("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td width="46%" bgcolor="f7f7f7"><font color="#666666"><%=rs6("title")%></font></td>
    <td width="35%" height="25" align="center" bgcolor="f7f7f7"><%=rs6("UserName")%>(<span class="td"><%=rs6("AddTime")%></span>)</td>
    <td width="15%" align="center" bgcolor="f7f7f7"><%if rs6("yn")=0 then Response.Write("<font color=red>未审</font>") else Response.Write("已审") end if%>|<a href="?action=edit&id=<%=rs6("ID")%>">回复</a>|<a href="?action=del&id=<%=rs6("ID")%>">删除</a></td>
    </tr>
<%

  rs6.movenext
  loop
 
  rs6.close
  set rs6=nothing
%>

<%
		page_count = page_count + 1
		rs.movenext
		wend
%>



<tr><td align=center><input name="Action" type="hidden"  value="Del">
    <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="3">
  	<input type="submit" value="删除" name="Del" id="Del">
    <input type="submit" value="批量未审" name="Del" id="Del">
	<input type="submit" value="批量审核" name="Del" id="Del">
    分页：
    <%Pcount=rs.PageCount
	if currentpage > 4 then
		response.write "<a href=""?page=1&class="&A_Class&""">[1]</a> ..."
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
        		response.write " <a href=""?page="&i&"&class="&A_Class&""">["&i&"]</a>"
			end if
		end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&class="&A_Class&""">["&Pcount&"]</a>"
	end if
%></td>
  </tr>
</table>
</form>
<%
	end if
	rs.close
end sub

sub del()
	id=request("id")
	set rs=conn.execute("delete from xiaowei_GuestBook where id="&id)
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Guestbook.asp"";</script>"
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_GuestBook where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">回复贴子</td>
</tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">标题</td>
  <td colspan=4 class=b1_1><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">发贴人</td>
  <td colspan=4 class=b1_1><input name="username" type="text" id="title2" value="<%=rs("username")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">IP</td>
  <td colspan=4 class=b1_1><%=rs("AddIP")%>[<a href="<%=rs("AddIP")%>" target="_blank">点击查看位置</a>]　发贴时间：<%=rs("AddTime")%></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">内容</td>
  <td colspan=4 class=b1_1><textarea name="content" style="width:600px;height:250px;visibility:hidden;"><%=rs("Content")%></textarea></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">回复</td>
  <td colspan=4 class=b1_1><textarea name="recontent" cols="80" rows="10" id="recontent"><%=rs("recontent")%></textarea></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td colspan=4 class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="提 交">
  <input name="yn" type="checkbox" id="yn" value="1" style="border:0" <%if rs("yn")=1 then Response.Write("checked") end if%>>
  审核</td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	Dim Article_title
	id=request.form("id")
	title				=trim(request.form("title"))
	username			=trim(request.form("username"))
	qq					=trim(request.form("qq"))
	email				=request.form("email")
	Content				=request.form("Content")
	ReContent			=request.form("ReContent")
	yn					=request.form("yn")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_GuestBook where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("title")				=title
		rs("UserName")			=UserName
		rs("Content")			=Content
		if recontent<>"" then
		rs("ReContent")			=ReContent
		end if
		rs("ReTime")			=Now()
		if yn=1 then
		rs("yn")=1
		else
		rs("yn")=0
		end if
		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_Guestbook.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_Guestbook.asp"";</script>"
	end if
	rs.close
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('请选择!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="批量未审" Then
   set rs=conn.execute("update xiaowei_GuestBook set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_Guestbook.asp"";</script>")
ElseIf Request("Del")="批量审核" Then
   set rs=conn.execute("update xiaowei_GuestBook set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_Guestbook.asp"";</script>")
ElseIf Request("Del")="删除" Then
	set rs=conn.execute("delete from xiaowei_GuestBook where ID In(" & ID & ")")
   	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Guestbook.asp"";</script>"
End If
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>