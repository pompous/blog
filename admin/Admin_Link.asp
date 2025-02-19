<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>vpCMS 系统管理</title>
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
	sql="select * from xiaowei_link order by id desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Response.Write("没有链接!<a href='?action=add'>[添加]</a>")
	else
%>
<form name="myform" method="POST" action="Admin_link.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="4" align=left class="admintitle">链接列表 <a href="?action=add">[添加]</a></td>
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
    <td width="4%" height="25" align="center" bgcolor="f7f7f7"><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td width="46%" bgcolor="f7f7f7"><%=NoI%>.　<%=rs("UserName")%>　<%=rs("title")%></td>
    <td width="35%" height="25" align="center" bgcolor="f7f7f7"><span class="td"><%=rs("AddTime")%></span></td>
    <td width="15%" align="center" bgcolor="f7f7f7"><%if rs("yn")=0 then Response.Write("<font color=red>未审</font>") else Response.Write("已审") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">编辑</a>|<a href="?action=del&id=<%=rs("ID")%>" onClick="JavaScript:return confirm('确认删除吗？')">删除</a></td>
    </tr>
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
	set rs=conn.execute("delete from xiaowei_link where id="&id)
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_link.asp"";</script>"
end sub

 sub add()
  %>

<script language=javascript>
function chk()
{
	if(document.form.UserName.value == "" || document.form.UserName.value.length > 20)
	{
	alert("不能提交申请，网站名称为空或大于20字符！");
	document.form.UserName.focus();
	document.form.UserName.select();
	return false;
	}
	if(document.form.title.value == "" || document.form.title.value.length > 50)
	{
	alert("不能提交申请，网站地址为空或大于50个字符！");
	document.form.title.focus();
	document.form.title.select();
	return false;
	}
	if(document.form.content.value == "")
	{
	alert("请填写网站简介！");
	document.form.content.focus();
	document.form.content.select();
	return false;
	}
	return true;
}
</script>

<form onSubmit="return chk();" method="post" name="form" action="?action=savenew">
 
 <table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="2" align=left class="admintitle">添加链接 </td>
</tr>

    <tr>
      <td height="45" class="b1_1" align="right" width="20%">网站名称：</td>
      <td height="45" class="b1_1"  align="left">
      <input name="UserName" type="text" id="UserName" maxlength="20"  value="" SIZE="40"> <font color=#ff0000>*</font></td>
    </tr>
    <tr>
      <td height="45" class="b1_1" align="right">网址URL：</td>
      <td height="45" class="b1_1"  align="left">
      <input name="title" type="text" id="title" maxlength="50"  value="http://" SIZE="40"> <font color=#ff0000>*</font> </td>
    </tr>
    
    <tr>
      <td height="45" class="b1_1" align="right">LOGO图：</td>
      <td height="45" class="b1_1"  align="left"><input name="Images" type="text" id="Images" size="40" maxlength="200" value="" SIZE="40" >
      　 没有LOGO请留空 大小： 88*31</td>
    </tr>
    
    <tr>
      <td height="150" class="b1_1" align="right" valign="top">网站简介：</td>
      <td height="150" class="b1_1"  align="left"  >
		<textarea name="content" cols="60" rows="6" id="content" ></textarea> <font color=#ff0000>*</font></td>
    </tr>
              <tr>
            <td height="25" bgcolor="f7f7f7" class="tdleft">链接是否首页显示：</td>
  <td class=b1_1><input name="linkynoff" type="radio" class="noborder" value="1" checked>
    是
      <input name="linkynoff" type="radio" class="noborder" value="0">
      否</td>
</tr>
          </tr>
    <tr>
      <td height="36" align="center" class="b1_1"></td> 
      <td height="36" class="b1_1" align="left"  ><input type="submit" name="Submit2" class="borderall1" value="  提  交   " style="height:25px;color:#fff;font-family:微软雅黑"> <Input name="Cancel" type=reset  class="borderall1" id=Reset value="   重  填  " style="height:25px;color:#fff;font-family:微软雅黑"></td>
    </tr>
  </table>
<br>
	
		
	 </form>
<% end sub 
	sub savenew()
	dim UserName,Title,Content,images
	UserName = trim(request.form("UserName"))
	images = trim(request.form("images"))
	Title = trim(request.form("Title"))
	Content = request.form("Content")
	linkynOFF = request.form("LINKynoff")
		

  	if Instr(Content,"傻B")>0 or Instr(Content,"操你妈")>0 or Instr(Content,"胡锦")>0 or Instr(Content,"江泽")>0 or Instr(Content,"娘")>0 or Instr(Content,"[url")>0 or HasChinese(Content)=false then
	Response.Write "<script language=javascript>alert('请不要乱发信息,谢谢!');javascript:history.back();</script>"
	Response.End
	end if

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_link"
	rs.open sql,conn,1,3
		if UserName="" or Title="" or Content="" then
			response.Redirect "admin_link.asp?action=add"
		end if

		rs.AddNew 
		rs("UserName")			=UserName
		rs("Title")				=Title
		rs("Content")			=Content
		rs("images")			=images
		rs("AddIP")				=Request.ServerVariables("REMOTE_ADDR")

		rs("yn")				=1
		if linkynoff=1 then
		rs("linkynoff")         =1
		else
		rs("linkynoff")         =0
		end if
		rs.update
	
		Response.Write("<script language=javascript>alert('恭喜你,提交成功!');location.href='admin_link.asp';</script>")
		
		rs.close
		Set rs = nothing
   %>









<%
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_link where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="5" class="admintitle">回复链接</td>
</tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">网站名称</td>
  <td colspan=4 class=b1_1><input name="username" type="text" id="title2" value="<%=rs("username")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">网站地址</td>
  <td colspan=4 class=b1_1><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">LOGO图片</td>
  <td colspan=4 class=b1_1><img src="../xwupload/link/<%=rs("images")%>">　</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">IP</td>
  <td colspan=4 class=b1_1><%=rs("AddIP")%>　申请时间：<%=rs("AddTime")%></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">网站简介</td>
  <td colspan=4 class=b1_1><textarea name="Content" cols="80" rows="10" id="Content"><%=rs("Content")%></textarea></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">回复</td>
  <td colspan=4 class=b1_1><textarea name="recontent" cols="80" rows="10" id="recontent"><%=rs("recontent")%></textarea></td>
</tr>
          <tr>
            <td height="25" bgcolor="f7f7f7" class="tdleft">链接是否首页显示：</td>
<td class=b1_1><input name="linkynoff" type="radio" class="noborder" value="1"<%If rs("linkynoff")=1 then Response.Write(" checked") end if%>>
是
  <input name="linkynoff" type="radio" class="noborder" value="0"<%If rs("linkynoff")=0 then Response.Write(" checked") end if%>>
否</td>
</tr>

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
	linkynOFF			=request.form("LINKynoff")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_link where ID="&id&""
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
		if linkynoff=1 then
		rs("linkynoff")=1
		else
		rs("linkynoff")=0
		end if

		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_link.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_link.asp"";</script>"
	end if
	rs.close
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('请选择!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="批量未审" Then
   set rs=conn.execute("update xiaowei_link set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_link.asp"";</script>")
ElseIf Request("Del")="批量审核" Then
   set rs=conn.execute("update xiaowei_link set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_link.asp"";</script>")
ElseIf Request("Del")="删除" Then
	set rs=conn.execute("delete from xiaowei_link where ID In(" & ID & ")")
   	Response.write"<script>alert(""删除成功！"");location.href=""Admin_link.asp"";</script>"
End If
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>