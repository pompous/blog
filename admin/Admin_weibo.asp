
<!--#include file="../xwInc/conn.asp"-->
<!--#include file="admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
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
					allowFileManager : true ,
 //�����ԣ��������д�����п��ޣ���Ӱ���ȡtextarea��ֵ
 //afterCreate: function(){this.sync();}
 //�������д�����ǹؼ������ڣ���ʧȥ����ʱִ�� this.sync();
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
		<td width="1" bgcolor="#006699">��</td>
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
	sql="select * from xiaowei_2weima order by id desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Response.Write("û��΢��!<a href='?action=add'>[���]</a>")
	else
%>

		
<form name="myform" method="POST" action="admin_weibo.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="4" align=left class="admintitle">΢���б� <a href="?action=add">[���]</a></td>
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
    <td width="46%" bgcolor="f7f7f7"><%=NoI%>.��<%=rs("UserName")%>��</td>
    <td width="35%" height="25" align="center" bgcolor="f7f7f7"><span class="td"><%=rs("AddTime")%></span></td>
    <td width="15%" align="center" bgcolor="f7f7f7"><%if rs("yn")=0 then Response.Write("<font color=red>δ��</font>") else Response.Write("����") end if%>|<a href="?action=edit&id=<%=rs("ID")%>">�༭</a>|<a href="?action=del&id=<%=rs("ID")%>" onClick="JavaScript:return confirm('ȷ��ɾ����')">ɾ��</a></td>
    </tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr><td align=center><input name="Action" type="hidden"  value="Del">
    <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="3">
  	<input type="submit" value="ɾ��" name="Del" id="Del">
    <input type="submit" value="����δ��" name="Del" id="Del">
	<input type="submit" value="�������" name="Del" id="Del">
    ��ҳ��
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
	set rs=conn.execute("delete from xiaowei_2weima where id="&id)
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""admin_weibo.asp"";</script>"
end sub

 sub add()
  %>

<script language=javascript>
function chk()
{
	if(document.form.UserName.value == "" || document.form.UserName.value.length > 200)
	{
	alert("�����ύ���룬΢������Ϊ�ջ����20�ַ���");
	document.form.UserName.focus();
	document.form.UserName.select();
	return false;
	}
	if(document.form.title.value == "" || document.form.title.value.length > 50)
	{
	alert("�����ύ���룬΢����ԴΪ�ջ����50���ַ���");
	document.form.title.focus();
	document.form.title.select();
	return false;
	}
	if(document.form.content.value == "")
	{
	alert("����д΢�����飡");
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
  <td colspan="2" align=left class="admintitle">���΢��</td>
</tr>

    <tr>
      <td height="45" class="b1_1" align="right" width="20%">΢�����⣺</td>
      <td height="45" class="b1_1"  align="left">
      <input name="UserName" type="text" id="UserName" maxlength="200"  value="<%=NOW()%>" SIZE="40"> <font color=#ff0000>*</font></td>
    </tr>
    <tr>
      <td height="45" class="b1_1" align="right">΢����Դ��</td>
      <td height="45" class="b1_1"  align="left">
      <input name="title" type="text" id="title" maxlength="50"  value="����" SIZE="40"> <font color=#ff0000>*</font> </td>
    </tr>
    
    <tr>
      <td height="45" class="b1_1" align="right">�����ˣ�&nbsp; </td>
      <td height="45" class="b1_1"  align="left"><input name="Images" type="text" id="Images" size="40" maxlength="200" value="Admin" SIZE="40" > 
		�� </td>
    </tr>
    
    <tr>
      <td height="150" class="b1_1" align="right" valign="top">΢�����ݣ�</td>
      <td height="150" class="b1_1"  align="left"  >
		<textarea name="content" style="width:700px;height:350px;visibility:hidden;"></textarea><font color=#ff0000>*</font></td>
    </tr>
          </tr>
    <tr>
      <td height="36" align="center" class="b1_1"></td> 
      <td height="36" class="b1_1" align="left"  ><input type="submit" name="Submit2" class="borderall1" value="  ��  ��   " style="height:25px;color:#fff;font-family:΢���ź�"> <Input name="Cancel" type=reset  class="borderall1" id=Reset value="   ��  ��  " style="height:25px;color:#fff;font-family:΢���ź�"></td>
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

		



	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_2weima"
	rs.open sql,conn,1,3
		if UserName="" or Title="" or Content="" then
			response.Redirect "admin_weibo.asp?action=add"
		end if

		rs.AddNew 
		rs("UserName")			=UserName
		rs("Title")				=Title
		rs("Content")			=Content
		rs("images")			=images
		rs("AddIP")				=Request.ServerVariables("REMOTE_ADDR")

		rs("yn")				=1
		

	
		rs.update
	
		Response.Write("<script language=javascript>alert('��ϲ��,�ύ�ɹ�!');location.href='admin_weibo.asp';</script>")
		
		rs.close
		Set rs = nothing
   %>









<%
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_2weima where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="2" class="admintitle">�ظ�΢��</td>
</tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">΢������</td>
  <td class=b1_1><input name="username" type="text" id="title2" value="<%=rs("username")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">΢����Դ</td>
  <td class=b1_1><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">������</td>
  <td class=b1_1><%=rs("images")%>��</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">IP</td>
  <td class=b1_1><%=rs("AddIP")%>������ʱ�䣺<%=rs("AddTime")%></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">΢������</td>
  <td class=b1_1><textarea name="content" style="width:700px;height:350px;visibility:hidden;"><%=server.htmlencode(rs("content"))%></textarea></td>
</tr>

          </tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="�� ��">
  <input name="yn" type="checkbox" id="yn" value="1" style="border:0" <%if rs("yn")=1 then Response.Write("checked") end if%>>
  ���</td>
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


	Content				=request.form("Content")

	yn					=request.form("yn")

	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_2weima where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("title")				=title
		rs("UserName")			=UserName
		rs("Content")			=Content
		if recontent<>"" then

		end if
		rs("ReTime")			=Now()
		if yn=1 then
		rs("yn")=1
		else
		rs("yn")=0
		end if
	

	

		rs.update
		Response.write"<script>alert(""��ϲ,�޸ĳɹ���"");location.href=""admin_weibo.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""admin_weibo.asp"";</script>"
	end if
	rs.close
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('��ѡ��!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="����δ��" Then
   set rs=conn.execute("update xiaowei_2weima set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""admin_weibo.asp"";</script>")
ElseIf Request("Del")="�������" Then
   set rs=conn.execute("update xiaowei_2weima set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""admin_weibo.asp"";</script>")
ElseIf Request("Del")="ɾ��" Then
	set rs=conn.execute("delete from xiaowei_2weima where ID In(" & ID & ")")
   	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""admin_weibo.asp"";</script>"
End If
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>