<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS ϵͳ����</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
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

	if request("action")="edit" then
		call edit()

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
	sql="select * from xiaowei_chz order by id desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Response.Write("û�г�ֵ��¼!")
	else
%>
<form name="myform" method="POST" action="admin_chz.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="4" align=left class="admintitle">��ֵ��¼�б� </td>
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
    <td width="46%" bgcolor="f7f7f7"><%=NoI%>.��<%=rs("UserName")%>��<%=rs("title")%></td>
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
	set rs=conn.execute("delete from xiaowei_chz where id="&id)
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""admin_chz.asp"";</script>"
end sub





sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_chz where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="2" class="admintitle">�༭��ֵ��¼</td>
</tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">�û���</td>
  <td class=b1_1><input name="username" type="text" id="username" value="<%=rs("username")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">��ˮ��</td>
  <td class=b1_1><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">���</td>
  <td class=b1_1>
	<input name="images" type="text" id="images" value="<%=rs("images")%>" size="30"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">IP</td>
  <td class=b1_1><%=rs("AddIP")%>������ʱ�䣺<%=rs("AddTime")%></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">�ظ�</td>
  <td class=b1_1><textarea name="recontent" cols="60" rows="5" id="recontent"><%=rs("recontent")%></textarea></td>
</tr>
          <tr>
            <td height="25" bgcolor="f7f7f7" class="tdleft">��ȷ�ϣ�</td>
<td class=b1_1><input name="linkynoff" type="radio" class="noborder" value="1"<%If rs("linkynoff")=1 then Response.Write(" checked") end if%>>
��
  <input name="linkynoff" type="radio" class="noborder" value="0"<%If rs("linkynoff")=0 then Response.Write(" checked") end if%>>
��</td>
</tr>

          </tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="�� ��">
  <input name="yn" type="checkbox" id="yn" value="1" style="border:0" <%if rs("yn")=1 then Response.Write("checked") end if%>> 
��ת���û��ʻ�</td>
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
	images					=trim(request.form("images"))

	ReContent			=request.form("ReContent")
	yn					=request.form("yn")
	linkynOFF			=request.form("LINKynoff")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_chz where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("title")				=title
		rs("UserName")			=UserName
	rs("images")               =images
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
		Response.write"<script>alert(""��ϲ,�޸ĳɹ���"");location.href=""admin_chz.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""admin_chz.asp"";</script>"
	end if
	rs.close
end sub

Sub delAll
ID=Trim(Request("ID"))
If ID="" Then
	  Response.Write("<script language=javascript>alert('��ѡ��!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="����δ��" Then
   set rs=conn.execute("update xiaowei_chz set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""admin_chz.asp"";</script>")
ElseIf Request("Del")="�������" Then
   set rs=conn.execute("update xiaowei_chz set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""�����ɹ���"");location.href=""admin_chz.asp"";</script>")
ElseIf Request("Del")="ɾ��" Then
	set rs=conn.execute("delete from xiaowei_chz where ID In(" & ID & ")")
   	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""admin_chz.asp"";</script>"
End If
End Sub
%>

				
		</td>
	</tr>
</table>


</body>

</html>