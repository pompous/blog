<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS ϵͳ����</title>
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
				editor = K.create('textarea[name="readme"]', {
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
	else
		call List()
	end if
 
sub List()
   Dim Sqlp,Rsp,TempStr
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
<tr> 
  <td colspan="6" align=left class="admintitle">��ҳ�б�[<a href="?action=add">���</a>]</td>
</tr>
  <tr align="center"> 
    <td width="30%" class="ButtonList">��ҳ����</td>
    <td width="10%" class="ButtonList">�û�Ͷ��</td>
    <td width="9%" class="ButtonList">����</td>
    <td width="9%" class="ButtonList">�˵���ʾ</td>
    <td width="9%" class="ButtonList">��ҳ��ʾ</td>
    <td width="37%" class="ButtonList">�� ��</td>
  </tr>
<%
   Sqlp ="select * from xiaowei_Class Where TopID = -1 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("û�з���")
   Else
   NoI=0
      Do while not Rsp.Eof   
	NoI=NoI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=NoI%> . <%=rsp("ClassName")%> <%If rsp("url")<>"" then Response.Write("<font color=blue>[��]</font>") else Response.Write("<font color=red>("&Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rsp("ID")&"",1)(0)&")</font>") end if%></td>
    <td height="25" align="center" class="tdleft"><%If rsp("IsUser")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td height="25" align="center"><%=rsp("Num")%></td>
    <td height="25" align="center"><%If rsp("IsMenu")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td height="25" align="center"><%If rsp("IsIndex")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td width="37%" align="center">  <a href="?action=edit&id=<%=rsp("ID")%>">�༭</a> | <a href="?action=del&id=<%=rsp("ID")%>" onClick="JavaScript:return confirm('ȷ��ɾ������')">ɾ��</a></td>
  </tr>
<%
		    Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			NoI1=0
			Do while not Rspp.Eof
			NoI1=NoI1+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft">����|- <%=rspp("ClassName")%> <font color=red>(<%=Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rspp("ID")&"",1)(0)%>)</font></td>
    <td height="25" align="center" class="tdleft"><%If rspp("IsUser")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td height="25" align="center"><%=rspp("Num")%></td>
    <td height="25" align="center"><%If rspp("IsMenu")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td height="25" align="center"><%If rspp("IsIndex")=1 then Response.Write("<font color=red>��</font>") else Response.Write("��") end if%></td>
    <td width="37%" align="center"> <a href="?action=edit&id=<%=rspp("ID")%>">�༭</a> | <a href="?action=del&id=<%=rspp("ID")%>" onClick="JavaScript:return confirm('ȷ��ɾ������')">ɾ��</a></td>
  </tr>
<%
			Rspp.Movenext   
      		Loop
			
      Rsp.Movenext   
      Loop   
   End if
%>  
</table>
<%
end sub

sub add()
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<tr> 
  <td colspan="2" class="admintitle">��ӵ�ҳ</th></tr>
<form action="?action=savenew" method=post>
<tr>
<td width="20%" class=b1_1>��ҳ����</td>
<td class=b1_1><input type="text" name="ClassName" size="30"></td>
</tr>
<tr>
  <td class=b1_1>��ҳ����</td>
  <td class=b1_1><select ID="TopID" name="TopID">
    <%call Admin_ShowClass_Option()%></select></td>
</tr>
<tr> 
<td width="20%" class=b1_1>�š�����</td>
<td class=b1_1><input name="num" type="text" value="10" size="4" maxlength="3"></td>
</tr>
<tr>
  <td class=b1_1>��ҳ����</td>
  <td class=b1_1><textarea name="readme" style="width:700px;height:350px;visibility:hidden;"></textarea></td>
</tr>
<tr>
  <td class=b1_1>�򿪷�ʽ</td>
  <td class=b1_1><select name="target2" id="target2">
      <option value="_top" selected>_top</option>
      <option value="_blank">_blank</option>
      <option value="_parent">_parent</option>
      <option value="_self">_self</option>
    </select></td>
</tr>
<tr>
  <td class=b1_1>�������Ƿ���ʾ</td>
  <td class=b1_1><input name="IsMenu" type="radio" class="noborder" value="1" checked>
    ��
      <input name="IsMenu" type="radio" class="noborder" value="0">
      ��&nbsp;&nbsp;&nbsp;&nbsp;  </td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td class=b1_1><input type="submit" name="Submit" value="�� ��"></td>
</tr></form>
</table>
<%
end sub

sub del()
	id=request("id")
	If Mydb("Select Count([ID]) From [xiaowei_Class] Where TopID="&ID&"",1)(0)>0 then
		Response.Write("<script language=javascript>alert('����ɾ���¼���ҳ!');history.back(1);</script>")
		Response.End
	else
		set rs=conn.execute("delete from xiaowei_Class where id="&id)
		set rs=conn.execute("delete from xiaowei_Article where ClassID In(" & ID & ")")
	end if
	Response.write"<script>alert(""ɾ���ɹ���"");location.href=""Admin_single.asp"";</script>"
end sub

sub savenew()
	if trim(request.form("ClassName"))="" then
		Response.write"<script>alert(""����д��ҳ���ƣ�"");location.href=""?action=add"";</script>"
		Response.End
	end if
	ClassName=trim(request.form("ClassName"))
	num=trim(request.form("num"))
	ReadMe=trim(request.form("ReadMe"))
	IsMenu=request.form("IsMenu")
		TopID=request.form("TopID")
		target=trim(request.form("target"))

	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Class where ClassName='"& ClassName &"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("ClassName")		=ClassName
		rs("num")			=num
		rs("ReadMe")		=ReadMe
		rs("IsMenu")		=IsMenu

		rs("TopID")			=TopID

		rs("target")		=target
		
		rs.update
		Response.write"<script>alert(""��ϲ,��ӳɹ���"");location.href=""Admin_single.asp"";</script>"
	else
		Response.write"<script>alert(""���ʧ�ܣ��õ�ҳ�Ѿ����ڣ�"");location.href=""Admin_single.asp"";</script>"
end if
	rs.close
end sub

sub edit()
id=request("id")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_Class where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form action="?action=savedit" method=post>
<tr> 
    <td colspan="2" class="admintitle">�޸ĵ�ҳ</td>
</tr>
<tr> 
<td width="20%" class="b1_1">��ҳ����</td>
<td class=b1_1><input type="text" name="ClassName" size="30" value="<%=rs("ClassName")%>"></td>
</tr>
<tr>
  <td class="b1_1">��ҳ����</td>
  <td class=b1_1><select ID="TopID" name="TopID">
    <%   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = -1 And Link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value=""-1"">��Ϊ��ҳ</option>") 
%>
  </select></td>
</tr>
<tr>
  <td class="b1_1">�š�����</td>
  <td class=b1_1><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="4" maxlength="3"></td>
</tr>
<tr>
  <td class="b1_1">��ҳ����</td>
  <td class=b1_1><textarea name="readme" style="width:700px;height:350px;visibility:hidden;"><%=server.htmlencode(rs("ReadMe"))%></textarea></td>
</tr>
<tr>
  <td class=b1_1>�򿪷�ʽ</td>
  <td class=b1_1><select name="target" id="target">
    <option value="_top"<%If rs("target")="_top" then Response.Write(" selected") end if%>>_top</option>
    <option value="_blank"<%If rs("target")="_blank" then Response.Write(" selected") end if%>>_blank</option>
    <option value="_parent"<%If rs("target")="_parent" then Response.Write(" selected") end if%>>_parent</option>
    <option value="_self"<%If rs("target")="_self" then Response.Write(" selected") end if%>>_self</option>
  </select></td>
</tr>
<tr>
  <td class=b1_1>�������Ƿ���ʾ</td>
  <td class=b1_1><input name="IsMenu" type="radio" class="noborder" value="1"<%If rs("IsMenu")=1 then Response.Write(" checked") end if%>>
��
  <input name="IsMenu" type="radio" class="noborder" value="0"<%If rs("IsMenu")=0 then Response.Write(" checked") end if%>>
��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	Dim ClassName
	id=request.form("id")
	ClassName=request.form("ClassName")
	Num=request.form("Num")
	TopID=request.form("TopID")
	ReadMe=trim(request.form("ReadMe"))
	IsMenu=request.form("IsMenu")
	target=trim(request.form("target"))

	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Class where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("ClassName")		= ClassName
		rs("Num")			= Num
		rs("ReadMe")		=ReadMe
		rs("IsMenu")		=IsMenu

		rs("TopID")			=TopID

		rs("target")		=target

		
		rs.update
		Response.write"<script>alert(""��ϲ,�޸ĳɹ���"");location.href=""Admin_single.asp"";</script>"
	else
		Response.write"<script>alert(""�޸Ĵ���"");location.href=""Admin_single.asp"";</script>"
	end if
	rs.close
end sub

sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = -1 And Link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 

    Response.Write("<option value=""-1"">��Ϊ��ҳ</option>") 

  end sub 
%>

				
		</td>
	</tr>
</table>


</body>

</html>