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
				<link rel="stylesheet" href="../kindeditor/themes/default/default.css" />
		<script src="../kindeditor/kindeditor-min.js"></script>
		<script src="../kindeditor/lang/zh_CN.js"></script>
		<script>
			KindEditor.ready(function(K) {
				var editor = K.editor({
					allowFileManager : true
				});
				K('#image1').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : K('#urlimg').val(),
							clickFn : function(url, title, width, height, border, align) {
								K('#urlimg').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
			});
		</script>
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
   Dim Sqlp,Rsp,TempStr
%>
<table border="0" cellspacing="2" cellpadding="3"  align="center" class="admintable">
<tr> 
  <td colspan="6" align=left class="admintitle">栏目列表　[<a href="?action=add">添加</a>]</td>
</tr>
  <tr align="center"> 
    <td width="30%" class="ButtonList">栏目名称</td>
    <td width="10%" class="ButtonList">用户投稿</td>
    <td width="9%" class="ButtonList">排序</td>
    <td width="9%" class="ButtonList">菜单显示</td>
    <td width="9%" class="ButtonList">首页显示</td>
    <td width="37%" class="ButtonList">操 作</td>
  </tr>
<%
   Sqlp ="select * from xiaowei_Class Where TopID = 0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("没有分类")
   Else
   NoI=0
      Do while not Rsp.Eof   
	NoI=NoI+1
%>
    <tr bgcolor="#f1f3f5" onMouseOver="this.style.backgroundColor='#EAFCD5';this.style.color='red'" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" class="tdleft"><%=NoI%> . <%=rsp("ClassName")%> <%If rsp("url")<>"" then Response.Write("<font color=blue>[外]</font>") else Response.Write("<font color=red>("&Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rsp("ID")&"",1)(0)&")</font>") end if%></td>
    <td height="25" align="center" class="tdleft"><%If rsp("IsUser")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td height="25" align="center"><%=rsp("Num")%></td>
    <td height="25" align="center"><%If rsp("IsMenu")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td height="25" align="center"><%If rsp("IsIndex")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td width="37%" align="center">  <a href="?action=edit&id=<%=rsp("ID")%>">编辑</a> | <a href="?action=del&id=<%=rsp("ID")%>" onClick="JavaScript:return confirm('删除栏目同时会删除该栏目下的文章！确定？')">删除</a></td>
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
    <td height="25" class="tdleft">　　|- <%=rspp("ClassName")%> <font color=red>(<%=Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rspp("ID")&"",1)(0)%>)</font></td>
    <td height="25" align="center" class="tdleft"><%If rspp("IsUser")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td height="25" align="center"><%=rspp("Num")%></td>
    <td height="25" align="center"><%If rspp("IsMenu")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td height="25" align="center"><%If rspp("IsIndex")=1 then Response.Write("<font color=red>√</font>") else Response.Write("ㄨ") end if%></td>
    <td width="37%" align="center"> <a href="?action=edit&id=<%=rspp("ID")%>">编辑</a> | <a href="?action=del&id=<%=rspp("ID")%>" onClick="JavaScript:return confirm('删除栏目同时会删除该栏目下的文章！确定？')">删除</a></td>
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
  <td colspan="2" class="admintitle">添加栏目</th></tr>
<form action="?action=savenew" onSubmit="return chk();" method="post" name="myForm">
<tr>
<td width="20%" class=b1_1>栏目名称</td>
<td class=b1_1><input type="text" name="ClassName" size="30"></td>
</tr>
<tr>
  <td class=b1_1>显示方式</td>
  <td class=b1_1><select name="sort" id="sort">
      <option value="1" selected>图文混排</option>
      <option value="2">图片方式</option>
     

    </select></td>
</tr>
<tr>
  <td class=b1_1>所属栏目</td>
  <td class=b1_1><select ID="TopID" name="TopID">
    <%call Admin_ShowClass_Option()%></select></td>
</tr>
<tr> 
<td width="20%" class=b1_1>排　　序</td>
<td class=b1_1><input name="num" type="text" value="10" size="4" maxlength="3"></td>
</tr>
<tr> 
<td width="20%" class=b1_1>栏目外连地址</td>
<td class=b1_1><input name="Url" type="text" id="Url" size="42">
<span class="note">不外连请留空！</span></td>
</tr>
<tr>
  <td class=b1_1>广告图片地址</td>
  <td class=b1_1>    <p><input name="urlimg"  type="text" id="urlimg" value="" size="55"/> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）<span class="note">栏目广告图片大小要求：250×100&nbsp;&nbsp;&nbsp;</span></p>    </td>
</tr>
<tr>
  <td class=b1_1>广告图片链接地址</td>
  <td class=b1_1><input name="Urllink" type="text" id="Urllink" size="55"> <span class="note">以 
	http:// 开头</span></td>
</tr>
<tr>
  <td class=b1_1>栏目介绍</td>
  <td class=b1_1><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"></textarea></td>
</tr>
<tr>
  <td class=b1_1>打开方式</td>
  <td class=b1_1><select name="target" id="target">
      <option value="_top" selected>_top</option>
      <option value="_blank">_blank</option>
      <option value="_parent">_parent</option>
      <option value="_self">_self</option>
    </select></td>
</tr>
<tr>
  <td class=b1_1>导航栏是否显示</td>
  <td class=b1_1><input name="IsMenu" type="radio" class="noborder" value="1" checked>
    是
      <input name="IsMenu" type="radio" class="noborder" value="0">
      否&nbsp;&nbsp;&nbsp;&nbsp; <span class="note">二级目录建议选择 否</span> </td>
</tr>
<tr>
  <td class=b1_1>首页是否显示</td>
  <td class=b1_1><input name="IsIndex" type="radio" class="noborder" value="1" checked>
是
  <input name="IsIndex" type="radio" class="noborder" value="0">
否&nbsp;&nbsp;&nbsp;&nbsp; <span class="note">二级目录无需设置</span></td>
</tr>
<tr>
  <td class=b1_1>首页显示数量</td>
  <td class=b1_1><input name="IndexNum" type="text" id="IndexNum" value="11" size="4" maxlength="2">&nbsp; 
	默认为 11&nbsp; 禁止修改</td>
</tr>
<tr>
  <td class=b1_1>该栏目是否允许用户投稿</td>
  <td class=b1_1><input name="IsUser" type="radio" class="noborder" value="1">
是
  <input name="IsUser" type="radio" class="noborder" value="0" checked>
否</td>
</tr>
<tr> 
<td width="20%" class=b1_1></td>
<td class=b1_1><input type="submit" name="Submit" value="添 加"></td>
</tr></form>
</table>
<%
end sub

sub del()
	id=request("id")
	If Mydb("Select Count([ID]) From [xiaowei_Class] Where TopID="&ID&"",1)(0)>0 then
		Response.Write("<script language=javascript>alert('请先删除下级栏目!');history.back(1);</script>")
		Response.End
	else
		set rs=conn.execute("delete from xiaowei_Class where id="&id)
		set rs=conn.execute("delete from xiaowei_Article where ClassID In(" & ID & ")")
	end if
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Class.asp"";</script>"
end sub

sub savenew()
	if trim(request.form("ClassName"))="" then
		Response.write"<script>alert(""请填写栏目名称！"");location.href=""?action=add"";</script>"
		Response.End
	end if
	ClassName=trim(request.form("ClassName"))
	num=trim(request.form("num"))
	ReadMe=trim(request.form("ReadMe"))
	IsMenu=request.form("IsMenu")
	IsIndex=request.form("IsIndex")
	Indexnum=trim(request.form("Indexnum"))
	TopID=request.form("TopID")
	Url=trim(request.form("Url"))
	Urlimg=trim(request.form("Urlimg"))
	Urlilink=trim(request.form("Urllink"))
	sort=trim(request.form("sort"))	
	target=trim(request.form("target"))
	IsUser=request.form("IsUser")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Class where ClassName='"& ClassName &"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("ClassName")		=ClassName
		rs("num")			=num
		rs("ReadMe")		=ReadMe
		rs("IsMenu")		=IsMenu
		rs("IsIndex")		=IsIndex
		rs("Indexnum")		=Indexnum
		rs("TopID")			=TopID
		rs("url")			=Url
		rs("urlimg")		=Urlimg
		rs("urllink")		=Urllink
		rs("sort")	    	=sort
		rs("target")		=target
		If rs("url")<>"" then
		rs("link")			=1
		else
		rs("link")			=0
		End if
		rs("IsUser")		=IsUser
		
		rs.update
		Response.write"<script>alert(""恭喜,添加成功！"");location.href=""Admin_Class.asp"";</script>"
	else
		Response.write"<script>alert(""添加失败，该栏目已经存在！"");location.href=""Admin_Class.asp"";</script>"
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
<form action="?action=savedit" onSubmit="return chk();" method="post" name="myform">
<tr> 
    <td colspan="2" class="admintitle">修改栏目</td>
</tr>
<tr> 
<td width="20%" class="b1_1">栏目名称</td>
<td class=b1_1><input type="text" name="ClassName" size="30" value="<%=rs("ClassName")%>"></td>
</tr>
<tr>
  <td class=b1_1>打开方式</td>
  <td class=b1_1><select name="sort" id="sort">
    <option value="1"<%If rs("sort")="1" then Response.Write(" selected") end if%>>图文混排</option>
    <option value="2"<%If rs("sort")="2" then Response.Write(" selected") end if%>>图片方式</option>
  

  </select></td>
</tr>
<tr>
  <td class="b1_1">所属栏目</td>
  <td class=b1_1><select ID="TopID" name="TopID">
    <%   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = 0 and ID<>"&ID&" And Link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value=""0"">做为顶级分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
		 If rs("topid")=Rsp("ID") then
			Response.Write(" selected" ) 
		 end if
         Response.Write(">|-" & Rsp("ClassName") & "")
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if%>
  </select></td>
</tr>
<tr>
  <td class="b1_1">排　　序</td>
  <td class=b1_1><input name="Num" type="text" id="Num" value="<%=rs("Num")%>" size="4" maxlength="3"></td>
</tr>
<tr>
  <td class="b1_1">栏目外连地址</td>
  <td class=b1_1>
	<input name="Url" type="text" id="Url" value="<%=rs("Url")%>" size="43">
    <span class="note">不外连请留空！</span></td>
</tr>
<tr>
  <td class=b1_1>广告图片地址</td>
  <td class=b1_1><p><input name="urlimg"  type="text" id="urlimg" value="<%=rs("Urlimg")%>" size="55"/> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）<span class="note">栏目广告图片大小要求：250×100&nbsp;&nbsp;&nbsp;</span></p> </td>
</tr>
<tr>
  <td class="b1_1">广告图片链接地址</td>
  <td class=b1_1><input name="Urllink" type="text" id="Urllink" value="<%=rs("Urllink")%>" size="56">
	<span class="note">以 http:// 开头</span></td>
</tr>
<tr>
  <td class="b1_1">栏目介绍</td>
  <td class=b1_1><textarea name="ReadMe" cols="40" rows="5" id="ReadMe"><%=rs("ReadMe")%></textarea></td>
</tr>
<tr>
  <td class=b1_1>打开方式</td>
  <td class=b1_1><select name="target" id="target">
    <option value="_top"<%If rs("target")="_top" then Response.Write(" selected") end if%>>_top</option>
    <option value="_blank"<%If rs("target")="_blank" then Response.Write(" selected") end if%>>_blank</option>
    <option value="_parent"<%If rs("target")="_parent" then Response.Write(" selected") end if%>>_parent</option>
    <option value="_self"<%If rs("target")="_self" then Response.Write(" selected") end if%>>_self</option>
  </select></td>
</tr>
<tr>
  <td class=b1_1>导航栏是否显示</td>
  <td class=b1_1><input name="IsMenu" type="radio" class="noborder" value="1"<%If rs("IsMenu")=1 then Response.Write(" checked") end if%>>
是
  <input name="IsMenu" type="radio" class="noborder" value="0"<%If rs("IsMenu")=0 then Response.Write(" checked") end if%>>
否&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  <span class="note">二级目录建议选择 否</span> </td>
</tr>
<tr>
  <td class=b1_1>首页是否显示</td>
  <td class=b1_1><input name="IsIndex" type="radio" class="noborder" value="1"<%If rs("IsIndex")=1 then Response.Write(" checked") end if%>>
是
  <input name="IsIndex" type="radio" class="noborder" value="0"<%If rs("IsIndex")=0 then Response.Write(" checked") end if%>>
否&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <span class="note">二级目录无需设置</span></td>
</tr>
<tr>
  <td class=b1_1>首页显示数量</td>
  <td class=b1_1><input name="IndexNum" type="text" id="IndexNum" value="<%=rs("IndexNum")%>" size="4" maxlength="2"> 
	默认为 11&nbsp; 禁止修改</td>
</tr>
<tr>
  <td class=b1_1>是否允许用户投稿</td>
  <td class=b1_1><input name="IsUser" type="radio" class="noborder" value="1"<%If rs("IsUser")=1 then Response.Write(" checked") end if%>>
是
  <input name="IsUser" type="radio" class="noborder" value="0"<%If rs("IsUser")=0 then Response.Write(" checked") end if%>>
否</td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input type="submit" name="Submit" value="提 交"></td>
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
	IsIndex=request.form("IsIndex")
	Indexnum=trim(request.form("Indexnum"))
	Url=trim(request.form("Url"))
	Urlimg=trim(request.form("Urlimg"))
	Urllink=trim(request.form("Urllink"))
	target=trim(request.form("target"))
	sort=trim(request.form("sort"))
	IsUser=request.form("IsUser")
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Class where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("ClassName")		= ClassName
		rs("Num")			= Num
		rs("ReadMe")		=ReadMe
		rs("IsMenu")		=IsMenu
		rs("IsIndex")		=IsIndex
		rs("Indexnum")		=Indexnum
		rs("TopID")			=TopID
		rs("url")			=Url
		rs("urlimg")		=Urlimg
		rs("urllink")		=Urllink
		rs("sort")		=sort
		rs("target")		=target
		If rs("url")<>"" then
		rs("link")			=1
		else
		rs("link")			=0
		End if
		rs("IsUser")		=IsUser
		
		rs.update
		Response.write"<script>alert(""恭喜,修改成功！"");location.href=""Admin_Class.asp"";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_Class.asp"";</script>"
	end if
	rs.close
end sub

sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = 0 And Link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value=""0"">做为顶级分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
         Response.Write(">|-" & Rsp("ClassName") & "")
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
end sub 
%>


				
		</td>
	</tr>
</table>


</body>

</html>
