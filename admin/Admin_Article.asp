<!--#include file="../xwInc/conn.asp"-->
<!--#include file="admin_check.asp"-->
<!--#include file="../xwInc/saveimage.asp"-->
<!--#include file="../xwInc/Function_Page.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理</title>
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
 //经测试，下面这行代码可有可无，不影响获取textarea的值
 //afterCreate: function(){this.sync();}
 //下面这行代码就是关键的所在，当失去焦点时执行 this.sync();
 afterBlur: function(){this.sync();}
				});

							});
		</script>
		
		<script>
			KindEditor.ready(function(K) {
				var editor = K.editor({
					allowFileManager : true
				});
				K('#image1').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : K('#images').val(),
							clickFn : function(url, title, width, height, border, align) {
								K('#images').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
			});
		</script>

</head>
<script src="js/admin.js"></script>
<body topmargin="0" leftmargin="0">



<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"  valign="top"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">　</td>
		<td valign="top"><br>	
			
<script language=javascript>
function CheckForm()
{ 
  if (document.myform.Title.value==""){
	alert("请填写标题！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.ClassID.value==""){
	alert("请选择分类！");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.Hits.value==""){
	alert("请填写浏览次数！");
	document.myform.Hits.focus();
	return false;
  }
  var filter=/^\s*[0-9]{1,6}\s*$/;
  if (!filter.test(document.myform.Hits.value)) { 
	alert("浏览次数填写不正确,请重新填写！"); 
	document.myform.Hits.focus();
	return false; 
  }
}
</script>

<table width="95%" border="0" cellspacing="2" cellpadding="3"  align=center class="admintable" style="margin-bottom:5px;">
    <tr><form name="form1" method="get" action="Admin_Article.asp">
      <td height="25" bgcolor="f7f7f7">快速查找：
        <SELECT onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')"  size="1" name="s">
        <OPTION value="" selected>-=请选择=-</OPTION>
        <OPTION value="?s=all">所有文章</OPTION>
        <OPTION value="?s=yn0">已审的文章</OPTION>
        <OPTION value="?s=yn1">未审的文章</OPTION>
        <OPTION value="?s=yn2">会员私有文章</OPTION>
          <OPTION value="?s=istop">固顶文章</OPTION>
        <OPTION value="?s=ishot">推荐文章</OPTION>
        <OPTION value="?s=img">有缩略图文章</OPTION>
        <OPTION value="?s=url">转向链接文章</OPTION>
        <OPTION value="?s=user">会员发表的文章</OPTION>
      </SELECT>      </td>
      <td align="center" bgcolor="f7f7f7">
        <a href="?hits=1"></a>
        <input name="keyword" type="text" id="keyword" value="<%=request("keyword")%>">
        <input type="submit" name="Submit2" value="搜索">
        <input onClick="window.location.href='?hits=1'" type='button' class="sub" name='Submit2' value='按浏览次数排序' />
      </form></td>
      <td align="right" bgcolor="f7f7f7">跳转到：
        <select name="ClassID" id="ClassID" onChange="javascript:window.open(this.options[this.selectedIndex].value,'main')">
    <%
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = 0 and link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """?ClassID=" & Rsp("ID") & """" & "")
		 If int(request("ClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & " ("&Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rsp("ID")&"",1)(0)&")")
		 
		    Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """?ClassID=" & Rspp("ID") & """" & "")
				If int(request("ClassID"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　|-" & Rspp("ClassName") & " ("&Mydb("Select Count([ID]) From [xiaowei_Article] Where ClassID="&rspp("ID")&"",1)(0)&")")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
	%>
        </select></td>
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
	elseif request("action")="yn1" then
		call yn1()
	elseif request("action")="yn2" then
		call yn2()
	elseif request("action")="del" then
		call del()
	elseif request("action")="delAll" then
		call delAll()
	else
		call List()
	end if

sub List()
%>
<form name="myform" method="POST" action="Admin_Article.asp?action=delAll">
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="1" bgcolor="#F2F9E8" class="admintable">
<tr> 
  <td colspan="6" align=left class="admintitle">文章列表　[<a href="?action=add">添加</a>]</td></tr>
    <tr bgcolor="#f1f3f5" style="font-weight:bold;">
    <td width="5%" height="30" align="center" class="ButtonList">　</td>
    <td width="45%" align="center" class="ButtonList">文章名称</td>
    <td width="18%" height="25" align="center" class="ButtonList">发布时间</td>
    <td width="7%" height="25" align="center" class="ButtonList">浏览</td>
    <td width="7%" height="25" align="center" class="ButtonList">顶一下</td>
    <td width="20%" height="25" align="center" class="ButtonList">管理</td>    
    </tr>
<%
page=request("page")
Hits=request("hits")
s=Request("s")
Articleclass=request("ClassID")
keyword=request("keyword")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from xiaowei_Article"
	if s="yn0" then
	mysql=mysql&" Where yn=0"
	elseif s="yn1" then
	mysql=mysql&" Where yn=1"
	elseif s="yn2" then
	mysql=mysql&" Where yn=2"
	elseif s="istop" then
	mysql=mysql&" Where istop=1"
	elseif s="ishot" then
	mysql=mysql&" Where ishot=1"
	elseif s="img" then
	mysql=mysql&" Where images<>''"
	elseif s="url" then
	mysql=mysql&" Where LinkUrl<>''"
	elseif s="user" then
	mysql=mysql&" Where UserName<>''"
	elseif Articleclass<>"" then
	mysql=mysql&" Where ClassID="&Articleclass&""
	elseif keyword<>"" then
	mysql=mysql&" Where Title like '%"&keyword&"%'"
	End if
mysql=mysql&" order by "
If Hits=1 then
mysql=mysql&"Hits desc"
Else
mysql=mysql&"DateAndTime desc"
End if
mypage.getsql=mysql
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%>
    <tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor='#EAF5FD';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
    <td height="25" align="CENTER"><input type="checkbox" value="<%=rs("ID")%>" name="ID" onClick="unselectall(this.form)" style="border:0;"></td>
    <td height="25"><a href="<%=SitePath%>xwArticle/?<%=rs("ID")%>.html" target="_blank"><%=rs("Title")%></a> <%if rs("IsTop")=1 then Response.Write("<font color=red>[顶]</font>") end if%><%if rs("IsHot")=1 then Response.Write("<font color=green>[荐]</font>") end if%><%if rs("Images")<>"" then Response.Write("<font color=blue>[图]</font>") end if%><%If rs("UserName")<>"" then Response.Write(" <font color=blue>["&rs("username")&"]</font>") end if%><%If rs("zffy")<>0 then Response.Write(" <font color=red>[费]</font>") end if%></td>
    <td height="25" align="center"><%=rs("DateAndTime")%></td>
    <td height="25" align="center"><%=rs("Hits")%></td>
    <td height="25" align="center"><%=rs("dig")%></td>
    <td align="center"><%If rs("yn")=0 then Response.Write("已审") end if:If rs("yn")=1 then Response.Write("<font color=red>未审</font>") end if:If rs("yn")=2 then Response.Write("<font color=blue>私有</font>") end if%>|<a href="?action=edit&id=<%=rs("ID")%>&page=<%=page%>">编辑</a></td>    
    </tr>
<%
        rs.movenext
    else
         exit for
    end if
next
%>
<tr><td align="center" bgcolor="f7f7f7"><input name="Action" type="hidden"  value="Del"><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" style="border:0"></td>
  <td colspan="5" bgcolor="f7f7f7"><font color=red>移动到：</font>
    <select id="ytype" name="ytype">
      <%call Admin_ShowClass_Option()%>
    </select>
    &nbsp;
    <input type="submit" value="移动" name="Del" id="Del">
	<input type="submit" value="更新时间" name="Del" id="Del">
    <input type="submit" value="删除" name="Del" id="Del">
    <input type="submit" value="批量未审" name="Del" id="Del">
    <input type="submit" value="批量审核" name="Del" id="Del">
    <input type="submit" value="推荐" name="Del" id="Del">
    <input type="submit" value="解除推荐" name="Del" id="Del">
    <input type="submit" value="固顶" name="Del" id="Del">
    <input type="submit" value="解除固顶" name="Del" id="Del"></td>
  </tr><tr><td bgcolor="f7f7f7" colspan="6">
  

<table border="0" cellspacing="5" cellpadding="2" align="center"><tr><%=mypage.showpage()%></tr></table>


</td></tr></table>
</form>
<%
	rs.close
end sub

sub add()
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" action="?action=savenew" name="myform" method=post>
<tr> 
    <td colspan="2" align=left class="admintitle">添加文章</td></tr>
<tr> 
<td width="20%" class="b1_1">标题</td>
<td class="b1_1"><input name="Title" type="text" id="Title" size="40" maxlength="50">

	  </td>
</tr>
<tr>
  <td class="b1_1">关键字</td>
  <td class="b1_1"><input name="KeyWord" type="text" id="KeyWord" size="40" maxlength="50">&nbsp;<span class="note">多个关键字用&nbsp; |&nbsp; 隔开</span></td>
</tr>
<tr>
  <td class="b1_1">作者</td>
  <td class="b1_1"><span class="td">
    <input name="Author" type="text" id="Author" value="" size="40" maxlength="200">
    </span></td>
</tr>
<tr>
  <td class="b1_1">来源</td>
  <td class="b1_1"><span class="td">
    <input name="CopyFrom" type="text" id="CopyFrom" value="" size="40" maxlength="200">
  </span></td>
</tr>
<tr>
  <td class="b1_1">分类</td>
  <td class="b1_1"><select ID="ClassID" name="ClassID">
    <%call Admin_ShowClass_Option()%></select>&nbsp;<span class="note">如果有二级分类,请选择二级分类.</span></td>
</tr>

<tr>
  <td class="b1_1">浏览次数</td>
  <td class="b1_1"><input name="Hits" type="text" id="Hits" value="1" size="6" maxlength="10"></td>
</tr>
<tr>
  <td class="b1_1">顶一下</td>
  <td class="b1_1"><input name="dig" type="text" id="dig" value="1" size="6" maxlength="10"></td>
</tr>
<tr>
  <td class="b1_1">查阅价格</td>
  <td class="b1_1"><input name="zffy" type="text" id="zffy" value="0" size="6" maxlength="10" <%if useroff="0" then %> Readonly <% end if %>>&nbsp;<span class="note">默认为“0”，无需支付积分。“用户系统”关闭时，无法设置价格！</span></td>
</tr>
<tr>
  <td class="b1_1">附件下载地址</td>
  <td class="b1_1"><input name="LinkUrl" type="text" id="linkurl" size="80">
		</td>
</tr>
<tr>
  <td class="b1_1">缩略图<br><span class="note">文章的缩略图，正文中的图片请通过编辑器的插入图片按钮上传!</span></td>
  <td class="b1_1"><input  name="images" type="text" id="images" value=""  size="55" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）&nbsp;&nbsp;       <input name="sSaveFileSelect" type="checkbox" class="noborder" id="sSaveFileSelect" value="1">显示到内容详情页中</td>
</tr>
<tr>
  <td valign="top" class="b1_1"><p>内容</p>
    <p>发布时间<br>
      <input name="DateAndTime" type="text" id="DateAndTime" value="<%=NOW()%>">
    </p>
    <p>
</p></td>
  <td class="b1_1"><textarea name="content" id="content" style="width:700px;height:350px;visibility:hidden;"></textarea></td>
</tr>
<tr>
  <td class="b1_1">附加选项</td>
  <td class="b1_1">固顶
    <input name="IsTop" type="checkbox" class="noborder" id="IsTop" value="1">
    推荐
    <input name="IsHot" type="checkbox" class="noborder" id="IsHot" value="1"></td>
</tr>
<tr>
  <td class="b1_1">手动分页符</td>
  <td class="b1_1"><font color="#FF0000">[xiaowei_page]</font>&nbsp;
	<font color="#CCCCCC">&nbsp;</font><font color="#999999">注：在内容需要分页的地方加上[xiaowei_page]即可，含“[]”</font></td>
</tr>
<tr>
  <td class="b1_1">自动分页字数</td>
  <td class="b1_1"><input name="PageNum" type="text" id="PageNum" value="0" size="6" maxlength="4">
    <span class="note">　注:如果在内容中加入了手动分页符,请填写0</span></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class="b1_1"><input type="submit" name="Submit" value="添 加"></td>
</tr>
</form>
</table>
<%
end sub

sub savenew()
	Title			=trim(request.form("Title"))
	Hits			=trim(request.form("Hits"))
	dig		    	=trim(request.form("dig"))
	zffy			=trim(request.form("zffy"))
	ClassID			=trim(request.form("ClassID"))
	Content			=request.form("Content")
	Author			=trim(request.form("Author"))
	CopyFrom		=trim(request.form("CopyFrom"))
	KeyWord			=trim(request.form("KeyWord"))
	IsTop			=request.form("IsTop")
	IsHot			=request.form("IsHot")

	Images			=trim(request.form("Images"))
	PageNum			=trim(request.form("PageNum"))
	LinkUrl			=trim(request.form("LinkUrl"))

	DateAndTime		=trim(request.form("DateAndTime"))
	sSaveFileSelect =request.Form("sSaveFileSelect")
	
	if Title="" or ClassID="" then
		Call Alert ("请填写完整再提交","-1")
	end if
		If Content="" then
		Call Alert ("你没有填写内容","-1")
	End if	

	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Article where Title='"&Title&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.AddNew 
		rs("Title")				=Title
		rs("Hits")				=Hits
		rs("dig")				  =dig
		rs("zffy")				  =zffy
		rs("ClassID")			=ClassID
		rs("LinkUrl")			=LinkUrl
		rs("Content")			=lContent
		rs("Author")			=Author
		rs("CopyFrom")			=CopyFrom
		rs("KeyWord")			=KeyWord
		If IsTop=1 then
		rs("IsTop")				=1
		else
		rs("IsTop")				=0
		end if
		
	If sSaveFileSelect=1 then
		rs("sSaveFileSelect")				=1
		else
		rs("sSaveFileSelect")				=0
		end if

		If IsHot=1 then
		rs("IsHot")				=1
		else
		rs("IsHot")				=0
		end if

		rs("Images")			=Images
		rs("yn")				=0
		rs("PageNum")			=PageNum
		rs("DateAndTime")		=DateAndTime
	
      		Rs("Content")=Content
		

		rs.update
		session("xiaoweiClassID")=ClassID
		Response.write"<script>alert(""添加成功！"");location.href=""Admin_Article.asp"";</script>"
	else
		Response.Write("<script language=javascript>alert('该文章已存在!');history.back(1);</script>")
	end if
	rs.close
end sub

sub del()
	id=request("id")
	adduserid=request("userid")
	set rs=conn.execute("delete from xiaowei_Article where id="&id)
	If adduserid<>"" then
		set rs=conn.execute("update xiaowei_User set UserMoney = UserMoney-"&money3&" where ID="&adduserid)
	end if
	Response.write"<script>alert(""删除成功！"");location.href=""Admin_Article.asp"";</script>"
end sub

sub edit()
id=request("id")
page=request("page")
set rs = server.CreateObject ("adodb.recordset")
sql="select * from xiaowei_Article where id="& id &""
rs.open sql,conn,1,1
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
<form onSubmit="return CheckForm();" name="myform" action="?action=savedit" method=post>
<tr> 
    <td colspan="2" class="admintitle">修改文章</td></tr>
<tr>
  <td width="20%" bgcolor="#FFFFFF" class="b1_1">标题</td>
  <td class=b1_1><input name="Title" type="text" value="<%=rs("Title")%>" size="40" maxlength="50">
</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">关键字</td>
  <td class=b1_1><input name="KeyWord" type="text" id="KeyWord" value="<%=rs("KeyWord")%>" size="40">&nbsp;<span class="note">多个关键字用 |&nbsp; 隔开</span></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">作者</td>
  <td class=b1_1><input name="Author" type="text" id="Author" value="<%=rs("Author")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">录入员</td>
  <td class=b1_1><input name="UserName" type="text" id="UserName" value="<%=rs("UserName")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">来源</td>
  <td class=b1_1><input name="CopyFrom" type="text" id="CopyFrom" value="<%=rs("CopyFrom")%>" size="40"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">分类</td>
  <td class=b1_1><select ID="ClassID" name="ClassID">
   <%
   Set Rsp=server.CreateObject("adodb.recordset") 
   Sqlp ="select * from xiaowei_Class Where TopID = 0 and link=0 order by num"   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
         If rs("ClassID")=Rsp("ID") Then
            Response.Write(" selected")
         End If
         Response.Write(">|-" & Rsp("ClassName") & "")
		 
		 Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """" & Rspp("ID") & """" & "")
				If rs("ClassID")=Rspp("ID") Then
            	Response.Write(" selected")
         		End If
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
   %>
  </select>  </td></tr>

<tr>
  <td bgcolor="#FFFFFF" class="b1_1">浏览次数</td>
  <td class=b1_1><input name="Hits" type="text" id="Hits" value="<%=rs("Hits")%>" size="6" maxlength="5"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">顶一下</td>
  <td class=b1_1><input name="dig" type="text" id="dig" value="<%=rs("dig")%>" size="6" maxlength="5"></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">查阅价格</td>
  <td class=b1_1><input name="zffy" type="text" id="zffy" value="<%=rs("zffy")%>" size="6" maxlength="5" <%if useroff="0" then %> Readonly <% end if %>></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">附件下载地址</td>
  <td class=b1_1><input name="LinkUrl" type="text" size="80" id="linkurl" value="<%=rs("Linkurl")%>">
			</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">缩略图<br><span class="note">文章的缩略图，正文中的图片请通过编辑器的插入图片按钮上传!</span></td>
  <td class=b1_1><input  name="images" type="text" id="images" value="<%=rs("Images")%>"  size="55" /> <input type="button" id="image1" value="选择图片" />（网络图片 + 本地上传）&nbsp;&nbsp;       <input name="sSaveFileSelect" type="checkbox" class="noborder" id="sSaveFileSelect" value="1">显示到内容详情页中</td></tr>
<tr>
  <td valign="top" bgcolor="#FFFFFF" class="b1_1"><p>内容</p>
    <p>发布时间<br>
      <input name="DateAndTime" type="text" id="DateAndTime" value="<%=rs("DateAndTime")%>">
</p>
    </td>
  <td class=b1_1><textarea name="content" style="width:700px;height:350px;visibility:hidden;"><%=server.htmlencode(rs("Content"))%></textarea></td>
</tr>
<tr>
  <td bgcolor="#FFFFFF" class="b1_1">附加选项</td>
  <td class=b1_1>固顶
    <input name="IsTop" type="checkbox" class="noborder" id="IsTop" value="1"<%if rs("IsTop")=1 then Response.Write("checked") end if%>>
推荐
<input name="IsHot" type="checkbox" class="noborder" id="IsHot" value="1"<%if rs("IsHot")=1 then Response.Write("checked") end if%>></td>
</tr>

<tr>
  <td class="b1_1">手动分页符</td>
  <td class="b1_1"><font color="#FF0000">[xiaowei_page]</font>&nbsp;
	<font color="#CCCCCC">&nbsp;</font><font color="#999999">注：在内容需要分页的地方加上[xiaowei_page]即可，含“[]”</font></td>
</tr>


<tr>
  <td bgcolor="#FFFFFF" class="b1_1">自动分页字数</td>
  <td class=b1_1><input name="PageNum" type="text" id="PageNum" value="<%=rs("PageNum")%>" size="6" maxlength="4"><span class="note">　注:如果在内容中加入了手动分页符,请填写0</span></td>
</tr>
<tr> 
<td width="20%" class="b1_1"></td>
<td class=b1_1><input name="id" type="hidden" value="<%=rs("ID")%>"><input name="page" type="hidden" value="<%=page%>"><input type="submit" name="Submit" value="提 交"></td>
</tr>
</form>
</table>
<%
end sub

sub savedit()
	Dim Title
	id=request.form("id")
	page=request.form("page")
	Title			=trim(request.form("Title"))
	Hits			=trim(request.form("Hits"))
	dig	    		=trim(request.form("dig"))
	zffy			=trim(request.form("zffy"))
	ClassID			=trim(request.form("ClassID"))
	Content			=request.form("Content")
	Author			=trim(request.form("Author"))
	UserName		=trim(request.form("UserName"))
	CopyFrom		=trim(request.form("CopyFrom"))
	KeyWord			=trim(request.form("KeyWord"))
	IsTop			=request.form("IsTop")
	IsHot			=request.form("IsHot")

	Images			=trim(request.form("Images"))
	PageNum			=trim(request.form("PageNum"))
	LinkUrl			=trim(request.form("LinkUrl"))

	DateAndTime		=trim(request.form("DateAndTime"))
	sSaveFileSelect =request.Form("sSaveFileSelect")
	
	if Title="" or ClassID="" then
		Call Alert ("请填写完整再提交","-1")
	end if
		
	If Content="" then
		Call Alert ("你没有填写内容","-1")
	End if
	
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Article where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
	
		rs("Title")				=Title
		rs("Hits")				=Hits
		rs("dig")				=dig
		rs("zffy")				=zffy
		rs("ClassID")			=ClassID
		rs("Content")			=Content
		rs("LinkUrl")			=LinkUrl
		rs("Author")			=Author
		rs("UserName")			=UserName
		rs("CopyFrom")			=CopyFrom
		rs("KeyWord")			=KeyWord
		If IsTop=1 then
		rs("IsTop")				=1
		else
		rs("IsTop")				=0
		end if
		If IsHot=1 then
		rs("IsHot")				=1
		else
		rs("IsHot")				=0
		end if
If sSaveFileSelect=1 then
		rs("sSaveFileSelect")				=1
		else
		rs("sSaveFileSelect")				=0
		end if


		rs("Images")			=Images
		rs("yn")				=0
		rs("PageNum")			=PageNum
		rs("DateAndTime")		=DateAndTime
	
      		Rs("Content")=Content
	
		
		rs.update
		Response.write"<script>alert(""修改成功！"");location.href=""Admin_Article.asp?page="&page&""";</script>"
	else
		Response.write"<script>alert(""修改错误！"");location.href=""Admin_Article.asp?page="&page&""";</script>"
	end if
	rs.close
end sub

sub yn1()
	id=request("id")
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Article where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("yn")=1
		
		rs.update
		Response.write"恭喜,修改成功！"
	else
		Response.write"错误!"
	end if
	rs.close
end sub

sub yn2()
	id=request("id")
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_Article where ID="&id&""
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		rs("yn")=0
		
		rs.update
		Response.write"恭喜,修改成功！"
	else
		Response.write"错误!"
	end if
	rs.close
end sub


sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = 0 and link=0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
		 If int(session("xiaoweiClassID"))=Rsp("ID") then
				Response.Write(" selected" ) 
		 End if
         Response.Write(">|-" & Rsp("ClassName") & "")
		 
		    Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" and link=0 order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """" & Rspp("ID") & """" & "")
				If int(session("xiaoweiClassID"))=Rspp("ID") then
				Response.Write(" selected" ) 
				End if
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
end sub 

Sub delAll
ID=Trim(Request("ID"))
ytype=Request("ytype")
page=request("page")
If ID="" Then
	  Response.Write("<script language=javascript>alert('请选择文章!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="批量未审" Then
   set rs=conn.execute("update xiaowei_Article set yn = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="更新时间" Then
   set rs=conn.execute("update xiaowei_Article set DateAndTime = Now() where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="批量审核" Then
   set rs=conn.execute("update xiaowei_Article set yn = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="推荐" Then
   set rs=conn.execute("update xiaowei_Article set IsHot = 1 where ID In(" & ID & ")")
   Call Alert ("操作成功!","Admin_Article.asp?page="&page&"")
ElseIf Request("Del")="解除推荐" Then
   set rs=conn.execute("update xiaowei_Article set IsHot = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="固顶" Then
   set rs=conn.execute("update xiaowei_Article set IsTop = 1 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="解除固顶" Then
   set rs=conn.execute("update xiaowei_Article set IsTop = 0 where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");history.back(1);</script>")
ElseIf Request("Del")="移动" Then
		If ytype="" then
			Response.Write("<script language=javascript>alert('请选择类别!');history.back(1);</script>")
			Response.End
		End if
   set rs=conn.execute("update xiaowei_Article set ClassID = "&ytype&" where ID In(" & ID & ")")
   Response.Write("<script>alert(""操作成功！"");location.href=""Admin_Article.asp"";</script>")
ElseIf Request("Del")="删除" Then
	'set rs=conn.execute("delete from xiaowei_Article where ID In(" & ID & ")")
    
			for i=1 to request("ID").count
				if request("ID").count=1 then
				ArticleID=request("ID")
				else
				ArticleID=replace(request("id")(i),"'","")
				end if
				'对用户分值操作
				Call EditUserMn(ArticleID,money3,0)
				'删除文章
				Conn.Execute("Delete from [xiaowei_Article] where ID = "&ArticleID&"")
				'删除文章相关评论
				Conn.Execute("Delete from [xiaowei_Pl] where ArticleID = "&ArticleID&"")
			next
			Call Alert ("删除成功","-1")
            
End If
End Sub
%>



				
		</td>
	</tr>
</table>


</body>

</html>