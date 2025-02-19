<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim SqlItem,RsItem,FoundErr,ErrMsg

Dim ItemID,ItemName,WebName,WebUrl,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ItemDemo
ItemID=Trim(Request("ItemID"))
If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
Else
   ItemID=Clng(ItemID)
End If
If FoundErr<>True Then
   SqlItem ="select ItemID,ItemName,WebName,WebUrl,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ItemDemo from Item where ItemID=" & ItemID
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If RsItem.Eof And RsItem.Bof  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，没有找到该项目！</li>"
   Else
      ItemName=RsItem("ItemName")
      ItemDemo=RsItem("ItemDemo")
      WebName=RsItem("WebName")
      WebUrl=RsItem("WebUrl")
      ClassID=RsItem("ClassID")
      ClassID=RsItem("ClassID")
      SpecialID=RsItem("SpecialID")
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
   End If
   RsItem.Close
   Set RsItem=Nothing
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>
<%Sub Main%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS 系统管理</title>
<link href="../Images/admin_css.css" rel="stylesheet" type="text/css" />
<%Call SetChannel()%>
</head>

<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">　</td>
		<td valign="top"><br>	


<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1">项目编辑 >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>"><font color=red>基本设置</font></a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="Admin_ItemModify4.asp?ItemID=<%=ItemID%>">正文设置</a> >> <a href="Admin_ItemModify5.asp?ItemID=<%=ItemID%>">采样测试</a> >> <a href="Admin_ItemAttribute.asp?ItemID=<%=ItemID%>">属性设置</a> >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemModify2.asp" name="myform">
    <tr> 
      <td colspan="2" class="admintitle">编辑项目--基本设置</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>项目名称：</strong></td>
      <td width="75%" class="b1_1">
	  <input name="ItemName" type="text" size="27" maxlength="30" value="<%=ItemName%>">&nbsp;&nbsp;<font color=red>*</font>如：新浪网－文章中心 既简单又明了      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong> 所属栏目：</strong></td>
      <td width="75%" class="b1_1"><select ID="ClassID" name="ClassID"><%call Admin_ShowChannel_Option(ClassID)%></select></td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> 网站名称：</strong></td>
      <td width="75%" class="b1_1">
	  <input name="WebName" type="text" size="27" maxlength="30" value="<%=WebName%>">      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> 网站网址：</strong></td>
      <td width="75%" class="b1_1"><input name="WebUrl" type="text" size="49" maxlength="150" value="<%=WebUrl%>">      </td>
    </tr>
   <tr> 
      <td width="20%" class="b1_1"><strong> 网站登录：</strong></td>
      <td class="b1_1">
		<input name="LoginType" type="radio" class="noborder" onClick="Login.style.display='none'" value="0" <%if LoginType=0 Then Response.Write "checked"%>>不需要登录<span lang="en-us">&nbsp;
		</span>
		<input name="LoginType" type="radio" class="noborder" onClick="Login.style.display=''" value="1" <%if LoginType=1 Then Response.Write "checked"%>>设置参数      </td>
    </tr>
   <tr id="Login" style="<%If LoginType=0 Then Response.write "display:none"%>"> 
      <td width="20%" class="b1_1"><strong> 登录参数：</strong></td>
      <td class="b1_1">
        登录地址：<input name="LoginUrl" type="text" size="40" maxlength="150" value="<%=LoginUrl%>"><br>
        提交地址：<input name="LoginPostUrl" type="text" size="40" maxlength="150" value="<%=LoginPostUrl%>"><br>
        用户参数：<input name="LoginUser" type="text" size="30" maxlength="150" value="<%=LoginUser%>"><br>
        密码参数：<input name="LoginPass" type="text" size="30" maxlength="150" value="<%=LoginPass%>"><br> 
		失败信息：<input name="LoginFalse" type="text" size="30" maxlength="150" value="<%=LoginFalse%>"></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>模板备注：</strong></td>
      <td width="75%" class="b1_1"><textarea name="ItemDemo" cols="49" rows="5"><%=ItemDemo%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="Cancel" type="button" id="Cancel" value=" 返&nbsp;&nbsp;回 " onClick="window.location.href='Admin_ItemStart.asp'">
<input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步"></td>
    </tr>
</form>
</table>



		</td>
	</tr>
</table>


</body>

</html>
<%End  Sub%>
<%
Sub SetChannel
Dim Arr_Channel,i_Channel,i_Class,i_Special,tmpDepth,i,ArrShowLine(20)
Dim Rs,Sql,ClassID,ClassName,SpecialID,SpecialName
Set Rs=server.createobject("adodb.recordset")
Sql = "select ID from xiaowei_Class order by num ASC"
Rs.Open Sql,Conn,1,1
If Not Rs.Eof Then
   Arr_Channel=Rs.GetRows()
End If
Rs.Close
Set Rs=Nothing

If Isarray(Arr_Channel)= True then
   i_Class=0
   i_Special=0
   For i=0 To Ubound(ArrShowLine)
      ArrShowLine(i)=False
   Next
End if
End sub
%>