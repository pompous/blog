<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Call Main()
'�ر����ݿ�����
Call CloseConn()
Call CloseConnItem()
Sub Main%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS ϵͳ����</title>
<link href="../Images/admin_css.css" rel="stylesheet" type="text/css" />
</head>

<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">��</td>
		<td valign="top"><br>	
	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">�����Ŀ</a> >> <font color=red>��������</font> >> �б����� >> �������� >> �������� >> �������� >> �������� >> ���</td>
  </tr>         
</table>     
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
<form method="post" action="Admin_ItemAddNew2.asp" name="myform">
    <tr> 
      <td colspan="2" class="admintitle">�������Ŀ--��������</td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>��Ŀ���ƣ�</strong></td>
      <td width="75%" class="b1_1">
	  <input name="ItemName" type="text" size="27" maxlength="30">&nbsp;&nbsp;<font color=red>*</font>�磺���������������� �ȼ�������      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong> ������Ŀ��</strong></td>
      <td width="75%" class="b1_1"><select ID="ClassID" name="ClassID"><%call Admin_ShowChannel_Option(1)%></select>      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> ��վ���ƣ�</strong></td>
      <td width="75%" class="b1_1">
	  <input name="WebName" type="text" size="27" maxlength="30">      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> ��վ��ַ��</strong></td>
      <td width="75%" class="b1_1"><input name="WebUrl" type="text" size="49" maxlength="150">      </td>
    </tr>
   <tr> 
      <td width="20%" class="b1_1"><strong> ��վ��¼��</strong></td>
      <td class="b1_1">
		<input name="LoginType" type="radio" class="noborder" onClick="Login.style.display='none'" value="0" checked>����Ҫ��¼<span lang="en-us">&nbsp;
		</span>
		<input name="LoginType" type="radio" class="noborder" onClick="Login.style.display=''" value="1">���ò���</td>
    </tr>
   <tr id="Login" style="display:none"> 
      <td width="20%" class="b1_1"><strong> ��¼������</strong></td>
      <td class="b1_1">
        ��¼��ַ��<input name="LoginUrl" type="text" size="40" maxlength="150" value=""><br>
        �ύ��ַ��<input name="LoginPostUrl" type="text" size="40" maxlength="150" value=""><br>
        �û�������<input name="LoginUser" type="text" size="30" maxlength="150" value=""><br>
        ���������<input name="LoginPass" type="text" size="30" maxlength="150" value=""><br> 
		ʧ����Ϣ��<input name="LoginFalse" type="text" size="30" maxlength="150" value=""></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1"><strong>��Ŀ��ע��</strong></td>
      <td width="75%" class="b1_1"><textarea name="ItemDemo" cols="49" rows="5"></textarea></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input name="Cancel" type="button" id="Cancel" value=" ��&nbsp;&nbsp;�� " onClick="window.location.href='Admin_ItemManage.asp'">
        &nbsp; 
        <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��"></td>
    </tr>
</form>
</table>   
		

				
		</td>
	</tr>
</table>


</body>

</html>
<%end Sub%>