
<%
Dim OK1,xiaoweimanage1
OK1=session("xiaoweiAdmin")
xiaoweimanage1=Request.Cookies("xiaoweimanage")("UserName")
if OK1="" and xiaoweimanage1="" then
	Response.Write("<script language=javascript>this.top.location.href='../Admin_Login.asp';</script>")
	Response.end
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
body {	margin-left: 0px;	margin-top: 0px;	margin-right: 0px;	margin-bottom: 0px;}
body,td,th,div,a,h3,textarea,input{font-family: "΢���ź�","Times New Roman","Courier New";font-size: 14px;color: #333333;}
.link {background:#eeeeee;color:#006699;padding:5px 8px;}


a{text-decoration:none;}
-->
</style>
</head>

<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="60" bgcolor="#006699">
  <tr>
    <td width="380" ><img src=../images/logo.jpg width="371" height="40"></td>
    <td  align="left">
    <a href="../index.asp" target="_self" class="link">��̨��ҳ</a>
    <a href="../Admin_Article.asp?action=add" target="_self"  class="link">��������</a>
   

    </td>
    <td width="185" style="text-align:right;">
    <p style="text-align: center">
    <a href="../../" target="_blank"  class="link">��վ��ҳ</a>
    <a href="../index.Asp?Sub=Logout" target="_top"  class="link">�˳���½</a></td>
  </tr>
</table>
</body>
</html>
<%end if%>