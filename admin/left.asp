
<%
Dim OK2,xiaoweimanage2
OK2=session("xiaoweiAdmin")
xiaoweimanage2=Request.Cookies("xiaoweimanage")("UserName")
if OK2="" and xiaoweimanage2="" then
	Response.Write("<script language=javascript>this.top.location.href='Admin_Login.asp';</script>")
	Response.end
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--

img{	border: none;}
form{	margin: 0px;	padding: 0px;}
input{	color: #000000;	height: 22px;	vertical-align: middle;}
textarea{	width: 80%;	font-weight: normal;	color: #000000;}
a{	text-decoration: underline;	color: #666666;}
a:hover{	text-decoration: none;}
.menuDiv,.menuDiv1{	background-color: #FFFFFF;}
.menuDiv1{	postion:relative;bottom:0px;top:50;}
.menuDiv h3,.menuDiv1 h3{
	font-weight:bold;font-size:14px;color:#ffffff;
	padding:8px 0 3px 15px;
	background-color:#006699;
	margin: 0px;cursor:pointer;
	border-top:#9FBCD4 1px solid; 
	height: 20px;
}
.menuDiv1 h3 {color:#ffcc00;height:30px;}
.menuDiv ul,.menuDiv1 ul{	margin:5px 0 0 0px;	padding: 5px;	list-style-type: none;}
.menuDiv ul li,.menuDiv1 ul li{
	color: #ffffff;
	background-color:#FFFFFF;
	padding: 5px 5px 5px 10px;
	font-size: 14px;
	height: 20px;border-bottom:1px solid #ffffff;
}
.menuDiv ul li a,.menuDiv1 ul li a{color: #333333;	text-decoration: none;border-bottom:#ccc 1px solid;border-right:#ccc 1px solid;padding:5px 10px;}
.menuDiv ul li a:hover,.menuDiv1 ul li a:hover{color: #009900;text-decoration:none;border-bottom:#ccc 1px solid; border-right:#ccc 1px solid;padding:5px 10px;}
.red{	color:#FF0000;}
-->
</style>

</head>

<body topmargin="0" leftmargin="0">
<table width="200" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td >
	<div class="menuDiv"> 
	  <h3>基本管理</h3> 	  
	  <ul> 	    
	<li><a href="Admin_Setting.asp" target="_self">网站配置</a>
	<a href="Admin_Admin.asp" target="_self">管 理 员</a></li>
	<li><a href="Admin_Ad.asp" target="_self">广告管理</a>
	<a href="Admin_weibo.asp" target="_self">微博管理</a></li>
	  </ul>
	</div>

	<div class="menuDiv"> 
	  <h3>文章管理</h3> 
	  <ul> 	    
		<li><a href="Admin_Class.asp" target="_self">栏目管理</a> <a href="Admin_Class.asp?action=add" target="_self">添加</a></li>
		<li><a href="Admin_Article.asp" target="_self">文章管理</a> <a href="Admin_Article.asp?action=add" target="_self">添加</a></li>
	  </ul>
	</div>
    <div class="menuDiv"> 
	  <h3>数据库管理</h3> 
	  <ul> 	    
		<li><a href="Admin_data.asp?action=SpaceSize" target="_self">空间查看</a> <a href="Admin_data.asp?action=CompressData" target="_self">数据压缩</a></li>
        <li><a href="Admin_data.asp?action=BackupData" target="_self">数据备份</a> <a href="Admin_data.asp?action=RestoreData" target="_self">数据恢复</a></li>

	  </ul>
	</div>
	<div class="menuDiv"> 
	  <h3>　</h3> </div>
          </td>
      </tr>
    </table></td>
  </tr>
</table>


</body>
</html>
<%end if%>