<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwInc/md5.asp"-->
<%
dim adminname
dim adminpwd

if request("action")="adminlogin" then
adminname	=trim(Request.form("adminname"))
adminpwd	=trim(Request.form("adminpwd"))
adminname	=Replace(adminname,"'","")
adminpwd	=Replace(adminpwd,"'","")
adminpwd	=md5(adminpwd,16)

mycode = trim(request.form("code"))
	if adminname="" or adminpwd="" then
	Response.Write("<script language=javascript>alert('请输入用户名和密码！');javascript:history.back();</script>") 
	end if
	if mycode<>Session("getcode") then
	Response.Write("<script language=javascript>alert('请输入正确的验证码！');javascript:history.back();</script>") 
	Response.End 
	end if
	
set rs=server.createobject("ADODB.Recordset")
sql="select * from xiaowei_Admin where Admin_Name='"&adminname&"' and Admin_Pass='"&adminpwd&"'"
rs.open sql,conn,1,3
If Not rs.Eof Then 
   session("xiaoweiAdmin")=rs("Admin_Name")
   'Response.Cookies("xiaoweimanage").Expires=Date+1
   Response.Cookies("xiaoweimanage")("UserName") = rs("Admin_Name")
   
   rs("Admin_Time")		= Now
   rs("Admin_IP")		= Request.ServerVariables("REMOTE_ADDR")
   rs.update
   
   response.Redirect "Index.asp"
   rs.close
   set rs=nothing
else	
   rs.close
   set rs=nothing
   Response.Write("<script language=javascript>alert('您输入的用户名或密码不正确!默认用户名admin密码123456');history.back(1);</script>")
end if

end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>管理员登录</title>
<link href="images/Admin_css.css" type=text/css rel=stylesheet>
<link rel="shortcut icon" href="<%=SitePath%>xwimg/myfav.ico" type="image/x-icon" />
</head>


<body bgcolor="#3D71BA">
<p>　</p>
<form name="form1" method="post" action="?action=adminlogin">
    
  <table width="633" height="369" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/loginbg.jpg" style="background-repeat:no-repeat">
    <tr> 
      <td width="634" height="246">　</td>
   </tr>
    <tr> 
      <td height="123" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="31" colspan="1" align="right"><input name="AdminName" type="text" id="AdminName"  style="width:75px; height:15px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1; border-color:#3D71BA" onFocus="this.select(); " onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" maxlength="20"></td>
		  
        </tr>
        <tr>
          <td width="51%" height="47"  align="right"><input name="adminpwd" type="password" id="adminpwd" style="width:75px; height:15px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1;border-color:#3D71BA" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          <td width="21%"  align="right"><input name="code" type="text" id="code" size="8" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1;height:15px;border-color:#3D71BA" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          <td width="28%">&nbsp;<img src="../xwInc/code.asp" border="0" alt="看不清楚请点击刷新验证码" style="cursor : pointer;" onClick="this.src='../xwInc/code.asp'"/>&nbsp;<input type="image" name="imageField" src="Images/submit.jpg"></td>
        </tr>
      </table></td>
   </tr>
  </table>
  <br>
  <p align="center"><br>
  </p>
</form>

</body>
</html>