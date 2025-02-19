<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwinc/Inc.asp"-->
<!--#include file="../xwInc/Function_Page.asp"-->
<!--#include file="../xwInc/function.asp"-->
<% classid = 999 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title>微博 - <%=SiteTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
	<meta name="keywords" content="<%=Sitekeywords%>" />
	<meta name="description" content="<%=Sitedescription%>" />
<link href="../css/default.css" rel="stylesheet" type="text/css">
<link href="../css/a.css" rel="stylesheet" type="text/css">
</head>
<body>
	<!--#include file="../xwinc/top.asp"-->	
	
<center>

<table border="0" width="1000" cellspacing="0" cellpadding="0" id="table1"  style="border: 20px solid #FFFFFF; padding: 0" bgcolor="#FFFFFF">
	<tr>
		<td valign="top">	
		
		


		
		
<%
set rs4 = server.CreateObject ("adodb.recordset")
sql="select UserMoney from xiaowei_User where UserName='"& UserName &"'"
rs4.open sql,conn,1,1
mymoney=rs4("UserMoney")
rs4.close
set rs4=nothing
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table2" height="40">
	<tr>
		<td align="left" background="../images/btbg.gif">&nbsp;<b><font style="font-size:16px;color:#333333;">微博</font></b></td>
	</tr>
</table>







<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="80" valign="middle"><ul>
    
    

			<%
Set mypage=new xdownpage
NoI=0
mypage.getconn=conn

mypage.getsql=server.createobject("adodb.recordset")

mypage.getsql="select top 100 * from xiaowei_2weima where yn=1 order by id desc"
mypage.pagesize=""&artlistnum&""
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
    NoI=NoI+1
%>	
	
 
            
              

 <li style="float:left;text-align:left;width:100%;">   
    
    
 <table border="0" width="100%" cellspacing="5" cellpadding="0">
		<tr>
			<td width="100%" height="40">
			<p align="left">
			&nbsp;<img border="0" src="1.png" width="11" height="11">&nbsp;
			<b><font color="#CCCCCC"><%=rs("addtime")%></font>
			</p></td>
		</tr>
		<tr>
			<td style="text-indent: 32; line-height: 200%; margin-left: 20; margin-right: 30">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("content")%> &nbsp;</td>
		</tr>
		</table>


	</li>


	<%
        rs.movenext
    else
         exit for
    end if
next
%>
</ul>　</td>
    </tr>
</table>

<br>	<table border="0" cellspacing="5" cellpadding="2" height="21" align="center"><tr><%=mypage.showpage()%></tr></table>
</td>
		
	</tr>
</table>
<!--#include file="../xwinc/bottom.asp"-->


</center>


<SCRIPT LANGUAGE=JAVASCRIPT><!-- 
if (top.location != self.location)top.location=self.location;
// --></SCRIPT>
</body>
</html>