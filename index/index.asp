<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwInc/md5.asp"-->
<!--#include file="../xwinc/Inc.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title> <%=SiteTitle%> - <%=Sitedescription%>,<%=Siteurl%></title>
<meta name="keywords" content="<%=Sitekeywords%>" />
<meta name="description" content="<%=Sitedescription%>" />
<meta name="author" content="Pompous">
<link rel="shortcut icon" href="favicon.ico" /> 
<link href="../css/default.css" rel="stylesheet" type="text/css">
<link href="../css/a.css" rel="stylesheet" type="text/css">
<link href="../css/img.css" rel="stylesheet" type="text/css">
</head>
<body>
	<!--#include file="../xwinc/top.asp"-->
	<p>
	<center>

			<td  valign="top" width="100%" height="200" bgcolor="#ffffff">

			<table width="1000" height="190" cellspacing="0" cellpadding="0" style="border: 20px solid #FFFFFF; padding: 0" bordercolor="#FFFFFF" id="table21" bgcolor="#FFFFFF">
				<tr>
					<td height="40" background="../images/btbg.gif" align="left">
					<b><font style="font-size:16px;color:#666666;">Î¢²©</font></b></td>
				</tr>
				<tr>
					<td align="left" valign="top">	<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select Top 5 * from xiaowei_2weima where yn = 1 order by ID desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err) 
%>

<p style="float:left;text-align:left;width:100%;line-height:30px;padding:3px 0 0 0;"><img border="0" src="../weibo/1.png" width="11" height="11"><b>&nbsp;<a href="../weibo/" target="_blank"><%=left(LoseHtml(rs1("username")),50)%></a></b> <br>
<font color="#006699"><a href="../weibo/" target="_blank"><%=left(LoseHtml(rs1("Content")),100)%></a></font></p> 



<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>


</td>
				</tr>
			</table>
			</td>







				<td  valign="top" width="1000" height="100%" bgcolor="#ffffff">

			<table border="0" width="1000" height="100%" id="table17" cellspacing="0" cellpadding="0" style="border: 20px solid #FFFFFF; padding: 0" bgcolor="#FFFFFF">
				<tr>
					<td height="40" align="left">
					<b><font style="font-size:16px;color:#666666;">×îÐÂÎÄÕÂ</font></b></td>
				</tr>
				<tr>
					<td valign="top"style="line-height: 400%" ><script src="../js/index.asp?topType=new&classNO=&num=20&maxlen=20&showdate=0&showhits=1&showClass=1"></script></td>
					</tr>
			</table>
			</td>

<!--#include file="../xwinc/bottom.asp"-->
<SCRIPT LANGUAGE=JAVASCRIPT><!-- 
if (top.location != self.location)top.location=self.location;
// --></SCRIPT>
<script src="js/jquery.js"></script> 
<script src="js/main.js"></script>
</body>
</html>