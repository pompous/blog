<!--#include file="../xwinc/conn.asp"-->
<!--#include file="../xwinc/Function_Page.asp"-->
<!--#include file="../xwinc/md5.asp"-->
<!--#include file="../xwinc/Inc.asp"-->
<%
id=CheckStr(Trim(request.QueryString("id")))
If id="" then
	Response.write"<script>alert(""URL错误"");location.href=""../"";</script>"	
end if
set rsclass=server.createobject("adodb.recordset")
sql="select * from xiaowei_Class where id="&id&""
rsclass.open sql,conn,1,1
if rsclass.eof and rsclass.bof then
Call Alert("URL错误","../")
else
if rsclass("Link")=1 then
Response.Redirect ""&rsclass("Url")&""
End if
classid = ""&rsclass("id")&"" 

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title><%=rsclass("ClassName")%> - <%=SiteTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk2312" />
	<meta name="keywords" content="<%=Sitekeywords%>" />
	<meta name="description"  />
<link href="../css/default.css" rel="stylesheet" type="text/css">
<link href="../css/a.css" rel="stylesheet" type="text/css">
<link href="../css/img.css" rel="stylesheet" type="text/css">
<script src="../js/jquery-1.8.3.min.js" type="text/javascript"></script>
<script type="text/javascript" src="../js/jquery.masonry.js"></script>
<script type="text/javascript" src="../js/jquery.infinitescroll.js"></script>




</head>

<body>
	<!--#include file="../xwinc/top.asp"-->
	<center>
	
<%If  rsclass("TOPid")=-1 then
	
			Sqlpp ="select * from xiaowei_Class Where Topid=-1 AND id="&id&""  
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof
%>	<hr color="#cccccc" size="4"><table border="0" width="100%"  id="table2" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
		<tr>

			<td valign="TOP">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table17" style="border: 20px solid #FFFFFF; padding: 0" bgcolor="#FFFFFF" height="121">
				<tr>
					<td height="40" align="left" background="../images/btbg.gif"><b><font style="font-size:16px;color:#666666;"><%=rspp("ClassName")%></font></b></td>
				</tr>
				<tr>
					<td valign="top" align="left" style="line-height: 200%"><%=rspp("README")%></td>
				</tr>
			</table>
			

			</td>
			<td width="6" bgcolor="#cccccc">　</td>
			<td width="340" valign="top">
			<table border="0" width="340" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" id="table19">
				<tr>
					<td><table border="0" width="340" cellspacing="0" cellpadding="0" id="table9" style="border: 20px solid #FFFFFF; padding: 0" bgcolor="#ffffff" >
					<tr>
						<td height="40"align="left" background="../images/btbg.gif"><b>站长推荐：</b></td>
					</tr>
					<tr>
						<td align="left" style="line-height: 150%"><script src="../js/right.asp?topType=IsHot&classNO=&num=12&maxlen=18&showdate=0&showhits=0&showClass=0"></script></td>
					</tr>
				</table><hr color="#cccccc" size="4">
</td>
				</tr>
			</table>
							<table border="0" width="100%" height="250" cellspacing="0" cellpadding="0" id="table10" style="border: 20px solid #FFFFFF; padding: 0">
					<tr>
						<td><%Call ShowAD(21)%></td>
					</tr>
				</table>
				
				<hr color="#cccccc" size="4">
			<br>

</td>
		</tr>
		</table><hr color="#cccccc" size="4">
						<%
			Rspp.Movenext   
      		Loop
   			rspp.close
   			set rspp=nothing
%> 

<% ELSE %>
	
<hr color="#cccccc" size="4">
 <table  width="1000"   id="table9" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
	<tr> <!-- 左边开始-->
		<td valign="top">
		
		 <table border="0" width="100%" cellspacing="0" cellpadding="0" id="table21" style="border: 20px solid #FFFFFF; padding: 0" background="listbg.gif">
			<tr>
				<td>		
				
                    <table border="0" width="100%" cellspacing="0" cellpadding="0" id="table18" bgcolor="#FFFFFF"  height="40" background="../images/btbg.gif">
	<tr>
		<td align="left" width="180"><b><font style="font-size:16px;color:#666666;"><%=rsclass("ClassName")%></font></b></td>
		<td  align="left" height="40"><font color="#CCCCCC"><%=rsclass("readme")%></font></td>
	</tr>
                  </table>

<br>
          <ul>           
				<%
Set mypage=new xdownpage
NoI=0
mypage.getconn=conn
mypage.getsql="select * from xiaowei_Article Where Classid="&rsclass("id")&" and yn=0 order by IsTop desc,id desc"
mypage.pagesize=""&artlistnum&""
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
    NoI=NoI+1
%>
		<li style="float:left;text-align:center;width:100%;line-height:35px;padding:0 0 5px 0;">

		
		<table border="0" width="100%" cellspacing="0" cellpadding="0" >
			<tr>
				<td>
								<table border="0" width="100%"  id="table11" cellspacing="0" cellpadding="0" style="border: 0px solid #666666; padding: 0">
					<tr>
					<li class="green">
						<td height="40" align="left" width="13%">
						<p style="line-height: 100%;margin-left:0px; margin-right:3px" align="right"><font color="#A3D8F7" size="2" style="background:#666666;color:#ffffff;padding:5px 8px;"><%=FormatDate(rs("DateAndTime"),5)%></font></td>
						
						<td height="40" align="left" width="95">
						<p align="center"><%If rs("Images")<>"" then %><img width="70" height="70" alt="<%=rs("title")%>"  src="<%=SitePath%><%=SiteUp%><%If AspJpeg=1 Then %>/<%Else%>/<%end if%><%=rs("Images")%>"  style="border: 1px solid #cccccc; padding: 5px; background-color: #f5f5f5"/><% else %><img border="0"  width="30" height="30"><% END IF %></td>
						<td height="40" align="left" width="72%"><a href="../page/?<%=rs("id")%>.html" target="_blank"><font  style="font-size:18px;color:#666666;"><b><%Response.Write(""&left(LoseHtml(rs("title")),50)&"")%></b></font></a>&nbsp;&nbsp;<%if rs("zffy") > 0 then%><span style="background:#FF6600;color:#ffffff;padding:0px 4px;">无权限</span><% end if %>&nbsp;&nbsp;<font color="#A3D8F7" class="daohan2"><%=rs("copyfrom")%></font><span><p style="line-height: 200%;margin-left:3px; margin-right:3px" align="left"><%=left(LoseHtml(rs("Content")),60)%></span><p align="left">
						</td>
						</li>
					</tr>
				</table>
				</td>
			</tr>
		</table>
<br>
</li>	
					<%
        rs.movenext
    else
         exit for
    end if
next
%>
 </ul>
</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
      <%end if%>
        <%end if%>
			</td>
		</tr>
	</table>
	<center>
<!--#include file="../xwinc/bottom.asp"-->

</body>
</html>