<!--#include file="../xwinc/conn.asp"-->
<% 
dim id 
id=request("id") 

If Request("Page") = "" then 
sql="update xiaowei_Article set hits = hits + 1 where ID= "&ID&""
conn.execute(sql)
End if

	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From xiaowei_Article Where id="&id&""
	Rs.Open Sql,Conn,3,3
'处理数据库hit+1代码省略..
hits=rs("hits")
'下面显示文章点击数，如果不显示就不要输出了
response.write("document.write('点击：" & hits & " 次')") 
%>