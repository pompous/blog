<!--#include file="../xwInc/conn.asp"-->
<%
Dim id
Dim Rs,Sql 
id = Replace(Trim(Request.QueryString("id")),"'","")
If Session("id"&id)<>"" Then
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From xiaowei_Article Where id="&id
	Rs.Open Sql,Conn,3,3
	If Rs.Eof And Rs.Bof Then
		Response.Write("NoData")
	Else
		Response.Write("Dig")
		Response.Write(",")
		Response.Write(Rs("Dig"))
	End If
Else
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From xiaowei_Article Where id="&id
	Rs.Open Sql,Conn,3,3
	If Rs.Eof And Rs.Bof Then
		Response.Write("NoData")
	Else
		Dim Dig
		Dig =Rs("Dig")
		Dig = Dig + 1
		Rs("Dig") = Dig
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Session("id"&id) = id
		Response.Write(Dig)
	End If
End If
%>


<%
Dim Fy_Post,Fy_Get,Fy_cook,Fy_In,Fy_Inf,Fy_Xh,aa
On Error Resume Next
Fy_In = "'|and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare|--|script"
aa="log.txt"
Fy_Inf = split(Fy_In,"|")
If Request.Form<>"" Then
For Each Fy_Post In Request.Form
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
flyaway1=""&Request.ServerVariables("REMOTE_ADDR")&","&Request.ServerVariables("URL")&"+'post'+"&Fy_post&"+"&replace(Request.Form(Fy_post),"'","*")&""
set fs=server.CreateObject("Scripting.FileSystemObject")
set file=fs.OpenTextFile(server.MapPath(aa),8,True)
file.writeline flyaway1
file.close
set file=nothing
set fs=nothing
call aaa()
End If
Next
Next
End If
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
flyaway2=""&Request.ServerVariables("REMOTE_ADDR")&","&Request.ServerVariables("URL")&"+'get'+"&Fy_get&"+"&replace(Request.QueryString(Fy_get),"'","*")&""
set fs=server.CreateObject("Scripting.FileSystemObject")
set file=fs.OpenTextFile(server.MapPath(aa),8,True)
file.writeline flyaway2
file.close
set file=nothing
set fs=nothing
call aaa()
End If
Next
Next
End If
If Request.Cookies<>"" Then
For Each Fy_cook In Request.Cookies
For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Cookies(Fy_cook)),Fy_Inf(Fy_Xh))<>0 Then
flyaway3=""&Request.ServerVariables("REMOTE_ADDR")&","&Request.ServerVariables("URL")&"+'cook'+"&Fy_cook&"+"&replace(Request.Cookies(Fy_cook),"'","*")&""
set fs=server.CreateObject("Scripting.FileSystemObject")
set file=fs.OpenTextFile(server.MapPath(aa),8,True)
file.writeline flyaway3
file.close
set file=nothing
set fs=nothing
call aaa()
End If
Next
Next
End If
Sub aaa()
Response.Write "<script>alert(""URL����"");location.href=""../"";</script>"
Response.Write "Sorry.Tel:110<br><hr>"
Response.End
end Sub
%>