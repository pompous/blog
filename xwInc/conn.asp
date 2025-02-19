<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<!--#include file="Function.asp"-->
<%
	dim conn,connstr,db
	db=""&SitePath&"xwdata/"&DataName&""
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)

	conn.Open connstr
	If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "<div style='margin:100px;font-size:14px;text-align:center'>数据库连接出错</div>"
	Response.End
	End If

xiaoweiuserID=Request.Cookies("xiaowei")("ID")
xiaoweiuserName=Request.Cookies("xiaowei")("UserName")


	
Sub ShowAD(ID)
	set rsad=conn.execute("select * from xiaowei_Ad Where ID = "&ID&"")
	If Not rsad.Eof Then 
	Response.Write(""&rsad("Content")&"")
	End if
	rsad.close
	set rsad=nothing    
End Sub

Sub Label(ID)
	set rsLabel=conn.execute("select * from xiaowei_Label Where ID = "&ID&"")
	If Not rsLabel.Eof Then 

	Response.Write(""&rsLabel("Content")&"")
	End if
	rsLabel.close
	set rsLabel=nothing    
End Sub



%>