<!--#include file="../../../xwinc/config.asp"-->
<!--#include file="../../admin_check.asp"-->
<%
dim connstr
dim db
dim conn
db=""&SitePath&"xwData/"&DataName&""
'On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.Open connstr
If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "数据库连接出错，请检查连接字串。"
	Response.End
End If

Sub CloseConn()
        'On Error Resume Next
	Conn.close
	set Conn=nothing
End sub


dim connstrItem
dim dbItem
dim connItem
dbItem="database/#Item.mdb"      '采集数据库文件的位置 
'On Error Resume Next
Set connItem = Server.CreateObject("ADODB.Connection")
connstrItem="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbItem)
connItem.Open connstrItem
If Err Then
   err.Clear
   Set ConnItem = Nothing
   Response.Write "采集数据库连接出错，请检查连接字串。"
   Response.End
End If

Sub CloseConnItem()
   On Error Resume Next
   ConnItem.close
   Set ConnItem=nothing
End sub
%>