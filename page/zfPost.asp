<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwinc/Inc.asp"-->

			<%
	dim UserName,ArticleID,zhifu
	UserName = trim(request.form("UserName"))

	ArticleID = trim(request.form("ArticleID"))
	zhifu = request.form("zhifu")


	
  
	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from xiaowei_gm"
	rs.open sql,conn,1,3
	
if UserName="" then
	Call Alert ("����!","../")
elseif articleid="" then
	Call Alert ("����!","../")
	elseif zhifu="" then
	Call Alert ("����!","../")


end if

		rs.AddNew 
		rs("UserName")			=UserName
		rs("ArticleID")				=ArticleID
		rs("zhifu")			=zhifu

        rs("yn")      =1
		rs("IP")				=Request.ServerVariables("REMOTE_ADDR")
	    rs("PostTime")     =  now()
		rs.update
	
		Response.Write("<script language=javascript>alert('���֧������ˢ��ҳ�棡');history.go(-1);</script>")
sqlmoney="update xiaowei_User set UserMoney = UserMoney-'"&rs("zhifu")&"' where UserName='"&UserName&"'"
			conn.execute(sqlmoney)

		rs.close
		Set rs = nothing
   %>