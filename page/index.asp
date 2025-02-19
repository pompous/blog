<!--#include file="../xwInc/conn.asp"-->
<!--#include file="../xwinc/ubb.asp"-->
<!--#include file="../xwinc/Inc.asp"-->




<%
dim htmlid,id1,id2,a,b
htmlid=Request.ServerVariables("QUERY_STRING") 

id1=replace(htmlid,".html","")
id2=split(id1,"_")
on error resume next
a=id2(0)
b=id2(1)
id=a
         dim zffy
        zffy= rs("ZFFY")
        
        UserName = Request.Cookies("xiaowei")("UserName")
        UserID=Request.Cookies("xiaowei")("ID")
        set rs4 = server.CreateObject ("adodb.recordset")
        sql="select UserMoney,userface from xiaowei_User where UserName='"& UserName &"'"
        rs4.open sql,conn,1,1
        mymoney=rs4("UserMoney")
        rs4.close
        set rs4=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from xiaowei_Article where id="&a
rs.open sql,conn,1,1


if rs.eof and rs.bof then

Response.write"<script>alert(""URL错误"");location.href=""../"";</script>"
response.end
   else







set rsClass=server.createobject("adodb.recordset")
sql = "select * from xiaowei_Class where ID="&rs("ClassID")&""
rsClass.open sql,conn,1,1  
if rsClass.eof and rsClass.bof then
  call Alert("URL错误","../")
  response.end
else
  ClassName=rsClass("ClassName")
rsClass.close
set rsClass=nothing
end if

If rs("PageNum")=0 then
	Content=ManualPagination(""&rs("ID")&"",""&rs("Content")&"")
else
	Content=AutoPagination(""&rs("ID")&"",""&rs("Content")&"",rs("PageNum"))
End if
%>




<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
	<title><%=rs("Title")%> - <%=SiteTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
	<meta name="keywords" content="<%=Sitekeywords%>" />
	<meta name="description" content="<%=Sitedescription%>" />
<link href="../css/default.css" rel="stylesheet" type="text/css">
<link href="../css/a.css" rel="stylesheet" type="text/css">
<link href="../css/img.css" rel="stylesheet" type="text/css">
  <script type="text/javascript" src="../xwinc/main.js"></script>
  <script type="text/javascript" src="../xwinc/Ajaxpl.asp"></script>
</head>
<body>

	<!--#include file="../xwinc/top.asp"-->
	<center>
<hr color="#cccccc" size="4">


<table border="0" width="1000" cellspacing="0" cellpadding="0" id="table2" height="40">
			<tr>
				<td width="1000"  valign="top"  bgcolor="#ffffff">
				
				
				

				
				<table border="0" width="1000" cellspacing="0" cellpadding="0" id="table3" style="border: 20px solid #FFFFFF; padding: 0" bgcolor="#ffffff" >
			<tr>
				<td height="40" colspan="2" align="left" background="../images/btbg.gif"><b><a href="/">首页</a> >> <%If TopID>0 then Response.Write("<a href="""&SitePath&"xwclass/Class.asp?ID="&TopID&""">"&Classlist(TopID)&"</a> >> ") End if%><a href="<%=SitePath%>xwclass/Class.asp?ID=<%=rs("ClassID")%>"><%=Classlist(rs("ClassID"))%></a></b>
				　</td>
			</tr>
			<tr>
				<td height="45" colspan="2">
				<p align="center"><b><font style="font-size:16px;color:#333333;"><%=rs("Title")%></font></b></td>
			</tr>
			<tr>
				<td height="40" colspan="2">
				<p align="right"><font color="#CCCCCC">作 者：<% if rs("username")<> "" then %><% else %><%=rs("author")%><% end if %> 来 源：<%=rs("CopyFrom")%> 时 间：20<%=FormatDate(rs("DateAndTime"),4)%> <script language="javascript" src="<%=SitePath%>xwinc/hits.asp?id=<%=rs("id")%>"></script>&nbsp;&nbsp;&nbsp;&nbsp; 
				</font> 
</td>
				</tr>

			<tr>
				<td align="left" style="line-height: 250%" width="620" colspan="2" > 
				
			

				
				
					<%
				If rs("yn")=2 then 
					If rs("UserName") = xiaoweiuserName then
					
						if rs("sSaveFileSelect") = 1 then %>
				<div style="height: 200px; width: 610px; "><a href="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" target="_blank"><img src="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" border="0" width="610"></a></div><p style="text-align: center">点击图片查看大图</p>
				<% end if
				
						Response.Write(""&Content&"")
					else
						Response.Write("<div style=""margin:40px auto;text-align:center;color:#ff0000;"">无权限</div>") 
					End if
				end if
				
				if rs("yn")=1 then 
				Response.Write("<div style=""margin:40px auto;text-align:center;color:#ff0000;"">未过审</div>") 
				end if  %>
				
				<% if rs("zffy") > 0 then 
				  set rs5=server.createobject("adodb.recordset")
                  sql="select * from xiaowei_gm where   username='"& UserName &"' and articleid='"&a&"'"

                 rs5.open sql,conn,1,1
        
                 if rs5.eof and rs5.bof then
                 
                 
                           if xiaoweiusername = "" then  
      
                               Call Alert ("无权限","../")
    
                           end if 
                           
                       %>
                <% else %>
							<%	if rs("yn")= 0 then 
				if rs("sSaveFileSelect") = 1 then %>
				<div style="height: 200px; width: 610px; float:left;padding:0,0,0,0;text-align:left;overflow:hidden;"><a href="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" target="_blank"><img src="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" border="0" width="610"></a></div><p style="text-align: center">点击图片查看大图</p>				<% end if
				
				Response.Write(""&Content&"") 
					
                 end if %>
	
            
				<%	if rs("linkurl")<> "" then %>
				     <table border="0" width="100%" cellspacing="0" cellpadding="0" id="table11">
					<tr>
						<td><a href="<%=rs("linkurl")%>" target="_blank"><img src="../images/down.png" border="0"></a></td>
					</tr>
					<tr>
						<td>本站的所有资源仅供学习与参考，请勿用于商业用途。</td>
					</tr>
				</table>

				<%end if%>
				
				<%end if 
				
				else%>
				
			<%	if rs("yn")= 0 then 
			
					if rs("sSaveFileSelect") = 1 then %>
				<div style="height: 200px; width: 610px; float:left;padding:0,0,0,0;text-align:left;overflow:hidden;overflow:hidden;"><a href="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" target="_blank"><img src="<% if left(LoseHtml(rs("images")),4) <> "http" then %>..<%end if%><%=rs("images")%>" border="0" width="610"></a></div><p style="text-align: center">点击图片查看大图</p>
				<% end if
				
				Response.Write(""&Content&"") 
					
                 end if %>
	
            
				<%	if rs("linkurl")<> "" then %>
				     <table border="0" width="100%" cellspacing="0" cellpadding="0" id="table11">
					<tr>
						<td><a href="<%=rs("linkurl")%>" target="_blank"><img src="../images/down.png" border="0"></a></td>
					</tr>
					<tr>
						<td>本站的所有资源仅供学习与参考，请勿用于商业用途。</td>
					</tr>
				</table>

				<%end if%>
				
				<% end if %>
				</td>
			</tr>
			<tr>
					<td align="left" height="40" bgcolor="#FFFFFF"><%=thehead%></td>
				<td align="left" height="40" bgcolor="#FFFFFF"><%=thenext%></td>
			</tr>
	
			</table>
</td>
			</tr>
		</table>

        
<!--#include file="../xwinc/bottom.asp"-->
</center>
				
				
				
				
				
				</body>

<%

function thehead 
headrs=server.CreateObject("adodb.recordset") 
sql="select top 1 ID,Title from xiaowei_Article where id<"&id&" and ClassID="&rs("ClassID")&" order by id desc" 
set headrs=conn.execute(sql) 
response.Write("<div class='vmovies'>") 
response.Write("<div class='movernrs'>")
response.Write("<ul>")
if headrs.eof then 
response.Write("<li class='titles'>上一篇：无</li>") 
else 
a0=headrs("id") 
a1=headrs("Title")

response.Write("<li class='titles'>上一篇：<a href='"&SitePath&"xwArticle/?"&a0&".html'>"&left(LoseHtml(a1),19)&"</a></li>")  
end if 
response.Write("</ul>")
response.Write("</div>")
response.Write("</div>")
end function

function thenext 
newrs=server.CreateObject("adodb.recordset") 
sql="select top 1 ID,Title from xiaowei_Article where id>"&id&" and ClassID="&rs("ClassID")&" order by id desc" 
set newrs=conn.execute(sql)  
response.Write("<div class='vmovies'>") 
response.Write("<div class='movernrs'>")
response.Write("<ul>")
if newrs.eof then 
response.Write("<li class='titles'>下一篇：无</li>")
else 
a0=newrs("id") 
a1=newrs("Title")

response.Write("<li class='titles'>下一篇：<a href='"&SitePath&"xwArticle/?"&a0&".html'>"&left(LoseHtml(a1),19)&"</a></li>") 

end if 
response.Write("</ul>")
response.Write("</div>")
response.Write("</div>")
end function
%></html>
        
         
      <%  end if

rs.close
set rs=nothing
   

 
 %>