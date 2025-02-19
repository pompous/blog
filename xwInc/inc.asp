

<%

Function Menu
	Response.Write("<table  height=""40"" cellspacing=""0"" cellpadding=""0"" class=""menu""><tr>") & VbCrLf
	if classid = 0 then
	Response.Write("<td><a href="""&SitePath&""" class=""select"">Ê× Ò³</a></td>") & VbCrLf
	else
	Response.Write("<td><a href="""&SitePath&""">Ê× Ò³</a></td>") & VbCrLf
  end if
  
set rs8=conn.execute("select * from xiaowei_class Where IsMenu=1 order by num asc")
do while not rs8.eof
NoI=NoI+1
	
	if classid = ""&rs8("id")&"" then
	Response.Write("<td><a href="""&SitePath&"menu/index.asp?id="&rs8("ID")&"""  class=""select"" target="""&rs8("target")&"""  onmouseover=""mouseover(this, "&rs8("id")&")"" onmouseout=""mouseout()"">"&rs8("className")&"</a></td>") & VbCrLf
	else
	Response.Write("<td><a href="""&SitePath&"menu/index.asp?id="&rs8("ID")&"""  target="""&rs8("target")&"""  onmouseover=""mouseover(this, "&rs8("id")&")"" onmouseout=""mouseout()"">"&rs8("className")&"</a></td>") & VbCrLf
  end if

	


  
  
rs8.movenext
loop
  
	if classid = 1 then
	Response.Write("<td><a href="""&SitePath&"weibo/"" class=""select"" title=""Î¢²©"">Î¢²©</a></td>") & VbCrLf
	else
	Response.Write("<td><a href="""&SitePath&"weibo/"" title=""Î¢²©"">Î¢²©</a></td>") & VbCrLf
  end if



		Response.Write("</tr></table>") & VbCrLf
rs8.close



set rs8=conn.execute("select * from xiaowei_class Where IsMenu=1 order by num asc")
do while not rs8.eof
NoI=NoI+1 %>
  <%If Mydb("Select Count([ID]) From [xiaowei_Class] Where TopID="&rs8("ID")&"",1)(0)>0 then %>
  <div id="menu<%=rs8("id")%>" class="menu-list" onmouseover="_mouseover()" onmouseout="_mouseout()">
          <ul><%
	Sqlpp ="select * from xiaowei_Class Where TopID="&Rs8("ID")&" order by num"   
  Set Rspp=server.CreateObject("adodb.recordset")   
  rspp.open sqlpp,conn,1,1
	Do while not Rspp.Eof
  %>
            <li><a href="<%=SitePath%>menu/index.asp?ID=<%=rspp("ID")%>"><span><%=rspp("ClassName")%></span></a></li><%
	Rspp.Movenext   
  Loop
  rspp.close
  set rspp=nothing
  %> 
          
        </ul></div>     
  <%else%>
  <div id="menu<%if rs8("id") <> 999 or rs8("id") <> 998 or rs8("id") <> 996 then %><%=rs8("id")%><%end if%>"></div>
	<%end if%>
<%
rs8.movenext
loop
rs8.close






End function
%>