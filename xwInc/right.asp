<!--#include file="../xwInc/conn.asp"-->
<%

dim path,classID,NclassID,NclassID1,showNclass,kind,dateNum,maxLen,listNum,bullet
dim hitColor,new_color,old_color

topType = Request("topType")
If Request("ClassNo") <> "" then
ClassNo = split(Request("ClassNo"),"|")
on error resume next
NClassID = ClassNo(0)
NClassID1 = ClassNo(1)
End if

num = request.querystring("num")
maxlen = Request.querystring("maxlen")
showdate = Request("showdate")
showhits = Request("showhits")
showClass = Request("showClass")

bullet="<img src='../images/dy.gif'>&nbsp;"		 '����ǰ��ͼƬ�����
hitColor="#cccccc"   '���������ɫ
hitColor1="#eeeeee" 
new_color="#FF0000"  '���������ڵ���ɫ
old_color="#999999"  '���������ڵ���ɫ
new_yan="background:#FF6600;color:#ffffff;padding:0px 4px;"

dim rs,sql,str,topic
Path="http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"right.asp","")

set rs=server.createObject("Adodb.recordset")
sql = "Select top "& num &" ID,Title,zffy,Author,ClassID,DateAndTime,Hits,IsTop,IsHot from xiaowei_Article Where yn = 0"

	If NclassID<>"" and NclassID1="" then
		If XIAOWEI_MyID(NclassID)="0" then
			SQL=SQL&" and ClassID="&NclassID&""
		else
			MyID = Replace(""&XIAOWEI_MyID(NclassID)&"","|",",")
			SQL=SQL&" and ClassID in ("&MyID&")"
		End if
	elseif NclassID<>"" and NclassID1<>"" then
		MyID = Replace(""&Request("ClassNo")&"","|",",")
		SQL=SQL&" and ClassID in ("&MyID&")"
	End if
	
select case topType
	case "new" sql=sql&" order by id desc"
	case "hot" sql=sql&" order by hits desc,ID desc"
	case "IsHot" sql=sql&"and IsHot = 1 order by ID desc"
	case "2" sql=sql&" DATEDIFF('d',intime,Now())<="&dateNum&" order by hits desc,Unid"
end select

set rs = conn.execute(sql)
if rs.bof and rs.eof then 
str=str+"û�з�������������"
else

rs.movefirst
do while not rs.eof
	topic=Left(LoseHtml(rs("title")),maxlen)
	topic=replace(server.HTMLencode(topic)," ","&nbsp;")
	topic=replace(topic,"'","&nbsp;")
	str=str+bullet
	if showClass = 1 then
		str=str+"<a href='../menu/index.asp?ID="&rs("ClassID")&"' target='_blank'>"&Classlist(rs("ClassID"))&"</a>&nbsp;|&nbsp; "
	end if
	str=str+"<a href='../"
		str=str+"page/index.asp?"+Cstr(rs("ID"))+""

	str=str+"' target='_blank'  title='"&replace(replace(server.HTMLencode(rs("title"))," ","&nbsp;"),"'","&nbsp;")&"') >"+Topic+"</a>"
	if showdate <> 0 then
		str=str & "��<font color="
			if rs("DateAndTime")>=date then
				str=str & new_color
		 	else
				str=str & old_color
			end if
			str=str &">" & FormatDate(rs("DateAndTime"),""&showdate&"")&"</font>��"
	end if
	if showhits = 0 then
        If rs("zffy")> 0 then
	 str=str&"&nbsp;<font style="& new_yan &">&nbsp;��&nbsp; </font> "
	end if
	end if
	str=str&"<p><font color="& hitcolor1 &">-------------------------------------------------</font></p>"
	rs.moveNext
loop
end if
rs.close : conn.close
set rs=nothing : set conn=nothing

response.write "document.write ("&Chr(34)&str&Chr(34)&");"
%>