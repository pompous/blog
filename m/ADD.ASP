<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
name = htmlencode2(Request.Form("name"))
qq = htmlencode2(Request.Form("qq"))
email = htmlencode2(Request.Form("email"))
sex = htmlencode2(Request.Form("sex"))
content = htmlencode2(Request.Form("content"))
bad=split(badstr,"|")
for i=0 to UBound(bad)
name=Replace(name,bad(i),"**")
content=Replace(content,bad(i),"**")
next	
'if len(name)>cint(""&n_num&"") or len(content)>cint(""&c_num&"") then
'Response.Write '"<script>window.alert('字数超过限止,称昵不能超过"&n_num&"个汉字,留言不能超过"&c_num&"个汉字！');javascript:window.close();</script>"
'Response.End
'end if
	
if name <> "" and content <> "" then

strSourceFile = Server.MapPath("siva.xml")
Set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
If objXML.parseError.ErrorCode <> 0 Then 
objXML.loadXML "<?xml version=""1.0"" encoding=""gb2312"" ?><xml><guestbook></guestbook></xml>"
End If
Set objRootlist = objXML.documentElement.selectSingleNode("guestbook")
If objRootlist.hasChildNodes then
id = objRootlist.lastChild.firstChild.text + 1
Else
id=1
End If
		
Set oListNode = objXML.documentElement.selectSingleNode("guestbook").AppendChild(objXML.createElement("item"))
Set oDetailsNode = oListNode.appendChild(objXML.createElement("id"))
oDetailsNode.Text = id
Set oDetailsNode = oListNode.appendChild(objXML.createElement("name"))
oDetailsNode.Text = name

Set oDetailsNode = oListNode.appendChild(objXML.createElement("qq"))
oDetailsNode.Text = qq

Set oDetailsNode = oListNode.appendChild(objXML.createElement("email"))
oDetailsNode.Text = email
Set oDetailsNode = oListNode.appendChild(objXML.createElement("sex"))
oDetailsNode.Text = sex

Set oDetailsNode = oListNode.appendChild(objXML.createElement("content"))
oDetailsNode.Text = content

Set oDetailsNode = oListNode.appendChild(objXML.createElement("addtime"))
oDetailsNode.Text = now()

Set oDetailsNode = oListNode.appendChild(objXML.createElement("reply"))
oDetailsNode.Text = ""
Set oDetailsNode = oListNode.appendChild(objXML.createElement("hftime"))
oDetailsNode.Text = ""
Set oDetailsNode = oListNode.appendChild(objXML.createElement("ip"))
oDetailsNode.Text = ip
objXML.save(strSourceFile)		
Set objXML=nothing
	
Response.Write "<script>window.alert('谢谢您的留言！');window.location='index.asp';</script>"
Response.End

end if
%>