<!--#include file="Safety.asp"-->
<!--#include file="char.inc"-->
<html>
<head>
<style type="text/css">
td{font-size:12px;}
</style>

<script language="javascript">

function checkspace(astr)

{

bstr=""
cd=astr.length
for(i=0;i<cd;i++)
 { if(astr.charAt(" ")>=0)
    {bstr=bstr+" "}
 }

if(bstr==astr)
 {
return true;
}
else{return false;}

}


function checkform() {

if(checkspace(form1.msgre.value))

{
   alert ("请输入回复内容!");
   form1.msgre.value="";
	form1.msgre.focus();
	return false;

}





}


</script>

</head>


<body>

<script language="javascript" src="../top.js"></script>

<%
	filePath = "siva.xml"
	Set objXML = Server.CreateObject("Microsoft.XMLDOM")
   'Set objXML = server.CreateObject("Msxml2.DOMDocument")
        objXML.async = False
        loadResult = objXML.Load(server.MapPath(filePath))
        If Not loadResult Then
             Response.Write ("加载XML文件出错!")
               Response.End
         End If		

      id =  Request.QueryString("ID")
      Set objNodes = objXML.selectSingleNode("xml/guestbook/item[id ='" & id & "']")



      if Not IsNull(objNodes) then
	  reply=objNodes.childNodes(7).text
       msgre = htmlencode2(Request.form("msgre"))
       if msgre <> "" then
		objNodes.childNodes(7).text  = msgre
		objNodes.childNodes(8).text  = now()
		objXML.save(server.MapPath(filePath))		
		Set objXML=nothing
		Response.Write "<script>window.alert('成功回复！');window.location='!!!!!1124lyadmin.asp';</script>"
		response.end
       end if
      end if
%>


<table width='729' border='0' align='center' cellpadding='0' cellspacing='0'>
<tr><td height='20' colspan='2' style='background-position: 0% 0%; font-size:12px;color:#000000;BORDER-RIGHT:1px solid #aaaaaa;padding-top:3px; background-image:url(&#039;images/bg1.gif&#039;); background-repeat:repeat; background-attachment:scroll'>
	<font color="#FFFFFF">当前位置：首页 >> 
    留言板</font></td></tr>
</table>

<table width='729' border='0' align='center' cellpadding='0' cellspacing='0'>
<tr><td height='10' colspan='2'></td></tr>
</table>
  
  
	  
<div align="center">
  
  
	  
<TABLE width=728 border=0 cellPadding=0 cellSpacing=0 style="BORDER: 2px solid #A5D318;">
  <TBODY>
    <TR style="BACKGROUND-COLOR: #f2fdee"> 
      <TD width=74 height=18>&nbsp;</TD>
      <TD width=572>　</TD>
      <TD width=80>　</TD>
    </TR>
    <TR style="BACKGROUND-COLOR: #f2fdee">
      <TD height=181>　</TD>
      <TD valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <!--DWLayoutTable-->
        <form name="form1" method="post" action="reply.asp?ID=<%=id%>" onSubmit="return checkform()">
            <tr> 
              <td width="80" height="133" align="right" valign="top">管理员回复：</td>
              <td width="492" valign="top"> <textarea name="msgre" cols="50" rows="9" id="msgre"><%=reply%></textarea> 
                <font color=#ff0000>**</font></td> 
            </tr> 
            <tr>  
              <td height="34" colspan="2" align="center" valign="middle"> <input name="submit" type="submit" class="input" value="管理员回复"> 
                &nbsp;  
                <input type="reset" value=" 重 填 " class="input">&nbsp;  
                <input type="button" value=" 返 回 " class="input" onclick="javascript:window.history.go(-1);"></td> 
            </tr> 
          </form> 
<%  
set objXML = nothing  
%> 
        </table></TD> 
      <TD>　</TD> 
    </TR> 
    <TR style="BACKGROUND-COLOR: #f2fdee"> 
      <TD height=25>　</TD> 
      <TD>　</TD> 
      <TD>　</TD> 
    </TR> 
  </TBODY> 
</TABLE> 
 
</div>
 
<script language="javascript" src="http://www.huaid.cn/blog/blogb1/blogbb.js"></script> 
 
</BODY></HTML> 
 
 
