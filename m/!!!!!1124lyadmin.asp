<!--#include file="safety.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<title><%=webname%>_��̨����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
td{font-size:12px;}
body {
	background-color: #006991;
}
.STYLE1 {color: #FFFFFF}
</style>
<script>

function checkImgSize(imgIn)
{ 

var maxWidth=140;
var maxHeight=90;

if(imgIn.width>maxWidth)
{imgIn.width=maxWidth;}

if(imgIn.height>maxHeight) 
{imgIn.height=maxHeight;}

}
</script>
</head>
<body topmargin="10">
<table width='729' border='0' align='center' cellpadding='0' cellspacing='0'>
  <tr>
    <td height='10' colspan='2'><span class="STYLE1"><strong>��ǰλ�ã���ҳ &gt;&gt; ���԰��̨<a href="http://www.51.la/report/0_menu.asp?id=523746" target="_blank"><img src="http://img.users.51.la/523746.asp" alt="&#x6211;&#x8981;&#x5566;&#x514D;&#x8D39;&#x7EDF;&#x8BA1;" width="4" height="4" border="0" style="border:none" /></a></strong></span></td>
  </tr>
</table>
<TABLE width=728 border=0 align="center" cellPadding=0 cellSpacing=0> 
  <TBODY> 
  <TR> 
      <TD>&nbsp; <div align="right"><span class="STYLE1"><strong>��վ���Բ���̨����</strong></span></div></TD>  
          
      <TD align=right width=400><a href="javascript:if(confirm('ȷ��Ҫ�˳���?')) window.location='logout.asp';" class="STYLE1">�˳����Բ���̨����</a>&nbsp;</TD>  
      </TR></TBODY></TABLE>  
	    
	    
	<%  
	strSourceFile = Server.MapPath("siva.xml")  
	Set objXML = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")  
	objXML.load(strSourceFile)  
	Set objRootsite = objXML.documentElement.selectSingleNode("guestbook")  
  
	'ÿҳ��ʾ*������  
	PageSize = cint(""&num&"")		  
	  
	'��ȡ�ӽڵ����ݣ���Ϊ�Ǵӽڵ�����0��ʼ��������ӽڵ���Ҫ��1��  
	AllNodesNum = objRootsite.childNodes.length - 1  
		  
	'�����ҳ��  
	PageNum = AllNodesNum\PageSize + 1   
	PageNo = cint(Request.querystring("PageNo"))  
	  
	'�����ÿһ�λ��ҳ����λ��ÿһҳ��ʾ���µ�����  
	if PageNo="" or PageNo="0" then  
		PageNo = 1  
	end if  
	  
	'�����ʼ�ڵ�  
	StarNodes = AllNodesNum - (PageNo - 1)*PageSize  
	  
	'��ý����ڵ�  
	EndNodes = StarNodes - PageSize + 1  
	  
	if EndNodes < 0 then  
		EndNodes = 0  
	end If  
	  
	'�ж���ʼ�ڵ����Ƿ񳬹��ܵĽڵ���  
	if StarNodes > AllNodesNum then  
		'�������������ڵ�Ҫ��ȥ(StarNodes-AllNodesNum)�Ĳ�ֵ�����±�ᳬ�����  
		EndNodes=EndNodes-(StarNodes-AllNodesNum)  
		StarNodes=AllNodesNum  
	end if  
	if EndNodes < 0 then  
		EndNodes=0  
	end if  
	while StarNodes >= EndNodes  
		id=objRootsite.childNodes.item(StarNodes).childNodes.item(0).text  
		name=objRootsite.childNodes.item(StarNodes).childNodes.item(1).text
		qq=objRootsite.childNodes.item(StarNodes).childNodes.item(2).text  
		email=objRootsite.childNodes.item(StarNodes).childNodes.item(3).text  
		sex=objRootsite.childNodes.item(StarNodes).childNodes.item(4).text  
		content=objRootsite.childNodes.item(StarNodes).childNodes.item(5).text  
		addtime=objRootsite.childNodes.item(StarNodes).childNodes.item(6).text	  
		reply=objRootsite.childNodes.item(starNodes).childNodes.item(7).text  
		hftime=objRootsite.childNodes.item(starNodes).childNodes.item(8).text  

	%>  
	    
	    
<TABLE width=728 border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" style="background-image: url('../images/left_top_bg01.gif'); background-repeat: repeat-x; border: 1px solid #dddddd">  
  <TBODY>  
  <TR>  
    <TD align=middle width=150><%=name%></TD>  
    <TD style="BACKGROUND-COLOR: #dddddd" align=middle width=1>��</TD>
    <TD height=20 style="border-bottom:#dddddd 1px solid;">&nbsp;<font color='#0066c9'>����ʱ�䣺</font><%=addtime%>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#0066c9'  style="FONT-FAMILY: georgia;font-size:11px;">Email��</font><a href="mailto:<%=email%>"><font style="FONT-FAMILY: georgia;font-size:11px;"><%=email%></font></a>&nbsp;QQ:<%=qq%>&nbsp;&nbsp;&nbsp;&nbsp;<a href="reply.asp?id=<%=id%>"><font color="#ff0000">�ظ�����</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="del.asp?id=<%=id%>"><font color="#ff0000">ɾ������</font></a>&nbsp;&nbsp;</TD></TR>
  <TR>
      <TD align=middle width=150 height=70><img border="0" src="images/<%=sex%>"></TD>
    <TD style="BACKGROUND-COLOR: #dddddd" align=middle width=1>��</TD>
    <TD>
      <TABLE cellSpacing=0 cellPadding=8 width="578" border=0>
          <!--DWLayoutTable-->
          <TBODY>
            <TR> 
              <TD width="562" height="66" valign="top"><%
			  
			  response.write replace(content,chr(13),"<br>")
			  if reply<>"" then
			  
			  response.write "<br><br><font color='#f60044'>����Ա�ظ���</font>"
			  response.write reply
			  
			  response.write "<p align=left><font color='#454545'>�ظ�ʱ�䣺"
			  
			  response.write hftime
			  
			  response.write "</font></p>"
			  
			  end if
			  
			  %>			  </TD>
            </TR>
          </TBODY>
        </TABLE></TD></TR></TBODY></TABLE>


<TABLE height=8 align="center" cellSpacing=0 cellPadding=0 width=728 border=0>
  <TBODY>
  <TR>
    <TD class="STYLE1"><script language="javascript" src="http://www.huaid.cn/blog/blogb1/blogbb.js"></script> </TD></TR></TBODY></TABLE>


<span class="STYLE1">
<% 
	StarNodes = StarNodes - 1
	wend 
	set objXML = nothing 
%>
</span>
<TABLE height=8 cellSpacing=0 cellPadding=0 width=728 border=0 align="center">
  <TBODY>
  <TR>
    <TD class="STYLE1"></TD></TR></TBODY></TABLE>
	

<TABLE width=728 border=0 align="center" cellPadding=0 cellSpacing=0>
  <!--DWLayoutTable-->
  <TBODY>
     <TR> 
        <TD width=526 height="22" valign="top" class="STYLE1">&nbsp;ҳ����<%=pageno%>/<%=pagenum%>&nbsp;   
        <%if pageno <> 1 then  
response.write "<a href='?pageno=1'>��ҳ</a>"  
response.write "&nbsp;<a href='?pageno="&pageno-1&"'>��һҳ</a>"  
else  
response.write "<FONT color=silver>��ҳ&nbsp;��һҳ</font>"  
end if                                                              
if pageno <> pagenum then  
response.write "&nbsp;<a href='?pageno="&pageno+1&"'>��һҳ</a>"  
response.write "&nbsp;<a href='?pageno="&pagenum&"'>βҳ</a>"  
else  
response.write "&nbsp;<FONT color=silver>��һҳ&nbsp;βҳ</font>"  
  
end if%>  
        &nbsp;&nbsp;&nbsp;����<%=AllNodesNum+1%>��(ÿҳ<%=PageSize%>��)</TD>  
      <TD width=202 align=right valign=top class="STYLE1" style="padding-right:4px;"><span class="STYLE1" style="padding-right:4px;">
        <input style="cursor:hand; BORDER-TOP: 2px solid #FFFFFF; FONT-SIZE: 14px; BACKGROUND: #00AED0; BORDER-LEFT: 2px solid #FFFFFF; WIDTH: 80px; HEIGHT: 22px" type="button" value="&gt;&gt;д������" name="btn3" onClick="javascrit:window.location='index.asp';">
        <a href="http://www.51.la/report/0_menu.asp?id=523746" target="_blank"></a></span></TD>  
    </TR>  
  </TBODY>  
</TABLE>  
  
	  
	  
	  
	  
	  
<span class="STYLE1">
<!--���Գ������-->  
</span>
<TABLE height=8 cellSpacing=0 cellPadding=0 width=728 border=0 align="center">  
  <TBODY>  
  <TR>  
    <TD class="STYLE1"></TD></TR></TBODY></TABLE>  
</BODY>
</HTML>  