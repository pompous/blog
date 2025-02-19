<html>
<!--#include file="conn.asp"-->
<head>

<style fprolloverstyle>
A:hover {
	color: #FFFFFF
}
</style>
<style type="text/css">
td{font-size:12px;}
.zi12 {
	font-size: 12px;
	color: #148AA8;
}
.style4 {font-size: 12px}
.tab {
	margin-bottom: 1px;
	padding-bottom: 1px;
	border-bottom-width: 1px;
	border-bottom-style: dashed;
	border-bottom-color: #333333;
	border-top-style: none;
	border-right-style: none;
	border-left-style: groove;
}
body {
	background-color: #999999;
}
.STYLE11 {color: #CCCCCC}
.STYLE14 {color: #333333}
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


function checkemail(qemail)

{

if(qemail.charAt(0)=="." ||qemail.charAt(0)=="@"||qemail.indexOf('@', 0) == -1 || qemail.indexOf('.', 0) == -1 || qemail.lastIndexOf("@")==qemail.value.length-1 || qemail.lastIndexOf(".")==qemail.value.length-1)
return true;

 }


function checkform() {




if(checkspace(form1.name.value)){
    alert ("请输入昵称!");
   form1.name.value="";
	form1.name.focus();
	return false;
}

if(checkspace(form1.qq.value)){
    alert ("别告诉我你没QQ号啊！呵呵");
   form1.qq.value="";
	form1.qq.focus();
	return false;
}

if(checkspace(form1.email.value)){
    alert ("留下EMAIL方便联系哦");
   form1.email.value="";
	form1.email.focus();
	return false;
}

if(checkspace(form1.content.value))

{
   alert ("请输入留言内容!");
   form1.content.value="";
	form1.content.focus();
	return false;

}


if(!checkspace(form1.email.value)){
  

if(checkemail(form1.email.value)){

    alert("Email地址格式不正确!");
	form1.email.value="";
    form1.email.focus();
	return false;
}
return true;
}

form1.subm.disabled=true;

}


</script>
<title><%=webname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link href="../css/default.css" rel="stylesheet" type="text/css">
<link href="../css/a.css" rel="stylesheet" type="text/css">
<link href="../css/img.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<!--#include file="../xwinc/top.asp"-->
  <td>
          <tr>
              <%       
	strSourceFile = Server.MapPath("siva.xml")       
	Set objXML = Server.CreateObject("Microsoft.FreeThreadedXMLDOM")       
	objXML.load(strSourceFile)       
	Set objRootsite = objXML.documentElement.selectSingleNode("guestbook")       
       
	'每页显示*条留言       
	PageSize = cint(""&num&"")		       
	       
	'获取子节点数据（因为是从节点数从0开始的所最大子节点数要减1）       
	AllNodesNum = objRootsite.childNodes.length - 1       
		       
	'算出总页数       
	PageNum = AllNodesNum\PageSize + 1        
	PageNo = cint(Request.querystring("PageNo"))       
	       
	'如果是每一次获得页面则定位到每一页显示最新的留言       
	if PageNo="" or PageNo="0" then       
		PageNo = 1       
	end if       
	       
	'获得起始节点       
	StarNodes = AllNodesNum - (PageNo - 1)*PageSize       
	       
	'获得结束节点       
	EndNodes = StarNodes - PageSize + 1       
	       
	if EndNodes < 0 then       
		EndNodes = 0       
	end If       
	       
	'判断起始节点数是否超过总的节点数       
	if StarNodes > AllNodesNum then       
		'如果超过则结束节点要减去(StarNodes-AllNodesNum)的差值否则下标会超界出错       
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
                <TABLE width="100%" border=0 cellPadding=8 cellSpacing=0>
                  <TBODY>
                    <TR>
                      <TD valign="top" style="font-size: 10pt"><b><font color="#333333"><%=name%></font> 发表于：<%=addtime%> </b></TD>
                    </TR>
         
                      <TD style="border-bottom:1px solid #333333; font-size: 10pt; border-left-width:1px; border-right-width:1px; border-top-width:1px">         <TR>
                          <%   
			     
			  response.write replace(content,chr(13),"<br>")   
			  if reply<>"" then   
			     
			  response.write "<br><br><font color='#f60044'>站长回复：</font>"   
			  response.write reply     
			     
			  end if   
			     
			  %>                      </TD>
                    </TR>
                  </TBODY>
                </TABLE>
              <%    
	StarNodes = StarNodes - 1   
	wend    
	set objXML = nothing    
%>
                <p align="right">&nbsp;页数：<FONT color=red><%=pageno%></FONT>/<%=pagenum%>&nbsp;
                    <%if pageno <> 1 then        
response.write "<a href='?pageno=1'>首页</a>"        
response.write "&nbsp;<a href='?pageno="&pageno-1&"'>上一页</a>"        
else        
response.write "<FONT color=silver>首页&nbsp;上一页</font>"        
end if                                                                    
if pageno <> pagenum then        
response.write "&nbsp;<a href='?pageno="&pageno+1&"'>下一页</a>"        
response.write "&nbsp;<a href='?pageno="&pagenum&"'>尾页</a>"        
else        
response.write "&nbsp;<FONT color=silver>下一页&nbsp;尾页</font>"        
        
end if%>
                  &nbsp;&nbsp;&nbsp;共有<FONT color=red><%=AllNodesNum+1%></FONT>篇(每页<%=PageSize%>篇)</td>
          </tr>
          <tr> </tr>
                </table>		  
		  <table border="0" cellpadding="0" cellspacing="0" width="100%">
            <form action="add.asp" method="post" name="form1" onSubmit="return checkform()">
              <tr>
                <td align="left" colspan="3" style="color: #999999; font-family: 宋体; font-size: 14pt; font-weight: bold">&nbsp;>
                  <input type="radio" name="sex" value="male.gif" checked>
                  <span class="STYLE14">发表新留言</span></td>
                <td align="left" style="color: #999999; font-family: 宋体; font-size: 14pt; font-weight: bold">&nbsp;</td>
                <td align="left" style="color: #999999; font-family: 宋体; font-size: 14pt; font-weight: bold">&nbsp;</td>
                <td align="left" style="color: #999999; font-family: 宋体; font-size: 14pt; font-weight: bold">&nbsp;</td>
              </tr>
              <tr>
                <td align="right">昵称：</td>
                <td><input type="text" name="name" maxlength="8" class="tab" size="10">
                    <br></td>
                <td><p align="right">QQ：</td>
                <td><input type="text" name="qq" maxlength="9" class="tab" size="10"></td>
                <td><p align="right">邮箱：</td>
                <td><input type="text" name="email" maxlength="30" class="tab" size="16"></td>
              </tr>
              <tr>
                <td align="right">内容：</td>
                <td height="80" colspan="5"><textarea name="content" cols="70" rows="6"></textarea></td>
              </tr>
              <tr>
                <td>　</td>
                <td align="center"><p align="left">&nbsp;
                  <input type="submit" name="subm" value=" 提 交 " style="color: #999999; background-color: #E0F0F8; font-family:宋体; font-size:10pt; font-weight:bold">
                </td>
                <td align="center"><input name="reset" type="reset" style="color: #999999; background-color: #E0F0F8; font-family:宋体; font-size:10pt; font-weight:bold" value=" 重 填 "></td>
              </tr>
            </form>
      </table></td>

</body>
</html>