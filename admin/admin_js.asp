<!--#include file="../xwInc/conn.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>vpCMS ϵͳ����</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<script src="js/admin.js"></script>
<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">��</td>
		<td valign="top"><br>	
			
<script>
function GenerateJs()
{
	var s;
	s="";
	s=s.substring(0,s.lastIndexOf("/")+1);
	s+="http://<%=SiteUrl%><%=SitePath%>xwinc/js.asp?topType=";
	s+=document.all.topType.value;
	s+="&classNO=";
	s+=document.all.classNO.value;
	s+="&num=";
	s+=document.all.topNum.value;
	s+="&maxlen=";
	s+=document.all.maxLength.value;
	s+="&showdate=";
	s+=document.all.showDate.checked ? "11" : "0";
	s+="&showhits=";
	s+=document.all.showHits.checked ? "1" : "0";
	s+="&showClass=";
	s+=document.all.showClass.checked ? "1" : "0";
	document.all.jstext.value="<script src=\""+s+"\"><\/script>";
}

function CopyJs()
{
	document.all.jstext.focus();
	document.all.jstext.select();
	document.execCommand("copy");
}
</script>
<table border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
<form name="form1" method="post" action="">
    <tr> 
      <td height="30" colspan="2" class="admintitle">�ⲿ��������</td>
    </tr>
    <tr> 
      <td width="15%" height="25"  valign="middle" bgcolor="#f1f3f5" class="tdleft">�������ͣ�	  </td>
	  <td bgcolor="#f1f3f5" class="tdleft">
	  <select id="topType">
			<option value="new">�������</option>
			<option value="hot">��������</option>
			<option value="IsHot">�Ƽ�����</option>
	  </select>	  </td>
    </tr>
	<tr> 
      <td width="15%" height="25"  valign="middle" bgcolor="#f1f3f5" class="tdleft">���÷��ࣺ	  </td>
	  <td bgcolor="#f1f3f5" class="tdleft">
	  <select name="classNO" id="classNO">
      <%call Admin_ShowClass_Option()%>
    </select>	  </td>
    </tr>
	<tr> 
      <td width="15%" height="25"  valign="middle" bgcolor="#f1f3f5" class="tdleft">����������	  </td>
	  <td bgcolor="#f1f3f5" class="tdleft"><input type="text" id="topNum" value=10 size=3 maxlength=3></td>
    </tr>
	<tr> 
      <td width="15%" height="25"  valign="middle" bgcolor="#f1f3f5" class="tdleft">���ⳤ�ȣ�	  </td>
	  <td bgcolor="#f1f3f5" class="tdleft"><input type="text" id="maxLength" value=20 size=3 maxlength=3></td>
    </tr>
	<tr> 
      <td width="15%" height="25"  valign="middle" bgcolor="#f1f3f5" class="tdleft">��ʾ���ݣ�	  </td>
	  <td bgcolor="#f1f3f5" class="tdleft">
	    <input type="checkbox" class="noborder" id="showClass" >
	    ��ʾ�������&nbsp; &nbsp;
		<input type="checkbox" class="noborder" id="showDate" >��ʾ����&nbsp; &nbsp;
		<input type="checkbox" class="noborder" id="showHits" >��ʾ�����&nbsp;&nbsp;	  </td>
    </tr>
	<tr> 
      <td width="15%" height="25" colspan="2"  valign="middle" bgcolor="#f1f3f5" class="tdleft">
	  <input type="button" value="���ɴ���" onClick="GenerateJs();"> &nbsp; &nbsp; 
		<input type="button" value="��������" onClick="CopyJs();">	  </td>
    </tr>
	<tr> 
      <td height="25" colspan="2"  valign="middle" bgcolor="#f1f3f5" class="tdleft">
		<br>
		<input type="text" id="jstext" value="" size="95%" />
		<br><br>
		ע�������ɵĽű�ճ����html�ļ��ڼ�������ʾָ��������	  </td>
    </tr>
</form>
</table>
<%
sub Admin_ShowClass_Option()
   Dim Sqlp,Rsp,TempStr
   Sqlp ="select * from xiaowei_Class Where TopID = 0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">���޷���</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">������ӷ���</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
         Response.Write(">|-" & Rsp("ClassName") & "")
		 
		    Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """" & Rspp("ID") & """" & "")
         		Response.Write(">��|-" & Rspp("ClassName") & "")
				Response.Write("</option>" ) 
			Rspp.Movenext   
      		Loop
			
         Response.Write("</option>" ) 
      Rsp.Movenext   
      Loop   
   End if
end sub 
%>
				
		</td>
	</tr>
</table>


</body>

</html>