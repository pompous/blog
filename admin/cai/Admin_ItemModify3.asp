<%@language=vbscript codepage=936%>
<%
option explicit
response.buffer=true
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim SqlItem,RsItem,ItemID,FoundErr,ErrMsg,Action
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim HsString,HoString,HttpUrlType,HttpUrlStr
Dim ListUrl,ListCode,ListPaingNext
ItemID=Trim(Request("ItemID"))
Action=Trim(Request("Action"))

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
Else
   ItemID=Clng(ItemID)
End If
If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr<>True Then
   Call GetTest()
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>

<%Sub Main%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS 系统管理</title>
<link href="../Images/admin_css.css" rel="stylesheet" type="text/css" />


</head>

<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">　</td>
		<td valign="top"><br>	



<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1">项目编辑 >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>"><font color=red>链接设置</font></a> >> <a href="Admin_ItemModify4.asp?ItemID=<%=ItemID%>">正文设置</a> >> <a href="Admin_ItemModify5.asp?ItemID=<%=ItemID%>">采样测试</a> >> <a href="Admin_ItemAttribute.asp?ItemID=<%=ItemID%>">属性设置</a> >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="22" colspan="2" class="admintitle">编辑项目--列表截取测试</td>
  </tr>
  <tr>
    <td height="22" colspan="2"><textarea name="Content" id="Content" style="width:100%;height:300px;"><%=ListCode%></textarea>
    </td>
  </tr>
</table>
<%If ListPaingNext<>"" And ListPaingNext<>"$False$" Then%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable">
  <tr>
    <td height="22" colspan="2" ><%Response.Write "<br>下一页列表：<a  href='" & ListPaingNext  &  "' target=_blank><font  color=red>"  &  ListPaingNext  &  "</font></a>"%>
    </td>
  </tr>
</table>
<%End If%>
<form method="post" action="Admin_ItemModify4.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
    <tr> 
      <td colspan="2" class="admintitle">编辑项目--链接设置></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>链接开始标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="HsString" cols="49" rows="7"><%=HsString%></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>链接结束标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="HoString" cols="49" rows="7"><%=HoString%></textarea></td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong> 链接特殊处理：</strong></td>
      <td width="75%" class="b1_1">
		<input name="HttpUrlType" type="radio" class="noborder" onClick="HttpUrl1.style.display='none'" value="0" <%If HttpUrlType=0 Then Response.Write "checked"%>> 
		自动处理&nbsp;
		<input name="HttpUrlType" type="radio" class="noborder" onClick="HttpUrl1.style.display=''" value="1" <%If HttpUrlType=1 Then Response.Write "checked"%>> 
		重新定位
      </td>
    </tr>
	<tr id="HttpUrl1" style="display:'<%If HttpUrlType=0 Then Response.Write "none"%>'">
      <td width="20%" class="b1_1"><strong>绝对链接字符：</strong></td>
      <td width="75%" class="b1_1">
	  <input name="HttpUrlStr" type="text" size="49" maxlength="200" value="<%=HttpUrlStr%>"></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="b1_1">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='Admin_ItemModify2.asp?ItemID=<%=ItemID%>'">
<input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步"></td>
    </tr>
</table>
</form>      
			

				
		</td>
	</tr>
</table>


</body>

</html>
<%End Sub%>
<%
Sub SaveEdit
   ListStr=Trim(Request.Form("ListStr"))
   LsString=Request.Form("LsString")
   LoString=Request.Form("LoString")
   ListPaingType=Request.Form("ListPaingType")
   LPsString=Request.Form("LPsString")
   LPoString=Request.Form("LPoString")
   ListPaingStr1=Trim(Request.Form("ListPaingStr1"))
   ListPaingStr2=Trim(Request.Form("ListPaingStr2"))
   ListPaingID1=Request.Form("ListPaingID1")
   ListPaingID2=Request.Form("ListPaingID2")
   ListPaingStr3=Request.Form("ListPaingStr3")

If ItemID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
If LsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空</li>"
End If
If LoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空</li>" 
End If
If ListPaingType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择列表索引分页类型</li>" 
Else
   ListPaingType=Clng(ListPaingType)
   Select Case ListPaingType
   Case 0,1
            If ListStr="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>列表索引页不能为空</li>"
            Else
               ListStr=Trim(ListStr)
            End If
      If  ListPaingType=1  Then
            If LPsString="" or LPoString="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>索引分页开始/结束标记不能为空</li>" 
            End If
            If ListPaingStr1<>"" and Len(ListPaingStr1)<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>索引分页重定向设置不正确(至少15个字符)</li>" 
            End If
      End  If
   Case 2
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成字符不能为空</li>"
      End If
      If isNumeric(ListPaingID1)=False or isNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围只能是数字</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成范围设置不正确</li>"
         End If
      End If
   Case 3
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引分页不能为空，请手动添加</li>"
      Else
         ListPaingStr3=Replace(ListPaingStr3,CHR(13),"|") 
      End If
   Case Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择列表索引分页类型</li>" 
   End Select
End if

If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3

   RsItem("LsString")=LsString
   RsItem("LoString")=LoString
   RsItem("ListPaingType")=ListPaingType
   Select Case ListPaingType
   Case 0,1
      RsItem("ListStr")=ListStr
      If ListPaingType=1  Then
         RsItem("LPsString")=LPsString
         RsItem("LPoString")=LPoString
         RsItem("ListPaingStr1")=ListPaingStr1
      End If
   Case 2
      RsItem("ListPaingStr2")=ListPaingStr2
      RsItem("ListPaingID1")=ListPaingID1
      RsItem("ListPaingID2")=ListPaingID2
   Case 3
      RsItem("ListPaingStr3")=ListPaingStr3
   End Select
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
End Sub


'==================================================
'过程名：GetTest
'作  用：测试
'参  数：无
'==================================================
Sub GetTest
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
   Else
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
      ListStr=RsItem("ListStr")
      LsString=RsItem("LsString")
      LoString=RsItem("LoString")
      ListPaingType=RsItem("ListPaingType")
      LPsString=RsItem("LPsString")
      LPoString=RsItem("LPoString")
      ListPaingStr1=RsItem("ListPaingStr1")
      ListPaingStr2=RsItem("ListPaingStr2")
      ListPaingID1=RsItem("ListPaingID1")
      ListPaingID2=RsItem("ListPaingID2")
      ListPaingStr3=RsItem("ListPaingStr3")
      HsString=RsItem("HsString")
      HoString=RsItem("HoString")
      HttpUrlType=RsItem("HttpUrlType")
      HttpUrlStr=RsItem("HttpUrlStr")
   End If
   RsItem.Close
   Set RsItem=Nothing     
   If LsString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空！</li>"
   End If
   If LoString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空！</li>"
   End If
   If ListPaingType=0 Or ListPaingType=1 Then
      If ListStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引页不能为空！</li>"
      End If
      If ListPaingType=1 Then    
         If LPsString="" Or LPoString="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>索引分页开始/结束标记不能为空！</li>"
         End If
         If ListPaingStr1<>"" And Len(ListPaingStr1)<15 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>索引分页绝对链接设置不正确(请留空或者字符>15个)！</li>"
         End If
      End If      
   ElseIf ListPaingType=2 Then
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成原字符串不能为空！</li>"
      End If
      If IsNumeric(ListPaingID1)=False or IsNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围不正确！无</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成的范围不正确！</li>"
         End If
      End If 
   ElseIf ListPaingType=3 Then
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页不能为空！</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请选择索引分页类型</li>"
   End If
 
   If LoginType=1 Then
      If LoginUrl="" or LoginPostUrl="" or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将登录信息填写完整</li>"
      End If
   End If

   If FoundErr<>True Then
      Select Case ListPaingType
      Case 0,1
         ListUrl=ListStr
      Case 2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      Case 3
         If Instr(ListPaingStr3,"|")> 0 Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
            ListUrl=ListPaingStr3
         End If
      End Select

      If LoginType=1 then
         LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
         LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
         If Instr(LoginResult,LoginFalse)>0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>登录网站时发生错误，请确认登录信息的正确性！</li>"
         End If
      End If
   End If
   If FoundErr<>True Then
      ListCode=GetHttpPage(ListUrl)
      If ListCode<>"$False$" Then
         If ListPaingType=1 Then
            ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
            If ListPaingNext<>"$False$" then
               If ListPaingStr1<>"" Then
                  ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
               Else
                  ListPaingNext=DefiniteUrl(ListPaingNext,ListUrl)
               End If
            End If
         End If
         ListCode=GetBody(ListCode,LsString,LoString,False,False)
         If ListCode="$False$" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
         End If
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
      End If
   End If 
End Sub
%>