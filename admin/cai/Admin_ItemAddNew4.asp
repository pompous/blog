<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim ListStr,LsString,LoString
Dim  ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim  ListUrl,ListCode,NewsArrayCode,NewsArray,UrlTest,NewsCode
dim Testi
ItemID=Trim(Request("ItemID"))
HsString=Request.Form("HsString")
HoString=Request.Form("HoString")
HttpUrlType=Trim(Request.Form("HttpUrlType"))
HttpUrlStr=Trim(Request.Form("HttpUrlStr"))


If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
If HsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>链接开始标记不能为空</li>"
End If
If HoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>链接结束标记不能为空</li>" 
End If
If HttpUrlType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择链接处理类型</li>"
Else
   HttpUrlType=Clng(HttpUrlType)
   If HttpUrlType=1 Then
      If HttpUrlStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>绝对链接字符设置不能为空</li>"
      Else
            If  Len(HttpUrlStr)<15  Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>绝对链接字符设置不正确(15个字符以上)</li>"
            End  If      
      End If
   End If
End If


If FoundErr<>True Then
   SqlItem="Select ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3

   RsItem("HsString")=HsString
   RsItem("HoString")=HoString
   RsItem("HttpUrlType")=HttpUrlType
   If HttpUrlType=1 Then
      RsItem("HttpUrlStr")=HttpUrlStr
   End If
   ListStr=RsItem("ListStr")
   LsString=RsItem("LsString")
   LoString=RsItem("LoString")
   ListPaingType=RsItem("ListPaingType")
   LPsString=RsItem("LPsString")
   ListPaingStr1=RsItem("ListPaingStr1")
   ListPaingStr2=RsItem("ListPaingStr2")
   ListPaingID1=RsItem("ListPaingID1")
   ListPaingID2=RsItem("ListPaingID2")
   ListPaingStr3=RsItem("ListPaingStr3")

   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
   
   Select  Case  ListPaingType
   Case  0,1
         ListUrl=ListStr
   Case  2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case  3
         If  Instr(ListPaingStr3,"|")>0  Then
         ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
         ListUrl=ListPaingStr3
      End  If
   End  Select

   ListCode=GetHttpPage(ListUrl)
   If ListCode<>"$False$" Then
      ListCode=GetBody(ListCode,LsString,LoString,False,False)
      If ListCode<>"$False$" Then 
         NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
         If NewsArrayCode<>"$False$" Then
               NewsArray=Split(NewsArrayCode,"$Array$")
               For Testi=0 To Ubound(NewsArray)
                  If HttpUrlType=1 Then
                     NewsArray(Testi)=Replace(HttpUrlStr,"{$ID}",NewsArray(Testi))
                  Else
                     NewsArray(Testi)=DefiniteUrl(NewsArray(Testi),ListUrl)
                  End If
               Next
               UrlTest=NewsArray(0)
               NewsCode=GetHttpPage(UrlTest)
         Else
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>在获取链接列表时出错。</li>"
         End If   
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
   End If
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
<%Sub Main()%>

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
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">添加项目</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>">链接设置</a> >> <font color=red>正文设置</font> >> 采样测试 >> 属性设置 >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable" >        
    <tr> 
      <td height="22" colspan="2" class="admintitle">添加新项目--列表新闻链接测试</td>
    </tr>
    <tr>
      <td height="22" colspan="2" class="b1_1">以下是分析后所得到的文章绝对链接地址，请查看是否正确。<br>
        <%
For Testi=0 To Ubound(NewsArray)
   Response.Write "<a href='" & NewsArray(Testi) & "' target=_blank>" & NewsArray(Testi) & "</a><br>"
Next
%>
        <br>
下一步将抽取第一条文章进行测试，在填写以下标记时尽量不要使用第一条文章。</td>
    </tr>
</table>
<form method="post" action="Admin_ItemAddNew5.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable" >
    <tr> 
      <td height="22" colspan="2" class="admintitle">添加新项目--正文设置      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>标题开始标记：</strong>
        <p>　</p><p>　</p>
      <strong>标题结束标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="TsString" cols="49" rows="7"></textarea><br>
      <textarea name="ToString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><strong>正文开始标记：</strong>
        <p>　</p><p>　</p>
      <strong>正文结束标记：</strong></td>
      <td width="75%" class="b1_1">
      <textarea name="CsString" cols="49" rows="7"></textarea><br>
      <textarea name="CoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>时<span lang="en-us">&nbsp;</span>间<span lang="en-us">&nbsp;</span>设<span lang="en-us">&nbsp;</span>置：</b></td>
      <td width="75%" class="b1_1">
      	<input name="DateType" type="radio" class="noborder" onClick="Date1.style.display='none'" value="0" checked>
      	不作设置&nbsp;
		<input name="DateType" type="radio" class="noborder" onClick="Date1.style.display=''" value="1">
		设置标签&nbsp;
    </tr>
    <tr id="Date1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>时间开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>时间结束标记：</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="DsString" cols="49" rows="7"></textarea><br>
      <textarea name="DoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>作 者<span lang="en-us">&nbsp;</span>设<span lang="en-us">&nbsp;</span>置：</b></td>
      <td width="75%" class="b1_1">
      	<input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='none';Author2.style.display='none'" value="0" checked>
      	不作设置&nbsp;
		<input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='';Author2.style.display='none'" value="1">
		设置标签&nbsp;
	  <input name="AuthorType" type="radio" class="noborder" onClick="Author1.style.display='none';Author2.style.display=''" value="2">
	  指定作者</td>
    </tr>
    <tr id="Author1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>作者开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>作者结束标记：</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="AsString" cols="49" rows="7"></textarea><br>
      <textarea name="AoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="Author2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>请指定作者：</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="AuthorStr" type="text" id="AuthorStr" value="">
      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>来 源&nbsp;设 置：</b></td>
      <td width="75%" class="b1_1">
      	<input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'" value="0" checked>
      	不作设置&nbsp;
		<input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='';CopyFrom2.style.display='none'" value="1">
		设置标签&nbsp;
	  <input name="CopyFromType" type="radio" class="noborder" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display=''" value="2">
	  指定来源</td>
    </tr>
    <tr id="CopyFrom1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>来源开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>来源结束标记：</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="FsString" cols="49" rows="7"></textarea><br>
      <textarea name="FoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="CopyFrom2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>请指定来源：</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="CopyFromStr" type="text" id="CopyFromStr" value="">
      </td>
    </tr>
    <tr> 
      <td width="20%" class="b1_1" ><b>关键字词设置：</b></td>
      <td width="75%" class="b1_1">
      	<input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='none';Key2.style.display='none'" value="0" checked>
      	标题生成&nbsp;
		<input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='';Key2.style.display='none'" value="1">
		标签生成&nbsp;
	  <input name="KeyType" type="radio" class="noborder" onClick="Key1.style.display='none';Key2.style.display=''" value="2">
	  自定义关键字</td>
    </tr>
    <tr id="Key1" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>关键词开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>关键词结束标记：</font></strong></td>
      <td width="75%" class="b1_1">
      <textarea name="KsString" cols="49" rows="7"></textarea><br>
      <textarea name="KoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="Key2" style="display:none"> 
      <td width="20%" class="b1_1" ><strong><font color=blue>请指定关键：</font></strong></td>
      <td width="75%" class="b1_1">
      <input name="KeyStr" type="text" id="KeyStr" value="">
      <span class="note">格式：关键字之间用<font color=red>|</font>分隔，如：文章|IT</span>
      </td>
    </tr>
    <tr>
      <td width="20%" class="b1_1"><strong>正文分页设置：</strong></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='none';NewsPaing12.style.display='none';NewsPaing13.style.display='none';NewsPaing2.style.display='none'" value="0" checked>
		不作设置&nbsp;
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='';NewsPaing12.style.display='';NewsPaing13.style.display='';NewsPaing2.style.display='none'" value="1">
		设置标签&nbsp;
		<input name="NewsPaingType" type="radio" class="noborder" onClick="NewsPaing1.style.display='none';NewsPaing12.style.display='none';NewsPaing13.style.display='none';NewsPaing2.style.display=''" value="2">
		智能设置
      </td>
    </tr>
	<tr id="NewsPaing1" style="display:none">
      <td width="20%" class="b1_1"><strong><font color=blue>下页开始标记：</font></strong>
        <p>　</p><p>　</p>
      <strong><font color=blue>下页结束标记：</font></strong></td>
      <td width="75%" class="b1_1">
		<textarea name="NPsString" cols="49" rows="7"></textarea><br>
	  <textarea name="NPoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr id="NewsPaing12" style="display:none"> 
      <td width="20%" class="b1_1"><b><font color="#0000FF">分页绝对链接：</font></b></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingStr" type="text" size="58"><br>
		格式：http://www.pcook.com.cn/Article_Show.asp?ID=1&Page={$ID}
    </tr>
    <tr id="NewsPaing13" style="display:'none'"> 
      <td width="20%" class="b1_1"><b><font color="#0000FF">分页链接字符：</font></b></td>
      <td width="75%" class="b1_1">
	  <input name="NewsPaingHtml" type="text" size="58" value=""></td>
    </tr>
    <tr id="NewsPaing2" style="display:none"> 
      <td width="20%" class="b1_1"><strong><font color=blue>智&nbsp; 能&nbsp; 设&nbsp; 置：</font></strong></td>
      <td width="75%" class="b1_1">
		<input name="NewsPaingStr2" type="text" value="预留功能" size="51">
      </td>
    </tr>

    <tr> 
      <td colspan="2" align="center" class="b1_1"><br>
          <input name="ItemID" type="hidden" value="<%=ItemID%>"> 
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='javascript:history.go(-1)'" >
        &nbsp;&nbsp;&nbsp;&nbsp; 
      <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步"></td>
        <input type="hidden" name="UrlTest" id="UrlTest" value="<%=UrlTest%>">
    </tr>
</table>
</form>     
       

				
		</td>
	</tr>
</table>


</body>

</html><%End Sub%>