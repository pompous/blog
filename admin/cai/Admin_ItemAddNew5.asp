<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim UrlTest,TsString,ToString,CsString,CoString
Dim DateType,DsString,DoString,UpDateTime
Dim AuthorType,AsString,AoString,AuthorStr
Dim CopyFromType,FsString,FoString,CopyFromStr
Dim KeyType,KsString,KoString,KeyStr
Dim NewsPaingType,NPsString,NPoString,NewsPaingStr,NewsPaingHtml
Dim NewsPaingNext,NewsPaingNextCode,ContentTemp
Dim NewsUrl,NewsCode
Dim Title,ConTent,Author,CopyFrom,Key
Dim UploadFiles,strInstallDir,strChannelDir

strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/"))
strChannelDir="Test"

ItemID=Trim(Request.Form("ItemID"))
UrlTest=Trim(Request.Form("UrlTest"))
TsString=Request.Form("TsString")
ToString=Request.Form("ToString")
CsString=Request.Form("CsString")
CoString=Request.Form("CoString")

DateType=Trim(Request.Form("DateType"))
DsString=Request.Form("DsString")
DoString=Request.Form("DoString")

AuthorType=Trim(Request.Form("AuthorType"))
AsString=Request.Form("AsString")
AoString=Request.Form("AoString")
AuthorStr=Trim(Request.Form("AuthorStr"))

CopyFromType=Trim(Request.Form("CopyFromType"))
FsString=Request.Form("FsString")
FoString=Request.Form("FoString")
CopyFromStr=Trim(Request.Form("CopyFromStr"))

KeyType=Trim(Request.Form("KeyType"))
KsString=Request.Form("KsString")
KoString=Request.Form("KoString")
KeyStr=Trim(Request.Form("KeyStr"))

NewsPaingType=Trim(Request.Form("NewsPaingType"))
NPsString=Request.Form("NPsString")
NPoString=Request.Form("NPoString")
NewsPaingStr=Trim(Request.Form("NewsPaingStr"))
NewsPaingHtml=Request.Form("NewsPaingHtml")


If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
If UrlTest="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，数据传递时发生错误</li>"
Else
   NewsUrl=UrlTest
End If
If TsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题开始标记不能为空</li>"
End If
If ToString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题结束标记不能为空</li>" 
End If
If CsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文开始标记不能为空</li>"
End If
If CoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文结束标记不能为空</li>" 
End If

If DateType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置时间类型</li>" 
Else
   DateType=Clng(DateType)
   If DateType=0 Then
   ElseIf DateType=1 Then
      If DsString="" or DoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将时间的开始/结束标记填写完整</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If AuthorType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置作者类型</li>" 
Else
   AuthorType=Clng(AuthorType)
   If AuthorType=0 Then
   ElseIf AuthorType=1 Then
      If AsString="" or AoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将作者的开始/结束标记填写完整</li>" 
      End If
   ElseIf AuthorType=2 Then
      If AuthorStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定作者</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

If CopyFromType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置来源类型</li>" 
Else
   CopyFromType=Clng(CopyFromType)
   If CopyFromType=0 Then
   ElseIf CopyFromType=1 Then
      If FsString="" or FoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将来源的开始/结束标记填写完整！</li>" 
      End If
   ElseIf CopyFromType=2 Then
      If CopyFromStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定来源</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

If KeyType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置关键字类型</li>" 
Else
   KeyType=Clng(KeyType)
   If KeyType=0 Then
   ElseIf KeyType=1 Then
      If KsString="" or KoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>关键字的开始/结束标记不能为空</li>" 
      End If
   ElseIf KeyType=2 Then
      If KeyStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定关键字</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If NewsPaingType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置文章分页类型</li>"
Else
   NewsPaingType=Clng(NewsPaingType)
   If NewsPaingType=0 Then
   ElseIf NewsPaingType=1 Then
      If NPsString="" or NPoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>分页开始/结束标记不能为空</li>" 
      End If
      If NewsPaingStr<>""  And  Len(NewsPaingStr)<15  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>分页绝对链接设置不正确(留空或者至少15个字符)</li>" 
      End  If            
   ElseIf NewsPaingType=2 Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>暂不支持手动设置分页类型</li>" 
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3

   RsItem("TsString")=TsString
   RsItem("ToString")=ToString
   RsItem("CsString")=CsString
   RsItem("CoString")=CoString

   RsItem("DateType")=DateType
   If DateType=1 Then
      RsItem("DsString")=DsString
      RsItem("DoString")=DoString
   End If

   RsItem("AuthorType")=AuthorType
   If AuthorType=1 Then
      RsItem("AsString")=AsString
      RsItem("AoString")=AoString
   ElseIf AuthorType=2 Then
      RsItem("AuthorStr")=AuthorStr
   End If

   RsItem("CopyFromType")=CopyFromType
   If CopyFromType=1 Then
      RsItem("FsString")=FsString
      RsItem("FoString")=FoString
   ElseIf CopyFromType=2 Then
      RsItem("CopyFromStr")=CopyFromStr
   End If

   RsItem("KeyType")=KeyType
   If KeyType=1 Then
      RsItem("KsString")=KsString
      RsItem("KoString")=KoString
   ElseIf KeyType=2 Then
      RsItem("KeyStr")=KeyStr
   End If

   RsItem("NewsPaingType")=NewsPaingType
   If NewsPaingType=1 Then
      RsItem("NPsString")=NPsString
      RsItem("NPoString")=NPoString
      If NewsPaingStr<>"" Then
         RsItem("NewsPaingStr")=NewsPaingStr
      End If
      RsItem("NewsPaingHtml")=NewsPaingHtml       
   End If
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If


If FoundErr<>True Then
   NewsCode=GetHttpPage(NewsUrl)
   If NewsCode<>"$False$" Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      Content=GetBody(NewsCode,CsString,CoString,False,False)
      If Title="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取标题的时候发生错误：" & NewsUrl & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
      End If

      If Content="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取正文的时候发生错误：" & NewsUrl & "</li>"
      Else
         '文章分页
         If NewsPaingType=1 Then
            NewsPaingNext=GetPaing(NewsCode,NPsString,NPoString,False,False)
            Do While NewsPaingNext<>"$False$"
               If NewsPaingStr="" or Isnull(NewsPaingStr)=True Then
                  NewsPaingNext=DefiniteUrl(NewsPaingNext,NewsUrl)
               Else
                  NewsPaingNext=Replace(NewsPaingStr,"{$ID}",NewsPaingNext)
               End If
               If NewsPaingNext="" or NewsPaingNext="$False$" Then Exit Do
               NewsPaingNextCode=GetHttpPage(NewsPaingNext)                  
               ContentTemp=GetBody(NewsPaingNextCode,CsString,CoString,False,False)
               If ContentTemp="$False$" Then
                  Exit Do
               Else
                  Content=Content & NewsPaingHtml & ContentTemp
                  NewsPaingNext=GetPaing(NewsPaingNextCode,NPsString,NPoString,False,False)
               End If
            Loop
         End If
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取源码时发生错误："& NewsUrl &"</li>"
   End If 
End If

If FoundErr<>True Then
      If DateType=0 then
         UpDateTime=Now()
      Else
         UpDateTime=GetBody(NewsCode,DsString,DoString,False,False)
         UpDateTime=FpHtmlEncode(UpDateTime)
         UpDateTime=Trim(Replace(UpDateTime,"&nbsp;"," "))
         If IsDate(UpDateTime)=True Then
            UpDateTime=CDate(UpDateTime)
         Else
            UpDateTime=Now()
         End If
      End If

      If AuthorType=1 Then
         Author=GetBody(NewsCode,AsString,AoString,False,False)
      ElseIf AuthorType=2 Then
         Author=AuthorStr
      End If
      If Author="$False$" Or Trim(Author)="" Then
         Author="佚名"
      Else
         Author=FpHtmlEnCode(Author)
      End If

      If CopyFromType=1 Then
         CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
      ElseIf CopyFromType=2 Then
         CopyFrom=CopyFromStr
      End If
      If CopyFrom="$False$" Or Trim(CopyFrom)="" Then
         CopyFrom="不详"
      Else
         CopyFrom=FpHtmlEnCode(CopyFrom)
      End If

      If KeyType=0 Then
         Key=Title
         Key=CreateKeyWord(Key,2)
      ElseIf KeyType=1 Then
         Key=GetBody(NewsCode,KsString,KoString,False,False)
         Key=FpHtmlEnCode(Key)
         Key=CreateKeyWord(Key,2)
      ElseIf KeyType=2 Then
         Key=KeyStr
         Key=FpHtmlEnCode(Key)
      End If
      If Key="$False$" Or Trim(Key)="" Then
         Key="南国都市"
      End If
End If

If FoundErr<>True Then
   Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,False,NewsUrl)
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End if
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
    <td height="30" class="b1_1"><a href="Admin_ItemAddNew.asp">添加项目</a> >> <a href="Admin_ItemModify.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="Admin_ItemModify2.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="Admin_ItemModify3.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="Admin_ItemModify4.asp?ItemID=<%=ItemID%>">正文设置</a> >> <font color=red>采样测试</font> >> 属性设置 >> 完成</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable" >
    <tr> 
      <td height="22" colspan="2" class="admintitle">添加新项目--采样测试</td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="b1_1"><%=Title%>　作者：<%=Author%>&nbsp;&nbsp;来源：<%=CopyFrom%>&nbsp;&nbsp;更新时间：<%=UpDateTime%></td>
    </tr>
    <tr>
      <td colspan="2" class="b1_1"><span lang="zh-cn"><%=Content%></span></td>
    </tr>
    <tr>
      <td colspan="2" class="b1_1"><b>关键字：<%=key%></b></td>
    </tr>
</table>

<form method="post" action="Admin_ItemAttribute.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td colspan="2" align="center">
        <input name="ItemID" type="hidden" value="<%=ItemID%>">
        <input name="button1" type="button" id="Cancel" value=" 上&nbsp;一&nbsp;步 " onClick="window.location.href='javascript:history.go(-1)'">
        &nbsp; 
        <input  type="submit" name="Submit" value="  下&nbsp;一&nbsp;步 "></td>
    </tr>
</table>
</form>       


				
		</td>
	</tr>
</table>


</body>

</html><%End Sub%>