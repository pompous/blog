<%@language=vbscript codepage=936 %>
<%
option explicit
Response.Buffer = True 
Server.ScriptTimeOut=999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/clsCache.asp"-->
<%
Dim ItemNum,ListNum,ListSuccesNum,ListFalseNum,NewsNumAll
Dim Rs,Sql,RsItem,SqlItem,FoundErr,ErrMsg,ItemEnd,ListEnd

'项目变量
Dim ItemID,ItemName,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr,CollecListNum,CollecNewsNum

'采集相关的变量
Dim Arr_i,NewsUrl

'其它变量
Dim LoginData,LoginResult
Dim Arr_Item,CacheTemp,CollecOrder,OrderTemp

'执行时间变量
Dim StartTime,OverTime

'列表
Dim ListUrl,ListCode,NewsArrayCode,NewsArray,ListArray,ListPaingNext,ListPaingTemp

CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

ItemNum=Clng(Trim(Request("ItemNum")))
ListNum=Clng(Trim(Request("ListNum")))
ListSuccesNum=Clng(Trim(Request("ListSuccesNum")))
ListFalseNum=Clng(Trim(Request("ListFalseNum")))
NewsNumAll=Clng(Trim(Request("NewsNumAll")))
ListPaingNext=Trim(Request("ListPaingNext"))

FoundErr=False
ItemEnd=False
ListEnd=False
CollecListNum=0
CollecNewsNum=0

Call SetCache()

If ItemEnd<>True Then
   If (ItemNum-1)>Ubound(Arr_Item,2) then
      ItemEnd=True
   Else
      Call SetItems()
   End If
End If

If ItemEnd<>True Then
   If ListPaingType=0 Then
      If ListNum=1 Then
         ListUrl=ListStr
      Else
         ListEnd=True
      End If
   ElseIf ListPaingType=1 Then
      If ListNum=1 Then
         ListUrl=ListStr
      Else
         If ListPaingNext="" or ListPaingNext="$False$" Then
            ListEnd=True
         Else
            ListPaingNext=Replace(ListPaingNext,"{$ID}","&")
            ListUrl=ListPaingNext
         End If
      End If
   ElseIf ListPaingType=2 Then
      If ListPaingID1>ListPaingID2 then
         If (ListPaingID1-ListNum+1)<ListPaingID2 or (ListPaingID1-ListNum+1)<0 Then
            Listend=True
         Else
            ListUrl=Replace(ListPaingStr2,"{$ID}",Cstr(ListpaingID1-ListNum+1))
         End if
      Else
         If (ListPaingID1+ListNum-1)>ListPaingID2 Then
            ListEnd=True
         Else
            ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1+ListNum-1))
         End If
      End If      
   ElseIf ListPaingType=3  Then
      ListArray=Split(ListPaingStr3,"|")
      If (ListNum-1)>Ubound(ListArray) Then
         ListEnd=True
      Else
         ListUrl=ListArray(ListNum-1)
      End If    
   End If
   If ListNum>CollecListNum And CollecListNum<>0 Then
      ListEnd=True
   End if
End If

If ItemEnd=True Then
   ErrMsg="<br>列表分析完成"
   ErrMsg=ErrMsg & "<br>成功分析： "  &  ListSuccesNum  &  "  页列表,失败： "    &  ListFalseNum  &  "  页,文章：" & NewsNumAll & "  条"
   ErrMsg=ErrMsg& "<br>正在整理数据，稍后进行文章的采集..."
   ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecNews.asp?ItemNum=1&NewsNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNumAll=" & NewsNumAll & """>"
Else
   If ListEnd=True Then
      ItemNum=ItemNum+1
      ListNum=1
      ErrMsg="<br>" & ItemName & "  项目所有列表分析完成，正在整理数据请稍后..."
      ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecSteady.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&ListSuccesNum=" & ListSuccesNum & "&ListFalseNum=" & ListFalseNum & "&NewsNumAll=" & NewsNumAll & """>"
   End If
End If

Call TopItem()
If ItemEnd<>True And ListEnd<>True Then
   FoundErr=False
   ErrMsg=""
   Call StartCollection()
End  If

Call WriteSucced(ErrMsg)
Call FootItem()
Response.Flush()
'关闭数据库链接
Call CloseConn()
Call CloseConnItem()
%>

<%
'==================================================
'过程名：StartCollection
'作  用：开始采集
'参  数：无
'==================================================
Sub StartCollection

'第一次采集时登录
If LoginType=1 And ListNum=1 then
   LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
   LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
   If Instr(LoginResult,LoginFalse)>0 Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在登录网站时发生错误，请确保登录信息的正确性！</li>"
   End If
End If


If FoundErr<>True then
   ListCode=GetHttpPage(ListUrl)
   Call GetListPaing()
   If ListCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取列表：" & ListUrl & "网页源码时发生错误！</li>"
   Else
      ListCode=GetBody(ListCode,LsString,LoString,False,False)
      If ListCode="$False$" Or ListCode="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取：" & ListUrl & "的文章列表时发生错误！</li>"
      End If
   End If
End If

If FoundErr<>True Then
   NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
   If NewsArrayCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在分析：" & ListUrl & "文章列表时发生错误！</li>"
   Else
      NewsArray=Split(NewsArrayCode,"$Array$")
      For Arr_i=0 to Ubound(NewsArray)
         If HttpUrlType=1 Then
            NewsArray(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(Arr_i)))
         Else
            NewsArray(Arr_i)=Trim(DefiniteUrl(NewsArray(Arr_i),ListUrl))           
         End If
         NewsArray(Arr_i)=CheckUrl(NewsArray(Arr_i))
      Next
      If CollecOrder=True Then
         For Arr_i=0 to Fix(Ubound(NewsArray)/2)
            OrderTemp=NewsArray(Arr_i)
            NewsArray(Arr_i)=NewsArray(Ubound(NewsArray)-Arr_i)
            NewsArray(Ubound(NewsArray)-Arr_i)=OrderTemp
         Next
      End If
   End If
End If

If FoundErr<>True Then
   ErrMsg=ErrMsg & "<br>本次运行 " & Ubound(Arr_Item,2)+1 & " 个项目"
   ErrMsg=ErrMsg & "<br>从第 " & ItemNum  & " 个项目 " & ItemName & " 的第 "  & ListNum & " 页列表分析出 " & Ubound(NewsArray) & " 条文章"
   If CollecNewsNum<>0 Then
      ErrMsg=ErrMsg & "，限制 " & CollecNewsNum & " 条。"
      If (CollecNewsNum-1)>Ubound(NewsArray) Then
         CollecNewsNum=Ubound(NewsArray)+1
      Else
         '保持不变CollecNewsNum
      End If
   Else
      CollecNewsNum=Ubound(NewsArray)+1
   End If
   ListSuccesNum=ListSuccesNum+1
   NewsNumAll=NewsNumAll+CollecNewsNum
   Call SaveNewsList()
Else
   ListFalseNum=ListFalseNum+1
End If
ErrMsg=ErrMsg & "<br>" & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecSteady.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&ListSuccesNum=" & ListSuccesNum & "&ListFalseNum=" & ListFalseNum & "&NewsNumAll=" & NewsNumAll & "&ListPaingNext=" & ListPaingNext  & """>"

End Sub


'==================================================
'过程名：SetCache
'作  用：存取缓存
'参  数：无
'==================================================
Sub SetCache()
   Dim myCache
   Set myCache=new clsCache

   '项目信息
   myCache.name=CacheTemp & "items"
   If myCache.valid then 
      Arr_Item=myCache.value
   Else
      ItemEnd=True
   End If
   Set myCache=Nothing
End Sub

Sub SetItems() 
      Dim ItemNumTemp
      ItemNumTemp=ItemNum-1
      ItemID=Arr_Item(0,ItemNumTemp)
      ItemName=Arr_Item(1,ItemNumTemp)
      LoginType=Arr_Item(9,ItemNumTemp)
      LoginUrl=Arr_Item(10,ItemNumTemp)          '登录
      LoginPostUrl=Arr_Item(11,ItemNumTemp)
      LoginUser=Arr_Item(12,ItemNumTemp)
      LoginPass=Arr_Item(13,ItemNumTemp)
      LoginFalse=Arr_Item(14,ItemNumTemp)
      ListStr=Arr_Item(15,ItemNumTemp)            '列表地址
      LsString=Arr_Item(16,ItemNumTemp)          '列表
      LoString=Arr_Item(17,ItemNumTemp)
      ListPaingType=Arr_Item(18,ItemNumTemp)
      LPsString=Arr_Item(19,ItemNumTemp)          
      LPoString=Arr_Item(20,ItemNumTemp)
      ListPaingStr1=Arr_Item(21,ItemNumTemp)
      ListPaingStr2=Arr_Item(22,ItemNumTemp)
      ListPaingID1=Arr_Item(23,ItemNumTemp)
      ListPaingID2=Arr_Item(24,ItemNumTemp)
      ListPaingStr3=Arr_Item(25,ItemNumTemp)
      HsString=Arr_Item(26,ItemNumTemp)  
      HoString=Arr_Item(27,ItemNumTemp)
      HttpUrlType=Arr_Item(28,ItemNumTemp)
      HttpUrlStr=Arr_Item(29,ItemNumTemp)
      CollecListNum=Arr_Item(80,ItemNumTemp)
      CollecNewsNum=Arr_Item(81,ItemNumTemp)
      CollecOrder=Arr_Item(84,ItemNumTemp)
End Sub

'==================================================
'过程名：GetListPaing
'作  用：获取列表下一页
'参  数：无
'==================================================
Sub GetListPaing()
   If ListPaingType=1 Then
      ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
      ListPaingNext=FpHtmlEnCode(ListPaingNext)
      If ListPaingNext<>"$False$" And ListPaingNext<>"" Then
         If ListPaingStr1<>""  Then  
            ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
         Else
            ListPaingNext=DefiniteUrl(ListPaingNext,ListUrl)
         End If
         ListPaingNext=Replace(ListPaingNext,"&","{$ID}")
      End If
   Else
      ListPaingNext="$False$"
   End If
End Sub

'==================================================
'过程名：SaveNewsList
'作  用：保存文章
'参  数：无
'==================================================
Sub SaveNewsList
   set rs=server.createobject("adodb.recordset")
   sql="select top 1 * from NewsList" 
   rs.open sql,connItem,1,3
   For Arr_i=1 To CollecNewsNum
      rs.addnew
      rs("ItemID")=ItemID
      rs("NewsUrl")=NewsArray(Arr_i-1)
      rs.update
   Next
   rs.close
   set rs=nothing
End Sub

'==================================================
'过程名：TopItem
'作  用：显示导航信息
'参  数：无
'==================================================
Sub TopItem()%>

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
			
<%End Sub%>

<%
'==================================================
'过程名：FootItem
'作  用：显示底部版权等信息
'参  数：无
'==================================================
Sub FootItem()%>    
				
		</td>
	</tr>
</table>


</body>

</html>
<%End Sub%>