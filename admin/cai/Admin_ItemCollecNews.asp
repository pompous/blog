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
Dim ItemNum,NewsNum,PaingNum,NewsSuccesNum,NewsFalseNum,NewsNumAll
Dim Rs,Sql,RsItem,SqlItem,FoundErr,ErrMsg,ItemEnd,NewsEnd

'项目变量
Dim ItemID,ItemName,ClassID,strChannelDir,SpecialID
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString
Dim CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NpoString,NewsPaingStr,NewsPaingHtml
Dim PaginationType,MaxCharPerPage,ReadLevel,Stars,ReadPoint,Hits,UpDateType,UpDateTime,IncludePicYn,DefaultPicYn,OnTop,Elite,Hot
Dim SkinID,TemplateID,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,CollecNewsNum,Passed,SaveFiles,CollecOrder,LinkUrlYn,InputerType,Inputer,EditorType,Editor,ShowCommentLink

'过滤变量
Dim Arr_Filters,FilterStr,Filteri

'采集相关的变量
Dim ContentTemp,NewsPaingNext,NewsPaingNextCode,Arr_i,NewsUrl,NewsCode

'文章保存变量
Dim ArticleID,Title,Content,Author,CopyFrom,Key,IncludePic,UploadFiles,DefaultPicUrl

'其它变量
Dim Arr_Item,Arr_News,CollecTest,Content_View

'历史记录
Dim Arr_Histrolys,His_Title,His_CollecDate,His_Result,His_Repeat,His_i 

'执行时间变量
Dim StartTime,OverTime

'图片统计
Dim Arr_Images,ImagesNum,ImagesNumAll

Dim strInstallDir,CacheTemp
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/"))

CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

ItemNum=Clng(Trim(Request("ItemNum")))
NewsNum=Clng(Trim(Request("NewsNum")))
NewsSuccesNum=Clng(Trim(Request("NewsSuccesNum")))
NewsFalseNum=Clng(Trim(Request("NewsFalseNum")))
ImagesNumAll=Clng(Trim(Request("ImagesNumAll")))
NewsPaingNext=Trim(Request("NewsPaingNext"))
ArticleID=Trim(Request("ArticleID"))
NewsNumAll=Trim(Request("NewsNumAll"))
If ArticleID="" Then
   ArticleID=0
Else
   ArticleID=Clng(ArticleID)
End If
If NewsNumAll="" Then
   NewsNumAll=0
Else
   NewsNumAll=Clng(NewsNumAll)
End If
FoundErr=False
ItemEnd=False
NewsEnd=False

Call SetCache
If ItemEnd<>True Then
   If (ItemNum-1)>Ubound(Arr_Item,2) then
      ItemEnd=True
   Else
      Call SetItems()
   End If
   If ItemEnd<>True Then
      If NewsNum=1 Then
         Call SetNews()
      Else
         Call GetNews()
      End if
      If NewsEnd<>True Then
         If (NewsNum-1)>Ubound(Arr_News,2) Then
            NewsEnd=True
         Else
            NewsUrl=Arr_News(0,NewsNum-1)
         End If
      End If
   End If
End If

If ItemEnd=True Then
   ErrMsg="<br>采集任务全部完成"
   ErrMsg=ErrMsg & "<br>全部文章：" & NewsNumAll & " 条，成功采集： "  &  NewsSuccesNum  &  "  条文章，失败： "    &  NewsFalseNum  &  "  条，图片： " & ImagesNumAll & "  张"
   Call DelCache()
Else
   If NewsEnd=True Then
      ItemNum=ItemNum+1
      NewsNum=1
      Call SetHistroly()
      ErrMsg="<br>" & ItemName & "  项目所有列表采集完成，正在整理数据请稍后..."
      ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecNews.asp?ItemNum=" & ItemNum & "&NewsNum=" & NewsNum & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&NewsNumAll=" & NewsNumAll & """>"
   End If
End If

Call TopItem()
Response.Flush
If ItemEnd=True Or NewsEnd=True Then
   Call WriteSucced(ErrMsg)
Else
   FoundErr=False
   ErrMsg=""
   Call TopItem2()
   Response.Flush
   Call StartCollection()
   Call FootItem2()
End  If
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
Sub StartCollection()
   '变量初始化
   UploadFiles=""
   DefaultPicUrl=""
   IncludePic=0
   ImagesNum=0
   NewsCode=""
   FoundErr=False
   ErrMsg=""
   His_Repeat=False
   Title=""
   PaingNum=1
   '……………………………………………… 
   If Response.IsClientConnected Then 
      Response.Flush 
   Else 
      Response.End 
   End If
   '……………………………………………… 

   If CollecTest=False Then
      His_Repeat=CheckRepeat(NewsUrl)
   Else
      His_Repeat=False
   End If
   If His_Repeat=True Then
      FoundErr=True
   End If

   If FoundErr<>True Then
      NewsCode=GetHttpPage(NewsUrl)
      If NewsCode="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br>在获取：" & NewsUrl & "文章源码时发生错误！"
         Title="分析源码错误"
      End If
   End If

   If FoundErr<>True Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      If Title="$False$" or Title="" then
         FoundErr=True
         ErrMsg=ErrMsg & "<br>在分析：" & NewsUrl & "的文章标题时发生错误"
         Title="<br>标题分析错误" 
      End If
      If FoundErr<>True Then
         Content=GetBody(NewsCode,CsString,CoString,False,False)
         If Content="$False$" or Content="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br>在分析：" & NewsUrl & "的文章正文时发生错误"
            Title=Title & "<br>正文分析错误" 
         End If
      End If
   End If
   If FoundErr<>True Then
      '文章分页
      If NewsPaingType=1 Then
         NewsPaingNext=GetPaing(NewsCode,NPsString,NPoString,False,False)
         NewsPaingNext=FpHtmlEnCode(NewsPaingNext)
         Do While NewsPaingNext<>"$False$" And NewsPaingNext<>""
            If NewsPaingStr="" or IsNull(NewsPaingStr)=True Then
               NewsPaingNext=DefiniteUrl(NewsPaingNext,NewsUrl)
            Else
               NewsPaingNext=Replace(NewsPaingStr,"{$ID}",NewsPaingNext)
            End If
            If NewsPaingNext="" or NewsPaingNext="$False$" Then
               Exit Do
            End If
            NewsPaingNextCode=GetHttpPage(NewsPaingNext)                  
            ContentTemp=GetBody(NewsPaingNextCode,CsString,CoString,False,False)
            If ContentTemp="$False$" Then
               Exit Do
            Else
               PaingNum=PaingNum+1
               Content=Content & NewsPaingHtml & ContentTemp
               NewsPaingNext=GetPaing(NewsPaingNextCode,NPsString,NPoString,False,False)
               NewsPaingNext=FpHtmlEnCode(NewsPaingNext)
            End If
         Loop
      End If
      '过滤
      Call Filters()
      Title=FpHtmlEnCode(Title)
      Call FilterScript()
      Content=Ubbcode(Content)
   End If

   '分开写（太长了照顾不过来）
   If FoundErr<>True Then
      '时间
      If UpDateType=0 Then
         UpDateTime=Now()
      ElseIf UpDateType=1 Then
         If DateType=0 then
            UpDateTime=Now()
         Else
            UpDateTime=GetBody(NewsCode,DsString,DoString,False,False)
            UpDateTime=Lcase(FpHtmlEncode(UpDateTime))
            UpDateTime=Trim(Replace(UpDateTime,"&nbsp;"," "))
            If IsDate(UpDateTime)=True Then
               UpDateTime=CDate(UpDateTime)
            Else
               UpDateTime=Now()
            End If
         End If
      ElseIf UpDateType=2 Then  
      Else
         UpDateTime=Now()
      End If
                
      '作者
      If AuthorType=1 Then
         Author=GetBody(NewsCode,AsString,AoString,False,False)
      ElseIf AuthorType=2 Then
         Author=AuthorStr
      Else
         Author="佚名"
      End If
      Author=FpHtmlEncode(Author)
      If Author="" or Author="$False$" then
         Author="佚名"
      Else
         If Len(Author)>255 then
            Author=Left(Author,255)
         End If
      End If
         
      '来源
      If CopyFromType=1 Then
         CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
      ElseIf CopyFromType=2 Then
         CopyFrom=CopyFromStr
      Else
         CopyFrom="不详"
      End If
      CopyFrom=FpHtmlEncode(CopyFrom)
      If CopyFrom="" or CopyFrom="$False$" Then
         CopyFrom="不详"
      Else
         If Len(CopyFrom)>255 Then 
            CopyFrom=Left(CopyFrom,255)
         End If
      End If

      '关键字
      If KeyType=0 Then
         Key=Title
         Key=CreateKeyWord(Key,2)
      ElseIf KeyType=1 Then
         Key=GetBody(NewsCode,KsString,KoString,False,False)
         Key=FpHtmlEncode(Key)
         Key=CreateKeyWord(Key)
      ElseIf KeyType=2 Then
         Key=KeyStr
         Key=FpHtmlEncode(Key)
         If Len(Key)>253 Then
            Key="|" & Left(Key,253) & "|"
         Else
            Key="|" & Key & "|"
         End If
      End If
      If Key="" or Key="$False$" Then
         Key="|南国都市|文章|"
      End If
   End If

   If FoundErr<>True Then 
      '转换图片相对地址为绝对地址/保存
      If CollecTest=False And SaveFiles=True then
         Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,True,NewsUrl)              
      Else
         Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,False,NewsUrl)
      End If
      '转换swf文件地址
      Content=ReplaceSwfFile(Content,NewsUrl)

      '图片统计、文章图片属性设置
      If UploadFiles<>"" Then
         If Instr(UploadFiles,"|")>0 Then
            Arr_Images=Split(UploadFiles,"|") 
            ImagesNum=Ubound(Arr_Images)+1
            DefaultPicUrl=Arr_Images(0)
         Else
            ImagesNum=1
            DefaultPicUrl=UploadFiles
         End If
         If DefaultPicYn=False then
            DefaultPicUrl=""
         End If
         If IncludePicYn=True Then
            IncludePic=-1
         Else
            IncludePic=0
         End If
         If SaveFiles<>True Then
            UploadFiles=""
         End If
      Else
         ImagesNum=0
         DefaultPicUrl=""
         IncludePic=0         
      End If
      ImagesNumAll=ImagesNumAll+ImagesNum
   End If

   If FoundErr<>True Then
      If CollecTest=False Then
         Call SaveArticle
         SqlItem="INSERT INTO Histroly(ItemID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ClassID & "','" & SpecialID & "','" & ArticleID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',True)"
         ConnItem.Execute(SqlItem)
         Content=Replace(Content,"[InstallDir_ChannelDir]",strInstallDir & strChannelDir & "/")
      End If
      NewsSuccesNum=NewsSuccesNum+1
      ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
      ErrMsg=ErrMsg & "文章标题："
      ErrMsg=ErrMsg & "<font color=red>" & Title & "</font><br>"
      ErrMsg=ErrMsg & "更新时间：" & UpDateTime & "<br>"
      ErrMsg=ErrMsg & "文章作者：" & Author & "<br>"
      ErrMsg=ErrMsg & "文章来源：" & CopyFrom & "<br>"
      ErrMsg=ErrMsg & "采集页面：<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
      ErrMsg=ErrMsg & "其它信息：分页--" & PaingNum & " 页，图片--" & ImagesNum & " 张<br>"
      ErrMsg=ErrMsg & "正文预览："
      If Content_View=True Then
         ErrMsg=ErrMsg & "<br>" & Content
      Else
         ErrMsg=ErrMsg & "您没有启用正文预览功能"
      End If
      ErrMsg=ErrMsg & "<br><br>关 键 字：" & Key & ""
   Else
      NewsFalseNum=NewsFalseNum+1
      If His_Repeat=True Then
         ErrMsg=ErrMsg & "<div class='admintable' style='text-align:left;'>No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
         ErrMsg=ErrMsg & "目标文章：<font color=red>"
         If His_Result=True Then
            ErrMsg=ErrMsg & His_Title
         Else
            ErrMsg=ErrMsg & NewsUrl
         End If
         ErrMsg=ErrMsg & "</font> 的记录已存在，不予采集。<br>"
         ErrMsg=ErrMsg & "采集时间：" & His_CollecDate & "<br>"
         ErrMsg=ErrMsg & "文章来源：<a href='" & NewsUrl & "' target=_blank>"&NewsUrl&"</a><br>"
         ErrMsg=ErrMsg & "采集结果："
         If His_Result=False Then
            ErrMsg=ErrMsg & "失败"
            ErrMsg=ErrMsg & "<br>失败原因：" & Title
         Else
            ErrMsg=ErrMsg & "成功"
         End If            
         ErrMsg=ErrMsg & "<br>提示信息：如想再次采集，请先将该文章的历史记录<font color=red>删除</font><br></div>"
      End If
      If CollecTest=False And His_Repeat=False Then
         SqlItem="INSERT INTO Histroly(ItemID,ClassID,SpecialID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ClassID & "','" & SpecialID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',False)"
         ConnItem.Execute(SqlItem)
      End If
   End If

   ErrMsg=ErrMsg & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class='admintable'>"
   ErrMsg=ErrMsg & "<tr>"
   ErrMsg=ErrMsg & "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
   ErrMsg=ErrMsg & "数据整理中，3秒后继续......3秒后如果还没反应请点击 <a href='Admin_ItemCollecNews.asp?ItemNum=" & ItemNum & "&NewsNum=" & NewsNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ArticleID=" & ArticleID & "&NewsNumAll=" & NewsNumAll & "'><font color=red>这里</font></a> 继续<br>"
   ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecNews.asp?ItemNum=" & ItemNum & "&NewsNum=" & NewsNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ArticleID=" & ArticleID & "&NewsNumAll=" & NewsNumAll & """>"
   ErrMsg=ErrMsg & "</td></tr>"
   ErrMsg=ErrMsg & "</table>"

   Call ShowMsg(ErrMsg)
   Response.Flush()'刷新
End Sub



'==================================================
'过程名：SetCache
'作  用：获取变量
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
      ErrMsg="<br><li>参数错误，请重新运行！</li>"
   End If

   '过滤信息
   myCache.name=CacheTemp & "filters"
   If myCache.valid then 
      Arr_Filters=myCache.value
   End If

   '历史记录
   myCache.name=CacheTemp & "histrolys"
   If myCache.valid then 
      Arr_Histrolys=myCache.value
   End If

   '其它信息
   myCache.name=CacheTemp & "collectest"
   If myCache.valid then 
      CollecTest=myCache.value
   Else
      CollecTest=False
   End If

   myCache.name=CacheTemp & "contentview"
   If myCache.valid then 
      Content_View=myCache.value
   Else
      Content_View=False
   End If

   Set myCache=Nothing
End Sub

'==================================================
'过程名：GetNews
'作  用：获取变量
'参  数：无
'==================================================
Sub GetNews()
   Dim myCache
   Set myCache=new clsCache

   '文章信息
   myCache.name=CacheTemp & "news"
   If myCache.valid then 
      Arr_News=myCache.value
   End If
   If IsArray(Arr_News)=False Then
      NewsEnd=True
   End If
   Set myCache=Nothing
End Sub

Sub DelCache()
   Dim myCache
   Set myCache=new clsCache
   myCache.name=CacheTemp & "items"
   Call myCache.clean()

   myCache.name=CacheTemp & "filters"
   Call myCache.clean()

   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()

   myCache.name=CacheTemp & "collectest"
   Call myCache.clean()

   myCache.name=CacheTemp & "contentview"
   Call myCache.clean()

   myCache.name=CacheTemp & "news"
   Call myCache.clean()

   Set myCache=Nothing
End Sub

'==================================================
'过程名：SetItems
'作  用：获取项目信息
'参  数：无
'==================================================
Sub SetItems()
      Dim ItemNumTemp
      ItemNumTemp=ItemNum-1
      ItemID=Arr_Item(0,ItemNumTemp)
      ItemName=Arr_Item(1,ItemNumTemp)
      ClassID=Arr_Item(2,ItemNumTemp)'栏目ID
      strChannelDir=Arr_Item(3,ItemNumTemp)'栏目目录
      ClassID=Arr_Item(4,ItemNumTemp)            '栏目
      SpecialID=Arr_Item(5,ItemNumTemp)        '专题
      TsString=Arr_Item(30,ItemNumTemp)          '标题
      ToString=Arr_Item(31,ItemNumTemp)
      CsString=Arr_Item(32,ItemNumTemp)          '正文
      CoString=Arr_Item(33,ItemNumTemp)
      DateType=Arr_Item(34,ItemNumTemp)      '作者
      DsString=Arr_Item(35,ItemNumTemp)          
      DoString=Arr_Item(36,ItemNumTemp)
      AuthorType=Arr_Item(37,ItemNumTemp)      '作者
      AsString=Arr_Item(38,ItemNumTemp)          
      AoString=Arr_Item(39,ItemNumTemp)
      AuthorStr=Arr_Item(40,ItemNumTemp)
      CopyFromType=Arr_Item(41,ItemNumTemp)  '来源
      FsString=Arr_Item(42,ItemNumTemp)          
      FoString=Arr_Item(43,ItemNumTemp)
      CopyFromStr=Arr_Item(44,ItemNumTemp)
      KeyType=Arr_Item(45,ItemNumTemp)            '关键词
      KsString=Arr_Item(46,ItemNumTemp)          
      KoString=Arr_Item(47,ItemNumTemp)
      KeyStr=Arr_Item(48,ItemNumTemp)
      NewsPaingType=Arr_Item(49,ItemNumTemp)            '关键词
      NPsString=Arr_Item(50,ItemNumTemp)          
      NPoString=Arr_Item(51,ItemNumTemp)
      NewsPaingStr=Arr_Item(52,ItemNumTemp)
      NewsPaingHtml=Arr_Item(53,ItemNumTemp)
      PaginationType=Arr_Item(55,ItemNumTemp)
      MaxCharPerPage=Arr_Item(56,ItemNumTemp)
      ReadLevel=Arr_Item(57,ItemNumTemp)
      Stars=Arr_Item(58,ItemNumTemp)
      ReadPoint=Arr_Item(59,ItemNumTemp)
      Hits=Arr_Item(60,ItemNumTemp)
      UpDateType=Arr_Item(61,ItemNumTemp)
      UpDateTime=Arr_Item(62,ItemNumTemp)
      IncludePicYn=Arr_Item(63,ItemNumTemp)
      DefaultPicYn=Arr_Item(64,ItemNumTemp)
      OnTop=Arr_Item(65,ItemNumTemp)
      Elite=Arr_Item(66,ItemNumTemp)
      Hot=Arr_Item(67,ItemNumTemp)
      SkinID=Arr_Item(68,ItemNumTemp)
      TemplateID=Arr_Item(69,ItemNumTemp)
      Script_Iframe=Arr_Item(70,ItemNumTemp)
      Script_Object=Arr_Item(71,ItemNumTemp)
      Script_Script=Arr_Item(72,ItemNumTemp)
      Script_Div=Arr_Item(73,ItemNumTemp)
      Script_Class=Arr_Item(74,ItemNumTemp)
      Script_Span=Arr_Item(75,ItemNumTemp)
      Script_Img=Arr_Item(76,ItemNumTemp)
      Script_Font=Arr_Item(77,ItemNumTemp)
      Script_A=Arr_Item(78,ItemNumTemp)
      Script_Html=Arr_Item(79,ItemNumTemp)
      CollecNewsNum=Arr_Item(81,ItemNumTemp)
      Passed=Arr_Item(82,ItemNumTemp)
      SaveFiles=Arr_Item(83,ItemNumTemp)
      CollecOrder=Arr_Item(84,ItemNumTemp)
      LinkUrlYn=Arr_Item(85,ItemNumTemp)
      InputerType=Arr_Item(86,ItemNumTemp)
      Inputer=Arr_Item(87,ItemNumTemp)
      EditorType=Arr_Item(88,ItemNumTemp)
      Editor=Arr_Item(89,ItemNumTemp)
      ShowCommentLink=Arr_Item(90,ItemNumTemp)
      If InputerType=1 Then
         Inputer=FpHtmlEnCode(Inputer)
      Else
         Inputer=session("AdminName")
      End If
      If EditorType=1 Then
         Editor=FpHtmlEnCode(Editor)
      Else
         Editor=session("AdminName")
      End If
      If IsObjInstalled("Scripting.FileSystemObject")=False or strChannelDir="" Then
         SaveFiles=False
      End if
End Sub

Sub SetNews()
   SqlItem ="select NewsUrl from NewsList where ItemID=" & ItemID
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_News=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   Dim myCache
   Set myCache=new clsCache
   myCache.name=CacheTemp & "news"
   Call myCache.clean()
   If IsArray(Arr_News)=True Then
      myCache.add Arr_News,Dateadd("n",1000,now)
   Else
      NewsEnd=True
   End If
   Set myCache=Nothing
End Sub

Sub SetHistroly()
   Dim myCache
   Set myCache=new clsCache
   '历史记录
   SqlItem ="select NewsUrl,Title,CollecDate,Result from Histroly"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Histrolys=RsItem.GetRows()
      myCache.name=CacheTemp & "histrolys"
      Call myCache.clean()
      myCache.add Arr_Histrolys,Dateadd("n",1000,now)
   End If
   RsItem.Close
   Set RsItem=Nothing
   Set myCache=Nothing
End Sub
'==================================================
'过程名：SaveArticle
'作  用：保存文章
'参  数：无
'==================================================
Sub SaveArticle
   'If ArticleID=0 Then
     ' set rs=server.createobject("adodb.recordset")
     ' sql="select top 1 ArticleID from LZ8_Article order by ArticleID desc" 
     ' rs.open sql,conn,1,1
     ' If rs.eof and rs.bof then
         'ArticleID=1
      'Else
        ' ArticleID=rs("ArticleID")+1
     ' End If
      'rs.close
      'set rs=nothing
   'Else
      'ArticleID=ArticleID+1
   'End If
   set rs=server.createobject("adodb.recordset")
   sql="select top 1 * from xiaowei_Article" 
   rs.open sql,conn,1,3
   rs.addnew
   'rs("ArticleID")=ArticleID
   rs("ClassID")=ClassID
   rs("Title")=Title
   rs("Keyword")=Left(Key,10)
   rs("Hits")=Hits
   rs("Author")=Author
   rs("CopyFrom")=CopyFrom
   rs("Content")=Content
   rs("yn")=0
   rs("IsTop")=OnTop
   rs("IsHot")=Hot
   rs("DateAndTime")=UpDateTime
   rs.update
   rs.close
   set rs=nothing
End Sub

'==================================================
'过程名：Filters
'作  用：过滤
'==================================================
Sub Filters()
If IsNull(Arr_Filters)=True or IsArray(Arr_Filters)=False Then
   Exit Sub
End if

   For Filteri=0 to Ubound(Arr_Filters,2)
      FilterStr="$False$"
      If Arr_Filters(1,Filteri)=ItemID Or Arr_Filters(10,Filteri)=True Then
         If Arr_Filters(3,Filteri)=1 Then'标题过滤
            If Arr_Filters(4,Filteri)=1 Then
               Title=Replace(Title,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Title=Replace(Title,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         ElseIf Arr_Filters(3,Filteri)=2 Then'正文过滤
            If Arr_Filters(4,Filteri)=1 Then
               Content=Replace(Content,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Content=Replace(Content,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         End If
      End If
   Next
End Sub

'==================================================
'过程名：FilterScript
'作  用：脚本过滤
'==================================================

Sub  FilterScript()
   If Script_Iframe=True Then
      Content=ScriptHtml(Content,"Iframe",1)
   End If
   If Script_Object=True Then
      Content=ScriptHtml(Content,"Object",2)
   End If
   If Script_Script=True Then
      Content=ScriptHtml(Content,"Script",2)
   End If
   If Script_Div=True Then
      Content=ScriptHtml(Content,"Div",3)
   End If
   If Script_Span=True Then
      Content=ScriptHtml(Content,"Span",3)
   End If
   If Script_Img=True Then
      Content=ScriptHtml(Content,"Img",3)
   End If
   If Script_Font=True Then
      Content=ScriptHtml(Content,"Font",3)
   End If
   If Script_A=True Then
      Content=ScriptHtml(Content,"A",3)
   End If
   If Script_Html=True Then
      Content=noHtml(Content)
   End If
End  Sub

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
Sub TopItem2%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="admintable">
    <tr>
      <td height="22" colspan="2" aling="left">本次运行：<%=Ubound(Arr_Item,2)+1%> 个项目，正在采集第 <font color=red><%= ItemNum%></font> 个项目 <font color=red><%=ItemName%></font> 的第 <font color=red><%=NewsNum%></font> 条，该项目文章 <%=Ubound(Arr_News,2)+1%> 条，全部文章 <%=NewsNumAll%> 条。
      <br>采集统计：成功采集--<%=NewsSuccesNum%>  条，失败--<%=NewsFalseNum%>  条，图片--<%=ImagesNumAll%> 张。<a href="Admin_ItemStart.asp"><font color=red>停止采集</font></a>
      </td>
    </tr>
</table>
<%StartTime=Timer()%>
<%End Sub%>

<%
Sub FootItem()%>
<!--#include file="../Admin_Copy.asp"-->       
				
		</td>
	</tr>
</table>


</body>

</html>
<%End Sub%>

<%
Sub FootItem2()
   Dim strTemp
   OverTime=Timer()
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class='admintable'>"       
   strTemp=strTemp & "<tr>"          
   strTemp=strTemp & "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
   strTemp=strTemp & "执行时间：" & CStr(FormatNumber((OverTime-StartTime)*1000,2)) & " 毫秒"
   strTemp=strTemp & "</td></tr><br>"
   strTemp=strTemp & "</table>"
   Response.write strTemp
End Sub

Sub ShowMsg(Msg)
   Dim strTemp
   strTemp= Msg
   Response.Write StrTemp     
End Sub

Function CheckRepeat(strUrl)
   CheckRepeat=False
   If IsArray(Arr_Histrolys)=True then
      For His_i=0 to Ubound(Arr_Histrolys,2)
         If Arr_Histrolys(0,His_i)=strUrl Then
            CheckRepeat=True
            His_Title=Arr_Histrolys(1,His_i)
            His_CollecDate=Arr_Histrolys(2,His_i)
            His_Result=Arr_Histrolys(3,His_i)
            Exit For
         End If
      Next
   End If
End Function
%>