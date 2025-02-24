<%@language=vbscript codepage=936 %>
'********************************************************
%>
<%
option explicit
response.buffer=true
Server.ScriptTimeOut=999
Response.Expires = -1
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/clsCache.asp"-->
<%
'传递变量：ItemNum--项目
'          ListNum--列表
'          NewsSuccesNum--成功采集的文章数量
'          NewsFalseNum--失败采集的文章数量
'          ImagesNum----图片数目
'          ListPaingNext--列表分页
Dim ItemNum,ListNum,NewsNum,PaingNum,NewsSuccesNum,NewsFalseNum
Dim Rs,Sql,RsItem,SqlItem,FoundErr,ErrMsg,ItemEnd,ListEnd,NewsEnd,PaingEnd

'项目变量
Dim ItemID,ItemName,ClassID,strChannelDir,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString
Dim CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NpoString,NewsPaingStr,NewsPaingHtml
Dim ItemCollecDate,PaginationType,MaxCharPerPage,ReadLevel,Stars,ReadPoint,Hits,UpDateType,UpDateTime,IncludePicYn,DefaultPicYn,OnTop,Elite,Hot
Dim SkinID,TemplateID,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,CollecListNum,CollecNewsNum,Passed,SaveFiles,CollecOrder,LinkUrlYn,InputerType,Inputer,EditorType,Editor,ShowCommentLink

'过滤变量
Dim Arr_Filters,i_Filter,FilterStr,SqlF,RsF,Filteri

'采集相关的变量
Dim ContentTemp,NewsPaingNext,NewsPaingNextCode,Arr_i,NewsUrl,NewsCode

'文章保存变量
Dim ArticleID,Title,Content,Author,CopyFrom,Key,IncludePic,UploadFiles,DefaultPicUrl

'其它变量
Dim LoginData,LoginResult,OrderTemp
Dim Arr_Item,Arr_Other,CollecTest,Content_View,CacheTemp,CollecNewsAll

'历史记录
Dim Arr_Histrolys,His_HistrolyID,His_Title,His_CollecDate,His_Result,His_Repeat,His_i 

'执行时间变量
Dim StartTime,OverTime

'图片统计
Dim Arr_Images,ImagesNum,ImagesNumAll

'列表
Dim ListUrl,ListCode,ListUrlArray,NewsArrayCode,NewsArray,ListArray,ListPaingNext,ListPaingTemp

Dim strInstallDir
'获得动易安装文件夹
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/"))

CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

CollecNewsNum=0
ArticleID=0
ItemNum=Clng(Trim(Request("ItemNum")))
ListNum=Clng(Trim(Request("ListNum")))
NewsSuccesNum=Clng(Trim(Request("NewsSuccesNum")))
NewsFalseNum=Clng(Trim(Request("NewsFalseNum")))
ImagesNumAll=Trim(Request("ImagesNumAll"))
ListPaingNext=Trim(Request("ListPaingNext"))
If ImagesNumAll="" Then
   ImagesNumAll=0
Else
   ImagesNumAll=Clng(ImagesNumAll)
End If
FoundErr=False
ItemEnd=False
ListEnd=False

Call SetCache

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
         If ListNum>CollecListNum And CollecListNum<>0 Then
            ListEnd=True
         Else
            If ListPaingNext="" or ListPaingNext="$False$" Then
               ListEnd=True
            Else
               ListPaingNext=Replace(ListPaingNext,"{$ID}","&")
               ListUrl=ListPaingNext
            End If
         End If
      End If
   ElseIf ListPaingType=2 Then
      If (ListPaingID1+ListNum-1)>ListPaingID2 Then
         ListEnd=True
      Else
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1+ListNum-1))
      End If
   ElseIf ListPaingType=3  Then
      ListArray=Split(ListPaingStr3,"|")
      If (ListNum-1)>Ubound(ListArray) Then
         ListEnd=True
      Else
         ListUrl=ListArray(ListNum-1)
      End If    
   End If
End If

If ItemEnd=True Then
   ErrMsg="<br>采集任务全部完成"
   ErrMsg=ErrMsg & "<br>成功采集： "  &  NewsSuccesNum  &  "  条,失败： "    &  NewsFalseNum  &  "  条,图片：" & ImagesNumAll & "  张"
   Call DelCache()
Else
   If ListEnd=True Then
      ItemNum=ItemNum+1
      ListNum=1
      ErrMsg="<br>" & ItemName & "  项目所有列表采集完成，正在存取缓存请稍后..."
      ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & """>"
   End If
End If

Call TopItem()
If ItemEnd=True Or ListEnd=True Then
   Call WriteSucced(ErrMsg)
Else
   FoundErr=False
   ErrMsg=""
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
   Call TopItem2()
   CollecNewsAll=0
   For Arr_i=0 to Ubound(NewsArray)
      'If CollecTest=True  And  Arr_i=10 Then
         'Exit For
      'End If
      If CollecNewsAll>=CollecNewsNum And CollecNewsNum<>0 Then
         Exit For
      End If
      CollecNewsAll=CollecNewsAll+1
      '变量初始化
      UploadFiles=""
      DefaultPicUrl=""
      IncludePic=0
      ImagesNum=0
      NewsCode=""
      FoundErr=False
      ErrMsg=""
      His_Repeat=False
      NewsUrl=NewsArray(Arr_i)
      Title=""

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
         If FoundErr<>True Then
            '文章分页
            If NewsPaingType=1 Then
               NewsPaingNext=GetPaing(NewsCode,NPsString,NPoString,False,False)
               Do While NewsPaingNext<>"$False$"
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
                     Content=Content & NewsPaingHtml & ContentTemp
                     NewsPaingNext=GetPaing(NewsPaingNextCode,NPsString,NPoString,False,False)
                  End If
               Loop
            End If

            '过滤
            Call Filters
            Title=FpHtmlEnCode(Title)
            Call FilterScript()
            Content=Ubbcode(Content)
         End If
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
               UpDateTime=FpHtmlEncode(UpDateTime)
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
            Key=CreateKeyWord(Key)
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
         
         '转换图片相对地址为绝对地址/保存
         If CollecTest=False And SaveFiles=True then
            Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,True,NewsUrl)              
         Else
            Content=ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,False,NewsUrl)
         End If
  
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
            SqlItem="INSERT INTO Histroly(ItemID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ClassID & "','" & ClassID & "','" & SpecialID & "','" & ArticleID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',True)"
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
         ErrMsg=ErrMsg & "图片信息：图片 " & ImagesNum & " 张<br>"
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
            ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
            ErrMsg=ErrMsg & "目标文章：<font color=red>"
            If His_Result=True Then
               ErrMsg=ErrMsg & His_Title
            Else
               ErrMsg=ErrMsg & NewsUrl
            End If
            ErrMsg=ErrMsg & "</font><br>"

            ErrMsg=ErrMsg & "采集时间：" & His_CollecDate & "<br>"
            ErrMsg=ErrMsg & "文章来源：<a href='" & NewsUrl & "' target=_blank>"&NewsUrl&"</a><br>"
            ErrMsg=ErrMsg & "采集结果："
            If His_Result=False Then
               ErrMsg=ErrMsg & "失败"
               ErrMsg=ErrMsg & "<br>失败原因：" & Title
            Else
               ErrMsg=ErrMsg & "成功"
            End If            
            ErrMsg=ErrMsg & "<br>提示信息：如想再次采集，请先将该文章的历史记录<font color=red>删除</font><br>"
         End If
         If CollecTest=False And His_Repeat=False Then
            SqlItem="INSERT INTO Histroly(ItemID,ClassID,SpecialID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ClassID & "','" & ClassID & "','" & SpecialID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',False)"
            ConnItem.Execute(SqlItem)
         End If
      End If
      Call ShowMsg(ErrMsg)
      Response.Flush()'刷新
   Next
Else
   Call ShowMsg(ErrMsg)
End If

Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"">"
Response.Write "<tr>"
Response.write "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
If CollecTest=False Then
   Response.Write "数据整理中，5秒后继续......5秒后如果还没反应请点击 <a href='Admin_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext & "'><font color=red>这里</font></a> 继续<br>"
   Response.Write "<meta http-equiv=""refresh"" content=""5;url=Admin_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext  & """>"
Else
   Response.Write "<a href='Admin_ItemCollecFast.asp?ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext & "'><font color=red>请 继 续</font></a>"
End If
Response.Write "</td></tr>"
Response.Write "</table>"
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
      FoundErr=True
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
      ItemCollecDate=Arr_Item(54,ItemNumTemp)
      PaginationType=Arr_Item(56,ItemNumTemp)
      MaxCharPerPage=Arr_Item(57,ItemNumTemp)
      ReadLevel=Arr_Item(58,ItemNumTemp)
      Stars=Arr_Item(59,ItemNumTemp)
      ReadPoint=Arr_Item(60,ItemNumTemp)
      Hits=Arr_Item(61,ItemNumTemp)
      UpDateType=Arr_Item(62,ItemNumTemp)
      UpDateTime=Arr_Item(63,ItemNumTemp)
      IncludePicYn=Arr_Item(64,ItemNumTemp)
      DefaultPicYn=Arr_Item(65,ItemNumTemp)
      OnTop=Arr_Item(66,ItemNumTemp)
      Elite=Arr_Item(67,ItemNumTemp)
      Hot=Arr_Item(68,ItemNumTemp)
      SkinID=Arr_Item(69,ItemNumTemp)
      TemplateID=Arr_Item(70,ItemNumTemp)
      Script_Iframe=Arr_Item(71,ItemNumTemp)
      Script_Object=Arr_Item(72,ItemNumTemp)
      Script_Script=Arr_Item(73,ItemNumTemp)
      Script_Div=Arr_Item(74,ItemNumTemp)
      Script_Class=Arr_Item(75,ItemNumTemp)
      Script_Span=Arr_Item(76,ItemNumTemp)
      Script_Img=Arr_Item(77,ItemNumTemp)
      Script_Font=Arr_Item(78,ItemNumTemp)
      Script_A=Arr_Item(79,ItemNumTemp)
      Script_Html=Arr_Item(60,ItemNumTemp)
      CollecListNum=Arr_Item(81,ItemNumTemp)
      CollecNewsNum=Arr_Item(82,ItemNumTemp)
      Passed=Arr_Item(83,ItemNumTemp)
      SaveFiles=Arr_Item(84,ItemNumTemp)
      CollecOrder=Arr_Item(85,ItemNumTemp)
      LinkUrlYn=Arr_Item(86,ItemNumTemp)
      InputerType=Arr_Item(87,ItemNumTemp)
      Inputer=Arr_Item(88,ItemNumTemp)
      EditorType=Arr_Item(89,ItemNumTemp)
      Editor=Arr_Item(90,ItemNumTemp)
      ShowCommentLink=Arr_Item(91,ItemNumTemp)
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

'==================================================
'过程名：GetListPaing
'作  用：获取列表下一页
'参  数：无
'==================================================
Sub GetListPaing()
   If ListPaingType=1 Then
      ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
      If ListPaingNext<>"$False$"  Then
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
'过程名：SaveArticle
'作  用：保存文章
'参  数：无
'==================================================
Sub SaveArticle
   If ArticleID=0 Then
      set rs=server.createobject("adodb.recordset")
      sql="select top 1 ArticleID from LZ8_Article order by ArticleID desc" 
      rs.open sql,conn,1,1
      If rs.eof and rs.bof then
         ArticleID=1
      Else
         ArticleID=rs("ArticleID")+1
      End If
      rs.close
      set rs=nothing
   Else
      ArticleID=ArticleID+1
   End If
   set rs=server.createobject("adodb.recordset")
   sql="select top 1 * from LZ8_Article" 
   rs.open sql,conn,1,3
   rs.addnew
   rs("ArticleID")=ArticleID
   rs("ClassID")=ClassID
   rs("ClassID")=ClassID
   rs("SpecialID")=SpecialID
   rs("Title")=Title
   rs("TitleFontType")=0
   If LinkUrlYn=False Then
      rs("Content")=Content
   Else
      rs("Content")=""
      rs("LinkUrl")=NewsUrl
   End If
   rs("Keyword")=Key
   rs("Hits")=Hits
   rs("Author")=Author
   rs("CopyFrom")=CopyFrom
   rs("IncludePic")=IncludePic
   rs("Passed")=Passed
   rs("OnTop")=OnTop
   rs("Hot")=Hot
   rs("Elite")=Elite
   rs("Stars")=Stars
   rs("UpdateTime")=UpDateTime
   rs("PaginationType")=PaginationType
   rs("MaxCharPerPage")=MaxCharPerPage
   rs("ReadLevel")=ReadLevel
   rs("ReadPoint")=ReadPoint
   rs("SkinID")=SkinID
   rs("TemplateID")=TemplateID
   rs("DefaultPicUrl")=DefaultPicUrl
   rs("UploadFiles")=UploadFiles
   rs("ShowCommentLink")=ShowCommentLink
   rs("Inputer")=Inputer
   if Editor="" then Editor="五月"
   rs("Editor")=Editor
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
			
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>采 集 系 统 采 集 管 理</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_ItemStart.asp">管理首页</a> >> <font color=red>文章采集</font>  <a href="Admin_ItemStart.asp">停止采集</a></td>         
  </tr>  
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">       
  <tr> 
    <td height="22" colspan="2" aling="center">采集需要一定的时间，请耐心等待，如果网站出现暂时无法访问的情况这是正常的，采集正常结束后即可恢复。
    </td>
  </tr>
</table>
<%End Sub%>

<%
Sub TopItem2%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
    <tr>
      <td height="22" colspan="2" aling="left">本次运行：<%=Ubound(Arr_Item,2)+1%> 个项目,正在采集第 <font color=red><%= ItemNum%></font> 个项目  <font color=red><%=ItemName%></font>  的第   <font color=red><%=ListNum%></font> 页列表,该列表待采集文章  <font color=red><%=Ubound(NewsArray)+1%></font> 条。
      <br>采集统计：成功采集--<%=NewsSuccesNum%>  条文章，失败--<%=NewsFalseNum%>  条，图片--<%=ImagesNumAll%>　张。
      </td>
    </tr>
</table>
<%StartTime=Timer()%>
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

<%
'==================================================
'过程名：FootItem2
'作  用：显示该列表采集时间等信息
'参  数：无
'==================================================
Sub FootItem2()
   Dim strTemp
   OverTime=Timer()
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"">"       
   strTemp=strTemp & "<tr>"          
   strTemp=strTemp & "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
   strTemp=strTemp & "执行时间：" & CStr(FormatNumber((OverTime-StartTime)*1000,2)) & " 毫秒"
   strTemp=strTemp & "</td></tr><br>"
   strTemp=strTemp & "</table>"
   Response.write strTemp
End Sub

'==================================================
'过程名：ShowMsg
'作  用：显示信息
'参  数：无
'==================================================
Sub ShowMsg(Msg)
   Dim strTemp
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"">"       
   strTemp=strTemp & "   <tr class='tdbg'>"          
   strTemp=strTemp & "      <td height=""22"" colspan=""2"" align=""left"">"
   strTemp=strTemp & Msg
   strTemp=strTemp & "      </td>"
   strTemp=strTemp & "   </tr><br>"
   strTemp=strTemp & "</table>"
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