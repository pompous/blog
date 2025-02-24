<%
'==================================================
'过程名：Admin_ShowChannel_Name
'作  用：显示栏目名称
'参  数：ClassID ------栏目ID
'==================================================
Sub Admin_ShowChannel_Name(ClassID)   
   Dim Sqlc,Rsc,TempStr
   ClassID=Clng(ClassID)
   Sqlc ="select top 1 ClassName from xiaowei_Class Where link = 0 and ID=" & ClassID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,Conn,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="无指定栏目"   
   Else   
      TempStr=Rsc("ClassName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   response.write TempStr   
End Sub  

'==================================================
'过程名：Admin_ShowChannel_Option
'作  用：显示栏目选项
'参  数：ClassID ------栏目ID
'==================================================
Sub Admin_ShowChannel_Option(ClassID)   
   Dim Sqlp,Rsp,TempStr,sqlpp,rspp
   Sqlp ="select * from xiaowei_Class Where link = 0 and TopID = 0 order by num"   
   Set Rsp=server.CreateObject("adodb.recordset")   
   rsp.open sqlp,conn,1,1 
   Response.Write("<option value="""">请选择分类</option>") 
   If Rsp.Eof and Rsp.Bof Then
      Response.Write("<option value="""">请先添加分类</option>")
   Else
      Do while not Rsp.Eof   
         Response.Write("<option value=" & """" & Rsp("ID") & """" & "")
		 		If ClassID=Rsp("ID") Then
            	Response.Write(" selected")
         		End If
         Response.Write(">|-" & Rsp("ClassName") & "")
		 Response.Write("</option>" )  & VbCrLf
		    Sqlpp ="select * from xiaowei_Class Where TopID="&Rsp("ID")&" order by num"   
   			Set Rspp=server.CreateObject("adodb.recordset")   
   			rspp.open sqlpp,conn,1,1
			Do while not Rspp.Eof 
				Response.Write("<option value=" & """" & Rspp("ID") & """" & "")
				If ClassID=Rspp("ID") Then
            	Response.Write(" selected")
         		End If
         		Response.Write(">　|-" & Rspp("ClassName") & "")
				Response.Write("</option>" )  & VbCrLf
			Rspp.Movenext   
      		Loop
      Rsp.Movenext   
      Loop   
   End if  
End sub 


'==================================================
'过程名：Admin_ShowItem_Name
'作  用：显示项目名称
'参  数：ItemID ------项目ID
'==================================================
Sub Admin_ShowItem_Name(ItemID)   
   Dim Sqlc,Rsc,TempStr
   ItemID=Clng(ItemID)
   Sqlc ="select top 1 ItemName from Item Where ItemID=" & ItemID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="无指定项目"   
   Else   
      TempStr=Rsc("ItemName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  


'==================================================
'过程名：Admin_ShowItem_Option
'作  用：显示项目选项
'参  数：ItemID ------项目ID
'==================================================
Sub Admin_ShowItem_Option(ItemID)   
   Dim SqlI,RsI,TempStr
   ItemID=Clng(ItemID)
   SqlI ="select ItemID,ItemName from Item order by ItemID desc"   
   Set RsI=server.CreateObject("adodb.recordset")   
   RsI.Open SqlI,ConnItem,1,1
   TempStr="<select Name=""ItemID"" ID=""ItemID"">"   
   If RsI.Eof and RsI.Bof Then
      TempStr=TempStr & "<option value=""0"">请添加项目</option>"   
   Else   
      TempStr=TempStr & "<option value=""0"">请选择项目</option>"
      Do while not RsI.Eof   
         TempStr=TempStr & "<option value=" & """" & RsI("ItemID") & """" & "" 
         If ItemID=RsI("ItemID") Then
            TempStr=TempStr & " Selected"
         End If
         TempStr=TempStr & ">" & RsI("ItemName")
         TempStr=TempStr & "</option>"  
      RsI.Movenext   
      Loop   
   End if
   RsI.Close   
   Set RsI=Nothing   
   TempStr=TempStr & "</select>"
   Response.Write TempStr   
End sub   

'==================================================
'函数名：GetHttpPage
'作  用：获取网页源码
'参  数：HttpUrl ------网页地址
'==================================================
Function GetHttpPage(HttpUrl)
   If IsNull(HttpUrl)=True Or Len(HttpUrl)<18 Or HttpUrl="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Dim Http
   Set Http=server.createobject("MSXML2.XMLHTTP")
   Http.open "GET",HttpUrl,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
   Set Http=Nothing
   If Err.number<>0 then
      Err.Clear
   End If
End Function

'==================================================
'函数名：BytesToBstr
'作  用：将获取的源码转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'==================================================
'函数名：PostHttpPage
'作  用：登录
'==================================================
Function PostHttpPage(RefererUrl,PostUrl,PostData) 
    Dim xmlHttp 
    Dim RetStr      
    Set xmlHttp = CreateObject("Msxml2.XMLHTTP")  
    xmlHttp.Open "POST", PostUrl, False
    XmlHTTP.setRequestHeader "Content-Length",Len(PostData) 
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.setRequestHeader "Referer", RefererUrl
    xmlHttp.Send PostData 
    If Err.Number <> 0 Then 
        Set xmlHttp=Nothing
        PostHttpPage = "$False$"
        Exit Function
    End If
    PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"GB2312")
    Set xmlHttp = nothing
End Function 

'==================================================
'函数名：UrlEncoding
'作  用：转换编码
'==================================================
Function UrlEncoding(DataStr)
    Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr,Si,1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
               InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode  And &HFF00)\ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'函数名：GetBody
'作  用：截取字符串
'参  数：ConStr ------将要截取的字符串
'参  数：StartStr ------开始字符串
'参  数：OverStr ------结束字符串
'参  数：IncluL ------是否包含StartStr
'参  数：IncluR ------是否包含OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
      GetBody="$False$"
      Exit Function
   End If
   Dim ConStrTemp
   Dim Start,Over
   ConStrTemp=Lcase(ConStr)
   StartStr=Lcase(StartStr)
   OverStr=Lcase(OverStr)
   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
   If Start<=0 then
      GetBody="$False$"
      Exit Function
   Else
      If IncluL=False Then
         Start=Start+LenB(StartStr)
      End If
   End If
   Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
   If Over<=0 Or Over<=Start then
      GetBody="$False$"
      Exit Function
   Else
      If IncluR=True Then
         Over=Over+LenB(OverStr)
      End If
   End If
   GetBody=MidB(ConStr,Start,Over-Start)
End Function


'==================================================
'函数名：GetArray
'作  用：提取链接地址，以$Array$分隔
'参  数：ConStr ------提取地址的原字符
'参  数：StartStr ------开始字符串
'参  数：OverStr ------结束字符串
'参  数：IncluL ------是否包含StartStr
'参  数：IncluR ------是否包含OverStr
'==================================================
Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or  IsNull(StartStr)=True Or IsNull(OverStr)=True Then
      GetArray="$False$"
      Exit Function
   End If
   Dim TempStr,TempStr2,objRegExp,Matches,Match
   TempStr=""
   Set objRegExp = New Regexp 
   objRegExp.IgnoreCase = True 
   objRegExp.Global = True
   objRegExp.Pattern = "("&StartStr&").+?("&OverStr&")"
   Set Matches =objRegExp.Execute(ConStr) 
   For Each Match in Matches
      TempStr=TempStr & "$Array$" & Match.Value
   Next 
   Set Matches=nothing

   If TempStr="" Then
      GetArray="$False$"
      Exit Function
   End If
   TempStr=Right(TempStr,Len(TempStr)-7)
   If IncluL=False then
      objRegExp.Pattern =StartStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   If IncluR=False then
      objRegExp.Pattern =OverStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   Set objRegExp=nothing
   Set Matches=nothing
   
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   TempStr=Replace(TempStr,"(","")
   TempStr=Replace(TempStr,")","")

   If TempStr="" then
      GetArray="$False$"
   Else
      GetArray=TempStr
   End if
End Function


'==================================================
'函数名：DefiniteUrl
'作  用：将相对地址转换为绝对地址
'参  数：PrimitiveUrl ------要转换的相对地址
'参  数：ConsultUrl ------当前网页地址
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
   Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
   If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" or ConsultUrl="$False$" Then
      DefiniteUrl="$False$"
      Exit Function
   End If
   If Left(Lcase(ConsultUrl),7)<>"http://" Then
      ConsultUrl= "http://" & ConsultUrl
   End If
   ConsultUrl=Replace(ConsultUrl,"\","/")
   ConsultUrl=Replace(ConsultUrl,"://",":\\")
   PrimitiveUrl=Replace(PrimitiveUrl,"\","/")

   If Right(ConsultUrl,1)<>"/" Then
      If Instr(ConsultUrl,"/")>0 Then
         If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then   
         Else
            ConsultUrl=ConsultUrl & "/"
         End If
      Else
         ConsultUrl=ConsultUrl & "/"
      End If
   End If
   ConArray=Split(ConsultUrl,"/")

   If Left(LCase(PrimitiveUrl),7) = "http://" then
      DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
   ElseIf Left(PrimitiveUrl,1) = "/" Then
      DefiniteUrl=ConArray(0) & PrimitiveUrl
   ElseIf Left(PrimitiveUrl,2)="./" Then
      PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-2)
      If Right(ConsultUrl,1)="/" Then   
         DefiniteUrl=ConsultUrl & PrimitiveUrl
      Else
         DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
      End If
   ElseIf Left(PrimitiveUrl,3)="../" then
      Do While Left(PrimitiveUrl,3)="../"
         PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
         Pi=Pi+1
      Loop            
      For Ci=0 to (Ubound(ConArray)-1-Pi)
         If DefiniteUrl<>"" Then
            DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
         Else
            DefiniteUrl=ConArray(Ci)
         End If
      Next
      DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
   Else
      If Instr(PrimitiveUrl,"/")>0 Then
         PriArray=Split(PrimitiveUrl,"/")
         If Instr(PriArray(0),".")>0 Then
            If Right(PrimitiveUrl,1)="/" Then
               DefiniteUrl="http:\\" & PrimitiveUrl
            Else
               If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
                  DefiniteUrl="http:\\" & PrimitiveUrl
               Else
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               End If
            End If      
         Else
            If Right(ConsultUrl,1)="/" Then   
               DefiniteUrl=ConsultUrl & PrimitiveUrl
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
            End If
         End If
      Else
         If Instr(PrimitiveUrl,".")>0 Then
            If Right(ConsultUrl,1)="/" Then
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=ConsultUrl & PrimitiveUrl
               End If
            Else
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
               End If
            End If
         Else
            If Right(ConsultUrl,1)="/" Then
               DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
            End If         
         End If
      End If
   End If
   If Left(DefiniteUrl,1)="/" then
     DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
   End if
   If DefiniteUrl<>"" Then
      DefiniteUrl=Replace(DefiniteUrl,"//","/")
      DefiniteUrl=Replace(DefiniteUrl,":\\","://")
   Else
      DefiniteUrl="$False$"
   End If
End Function

'==================================================
'函数名：ReplaceSaveRemoteFile
'作  用：替换、保存远程图片
'参  数：ConStr ------ 要替换的字符串
'参  数：SaveTf ------ 是否保存文件，False不保存，True保存
'参  数: TistUrl------ 当前网页地址
'==================================================
Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
   If ConStr="$False$" or ConStr="" or strInstallDir="" or strChannelDir="" Then
      ReplaceSaveRemoteFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<img.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
      Re.Pattern ="src\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   Set Matches=nothing
   Set Re=nothing
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSaveRemoteFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")

   Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
   DtNow=Now()
   If SaveTf=True then
 '***********************************
      SavePath= strChannelDir & ""&SitePath&""&SiteUp&"/" & year(DtNow) & right("0" & month(DtNow),2) & "/"
	  response.write "链接路径：" & savepath & "<br>"
      Arr_Path=Split(SavePath,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
   End If

   '去掉重复图片开始
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   '去掉重复图片结束

   '转换相对图片地址开始
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   '转换相对图片地址结束

   '图片替换/保存
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True

   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      If RemoteFileUrl<>"$False$" And SaveTf=True Then'保存图片
         ArrSaveFileName = Split(RemoteFileurl,".")
	 strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'文件类型
         If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
            UploadFiles=""
            ReplaceSaveRemoteFile=ConStr
            Exit Function
         End If

         Randomize
         RanNum=Int(900*Rnd)+100
	 strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
         Re.Pattern =TempArray(Tempi)
	 If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
'********************************
            PathTemp=SavePath & strFileName
            ConStr=Re.Replace(ConStr,PathTemp)
            Re.Pattern=strInstallDir & strChannelDir & "/"
            UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")

If aspjpeg=1 then
			
Dim Jpeg,RV_img,ImageMode,Uprequest
RV_img=SavePath&strFileName

Set Jpeg = Server.CreateObject("Persits.Jpeg") 
Jpeg.Open Server.MapPath(RV_img)
   
Jpeg.Canvas.Font.Color = "&H"&""&Color1&""
Jpeg.Canvas.Font.Size = ""&FontSize&""
Jpeg.Canvas.Font.Family = ""&FontFamily&""
Jpeg.Canvas.Font.ShadowColor = "&H"&""&Color2&""
Jpeg.Canvas.Font.ShadowXoffset = 1
Jpeg.Canvas.Font.ShadowYoffset = 1 
'Jpeg.Canvas.Font.Quality = 1
Jpeg.Canvas.Font.Bold = False
Jpeg.Canvas.Print 10, 10, ImageMode
Jpeg.Canvas.Print 8,5,""&Fonttext&""
Jpeg.Save Server.MapPath(RV_img)

Set Jpeg = Nothing 
Set Uprequest=nothing

End if

         Else
            PathTemp=RemoteFileUrl
            ConStr=Re.Replace(ConStr,PathTemp)
            'UploadFiles=UploadFiles & "|" & RemoteFileUrl
         End If
      ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'不保存图片
         Re.Pattern =TempArray(Tempi)
         ConStr=Re.Replace(ConStr,RemoteFileUrl)
         UploadFiles=UploadFiles & "|" & RemoteFileUrl
      End If
   Next   
   Set Re=nothing
   If UploadFiles<>"" Then
      UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
   End If
   ReplaceSaveRemoteFile=ConStr
End function

'==================================================
'函数名：ReplaceSwfFile
'作  用：解析动画路径
'参  数：ConStr ------ 要替换的字符串
'参  数: TistUrl------ 当前网页地址
'==================================================
Function ReplaceSwfFile(ConStr,TistUrl)
   If ConStr="$False$" or ConStr="" or TistUrl="" or TistUrl="$False$" Then
      ReplaceSwfFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<object.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="value\s*=\s*.+?\.swf"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
      Re.Pattern ="value\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSwfFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")

   Set Matches=nothing
   Set Re=nothing

   '去掉重复文件开始
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   '去掉重复文件结束

   '转换相对地址开始
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   '转换相对地址结束

   '替换
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True
   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      Re.Pattern =TempArray(Tempi)
      ConStr=Re.Replace(ConStr,RemoteFileUrl)
   Next   
   Set Re=nothing
   ReplaceSwfFile=ConStr
End function

'==================================================
'过程名：SaveRemoteFile
'作  用：保存远程的文件到本地
'参  数：LocalFileName ------ 本地文件名
'参  数：RemoteFileUrl ------ 远程文件URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
    SaveRemoteFile=True
	dim Ads,Retrieval,GetRemoteData
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
        If .Readystate<>4 then
            SaveRemoteFile=False
            Exit Function
        End If
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set Ads=nothing
end Function

'==================================================
'函数名：FpHtmlEnCode
'作  用：标题过滤
'参  数：fString ------字符串
'==================================================
Function FpHtmlEnCode(fString)
   If IsNull(fString)=False or fString<>"" or fString<>"$False$" Then
       fString=nohtml(fString)
       fString=FilterJS(fString)
       fString = Replace(fString,"&nbsp;"," ")
       fString = Replace(fString,"&quot;","")
       fString = Replace(fString,"&#39;","")
       fString = replace(fString, ">", "")
       fString = replace(fString, "<", "")
       fString = Replace(fString, CHR(9), " ")'&nbsp;
       fString = Replace(fString, CHR(10), "")
       fString = Replace(fString, CHR(13), "")
       fString = Replace(fString, CHR(34), "")
       fString = Replace(fString, CHR(32), " ")'space
       fString = Replace(fString, CHR(39), "")
       fString = Replace(fString, CHR(10) & CHR(10),"")
       fString = Replace(fString, CHR(10)&CHR(13), "")
       fString=Trim(fString)
       FpHtmlEnCode=fString
   Else
       FpHtmlEnCode="$False$"
   End If
End Function

'==================================================
'函数名：GetPaing
'作  用：获取分页
'==================================================
Function GetPaing(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" Or StartStr="" Or OverStr="" or IsNull(ConStr)=True or IsNull(StartStr)=True Or IsNull(OverStr)=True Then
   GetPaing="$False$"
   Exit Function
End If

Dim Start,Over,ConTemp,TempStr
TempStr=LCase(ConStr)
StartStr=LCase(StartStr)
OverStr=LCase(OverStr)
Over=Instr(1,TempStr,OverStr)
If Over<=0 Then
   GetPaing="$False$"
   Exit Function
Else
   If IncluR=True Then
      Over=Over+Len(OverStr)
   End If
End If
TempStr=Mid(TempStr,1,Over)
Start=InstrRev(TempStr,StartStr)
If IncluL=False Then
   Start=Start+Len(StartStr)
End If

If Start<=0 Or Start>=Over Then
   GetPaing="$False$"
   Exit Function
End If
ConTemp=Mid(ConStr,Start,Over-Start)

ConTemp=Trim(ConTemp)
ConTemp=Replace(ConTemp," ","")
ConTemp=Replace(ConTemp,",","")
ConTemp=Replace(ConTemp,"'","")
ConTemp=Replace(ConTemp,"""","")
ConTemp=Replace(ConTemp,">","")
ConTemp=Replace(ConTemp,"<","")
ConTemp=Replace(ConTemp,"&nbsp;","")
GetPaing=ConTemp
End Function

'==================================================
'函数名：ScriptHtml
'作  用：过滤html标记
'参  数：ConStr ------ 要过滤的字符串
'==================================================
Function ScriptHtml(Byval ConStr,TagName,FType)
    Dim Re
    Set Re=new RegExp
    Re.IgnoreCase =true
    Re.Global=True
    Select Case FType
    Case 1
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 2
       Re.Pattern="<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 3
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
       Re.Pattern="</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    End Select
    ScriptHtml=ConStr
    Set Re=Nothing
End Function

Function CheckDir(byval FolderPath)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(folderpath)) then
	'存在
		CheckDir = True
	Else
	'不存在
		CheckDir = False
	End if
	Set fso = nothing
End Function
Function MakeNewsDir(byval foldername)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(Server.MapPath(foldername))
        If fso.FolderExists(Server.MapPath(foldername)) Then
           MakeNewsDir = True
        Else
           MakeNewsDir = False
        End If
	Set fso = nothing
End Function

'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'**************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>[90站长]采集系统错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='../images/Admin_css.css' rel='stylesheet' type='text/css'></head><body topmargin='0' leftmargin='0'><table border='0' width='100%' cellspacing='0' cellpadding='0' height='126' id='table1'><tr><td width='300'></td><td valign='top'><br><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='admintable' align=center>" & vbcrlf
	strErr=strErr & "  <tr><td class='admintitle' style=""text-align:center;"">错误信息</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td style='padding:10px;'><b>产生错误的可能原因：</b>" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table></td><td valign='top' width='300'>　</td></tr></table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'过程名：WriteSucced
'作  用：显示成功提示信息
'参  数：无
'**************************************************
sub WriteSucced(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>[90站长]采集系统成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='../images/Admin_css.css' rel='stylesheet' type='text/css'></head><body topmargin='0' leftmargin='0'><table border='0' width='100%' cellspacing='0' cellpadding='0' height='126' id='table1'><tr><td width='300'></td><td valign='top'><br><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=3 cellspacing=2 border=0 class='admintable' align=center>" & vbcrlf
	strErr=strErr & "  <tr><td class='admintitle' style=""text-align:center;"">恭喜你</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td align='center' style=""padding:20px;"">" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td align='center'><a href='Admin_ItemStart.asp'>&lt;&lt; 返回首页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table></td><td valign='top' width='300'>　</td></tr></table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'函数名：ShowPage
'作  用：显示“上一页 下一页”等信息
'参  数：sFileName  ----链接地址
'       TotalNumber ----总数量
'       MaxPerPage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
function ShowPage(sFileName,TotalNumber,MaxPerPage,ShowTotal,ShowAllPages,strUnit)
	dim TotalPage,strTemp,strUrl,i

	if TotalNumber=0 or MaxPerPage=0 or isNull(MaxPerPage) then
		ShowPage=""
		exit function
	end if
	if totalnumber mod maxperpage=0 then
    	TotalPage= totalnumber \ maxperpage
  	else
    	TotalPage= totalnumber \ maxperpage+1
  	end if
	if CurrentPage>TotalPage then CurrentPage=TotalPage
		
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    	strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    	strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if CurrentPage>=TotalPage then
    	strTemp=strTemp & "下一页 尾页"
  	else
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    	strTemp=strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>页 "
        strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;&nbsp;转到第<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;""'>页"
	end if
	strTemp=strTemp & "</td></tr></table>"
	ShowPage=strTemp
end function
Dim fengyename
fengyename="["&"y"&"a"&"o"&"_"&"p"&"a"&"g"&"e"&"]"
'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'**************************************************
'函数名：CreateKeyWord
'作  用：由给定的字符串生成关键字
'参  数：Constr---要生成关键字的原字符串
'返回值：生成的关键字
'**************************************************
Function CreateKeyWord(byval Constr,Num)
   If Constr="" or IsNull(Constr)=True or Constr="$False$" Then
      CreateKeyWord="$False$"
      Exit Function
   End If
   If Num="" or IsNumeric(Num)=False Then
      Num=2
   End If
   Constr=Replace(Constr,CHR(32),"")
   Constr=Replace(Constr,CHR(9),"")
   Constr=Replace(Constr,"&nbsp;","")
   Constr=Replace(Constr," ","")
   Constr=Replace(Constr,"(","")
   Constr=Replace(Constr,")","")
   Constr=Replace(Constr,"<","")
   Constr=Replace(Constr,">","")
   Constr=Replace(Constr,"""","")
   Constr=Replace(Constr,"?","")
   Constr=Replace(Constr,"*","")
   Constr=Replace(Constr,"|","")
   Constr=Replace(Constr,",","")
   Constr=Replace(Constr,"，","")
   Constr=Replace(Constr,".","")
   Constr=Replace(Constr,"/","")
   Constr=Replace(Constr,"\","")
   Constr=Replace(Constr,"-","")
   Constr=Replace(Constr,"@","")
   Constr=Replace(Constr,"#","")
   Constr=Replace(Constr,"$","")
   Constr=Replace(Constr,"%","")
   Constr=Replace(Constr,"&","")
   Constr=Replace(Constr,"+","")
   Constr=Replace(Constr,":","")
   Constr=Replace(Constr,"：","")   
   Constr=Replace(Constr,"‘","")
   Constr=Replace(Constr,"“","")
   Constr=Replace(Constr,"”","")         
   Dim i,ConstrTemp
   For i=1 To Len(Constr)
      ConstrTemp=ConstrTemp & Mid(Constr,i,Num) & "|"
   Next
   If Len(ConstrTemp)<11 Then
      ConstrTemp=ConstrTemp
   Else
      ConstrTemp=Left(ConstrTemp,11)
   End If
   CreateKeyWord=ConstrTemp
End Function

Function CheckUrl(strUrl)
   Dim Re
   Set Re=new RegExp
   Re.IgnoreCase =true
   Re.Global=True
   Re.Pattern="http://([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
   If Re.test(strUrl)=True Then
      CheckUrl=strUrl
   Else
      CheckUrl="$False$"
   End If
   Set Rs=Nothing
End Function
%>
