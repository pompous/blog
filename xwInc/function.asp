<%
Function Mydb(MySqlstr,MyDBType)
	Select Case MyDBType
	Case 0 : Conn.Execute(MySqlstr) : Dataquery = Dataquery + 1
	Case 1 : Set Mydb = Conn.Execute(MySqlstr) : Dataquery = Dataquery + 1
	Case 2 : Set Mydb = Server.CreateObject("Adodb.Recordset") : Mydb.Open MySqlstr,Conn,1,1 : Dataquery = Dataquery + 1
	case 3:
		set db = server.createobject("Adodb.Recordset")
		db.open sqlstr, conn, 1, 3
	End Select
End Function

function CheckStr(str) 
    CheckStr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","") 
   CheckStr=replace(replace(replace(replace(CheckStr,"'",""),"and",""),"insert",""),"set","") 
    CheckStr=replace(replace(replace(replace(CheckStr,"select",""),"update",""),"delete%20from",""),chr(34),"&quot;") 
end function

Function LoseHtml(ContentStr)
 Dim ClsTempLoseStr,regEx
        ClsTempLoseStr = Cstr(ContentStr)
 		Set regEx = New RegExp
       regEx.Pattern = "(\<.+?\>)"
       regEx.IgnoreCase = True
       regEx.Global = True
       ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	RegEx.Pattern = "(&.+?;)"
	ClsTempLoseStr = RegEx.Replace(ClsTempLoseStr, "")
	ClsTempLoseStr = Replace(ClsTempLoseStr,VbCrlf,"")
	ClsTempLoseStr = Replace(ClsTempLoseStr,VbCr,"")
	ClsTempLoseStr = Replace(ClsTempLoseStr,VbLf,"")
	ClsTempLoseStr = Replace(ClsTempLoseStr,"  ","")
	ClsTempLoseStr = Replace(ClsTempLoseStr,"  ","")
	ClsTempLoseStr = Replace(ClsTempLoseStr,"""","'")
	ClsTempLoseStr = Replace(ClsTempLoseStr,"[code]","")
	ClsTempLoseStr = Replace(ClsTempLoseStr,"<!--","")
	ClsTempLoseStr = Trim(ClsTempLoseStr)
 LoseHtml = ClsTempLoseStr
End function

Function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    'fString = Replace(fString, "http://static3.photo.sina.com.cn", "photo.asp?url=http://static3.photo.sina.com.cn")

    dvHTMLEncode = fString
end if
end Function

function HasChinese(str) 
HasChinese = false 
dim i 
for i=1 to Len(str) 
if Asc(Mid(str,i,1)) < 0 then 
HasChinese = true 
exit for 
end if 
next 
end function

Function replacecolor(Str)
	Dim re,s
	S=Str
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="("& KeyWord &")"
	s=re.Replace(s,"<font color='red'>"& KeyWord &"</font>")
	Set Re=Nothing
	replacecolor=s
End Function

function iparray(ipstr)
 dim t,ipx,ipfb
 if not isnull(ipstr) then
        t = 0
 ipx=""
 ipfb = split(ipstr, ".",4)
  for t = 0 to 2
  ipx = ipx&ipfb(t)&"."
  next
 iparray = ipx&"*"
 end if
end function

Public Function GlHtml(ByVal str) 
    If IsNull(str) Or Trim(str) = "" Then
        GlHtml = ""
        Exit Function
    End If
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "(\<.[^\<]*\>)"
    str = re.Replace(str, " ")
    re.Pattern = "(\<\/[^\<]*\>)"
    str = re.Replace(str, " ")
    Set re = Nothing
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    GlHtml = str
End Function

Function FormatDate(DateAndTime,para)
	On Error Resume Next
	Dim y, m, d, h, mi, s, strDateTime
	FormatDate = DateAndTime
	
	If Not IsNumeric(para) Then Exit Function
	If Not IsDate(DateAndTime) Then Exit Function
	y = Mid(CStr(Year(DateAndTime)),3)
	m = CStr(Month(DateAndTime))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(DateAndTime))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(DateAndTime))
	If Len(h) = 1 Then h = "0" & h
	mi = CStr(Minute(DateAndTime))
	If Len(mi) = 1 Then mi = "0" & mi
	s = CStr(Second(DateAndTime))
	If Len(s) = 1 Then s = "0" & s
	
	Select Case para
		Case "1"
			strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
		Case "2"
			strDateTime = y & "-" & m & "-" & d
		Case "3"
			strDateTime = y & "/" & m & "/" & d
		Case "4"
			strDateTime = y & "年" & m & "月" & d & "日"
		Case "5"
			strDateTime = m & "-" & d
		Case "6"
			strDateTime = m & "/" & d
		Case "7"
			strDateTime = m & "月" & d & "日"
		Case "8"
			strDateTime = y & "年" & m & "月"
		Case "9"
			strDateTime = y & "-" & m
		Case "10"
			strDateTime = y & "/" & m
		Case "11"
			y = CStr(Year(DateAndTime))
			strDateTime = y & "-" & m & "-" & d
		Case "12"
			y = CStr(Year(DateAndTime))
			strDateTime = y & m & d & "_" & h & mi & s			
		Case Else
			strDateTime = DateAndTime
		End Select
		
	FormatDate = strDateTime
End Function


'=================================================
'过程名：ManualPagination
'作  用：采用手动分页方式显示文章具体的内容
'参  数：str1,str2,str3
'=================================================
Function ManualPagination(str1,str2)
	dim ArticleId,strContent,CurrentPage
	dim ContentLen,pages,i
	dim arrContent,ManualPagination_Tmp
	ArticleId = str1
	strContent = str2
	ContentLen=strContent
	CurrentPage=b
	if Instr(strContent,"[xiaowei_page]")<=0 then
		ManualPagination_Tmp = ManualPagination_Tmp & strContent
		ManualPagination_Tmp = ManualPagination_Tmp & "</p>"
	else
		arrContent=split(strContent,"[xiaowei_page]")

		pages=Ubound(arrContent)+1
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		ManualPagination_Tmp = ManualPagination_Tmp & arrContent(CurrentPage-1)

		ManualPagination_Tmp = ManualPagination_Tmp & "</p><div id=""clear""></div><div id=""vmoviesab""><table border=""0""  cellspacing=""5"" cellpadding=""2"" align=""center""><tr>"

		if CurrentPage>1 then

			ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""55"" class=""page_css_1_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & CurrentPage-1 & ".html"
						ManualPagination_Tmp = ManualPagination_Tmp & "'>上一页</a></TD>"
		else
			ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">上一页</TD>"
		end if
		for i=1 to pages
			if i=CurrentPage then
				ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""25"" class=""page_css_2_1"">" & cstr(i) & "</TD>"
			else
				ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""25"" class=""page_css_2_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & i & ".html" 
				ManualPagination_Tmp = ManualPagination_Tmp & "'>" & i & "</a></TD>"
			end if
			'if (i Mod 10) = 0 then ManualPagination_Tmp = ManualPagination_Tmp & "<br>"
		next
		if CurrentPage<pages then
			ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""55""  class=""page_css_1_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & CurrentPage+1 & ".html"
			ManualPagination_Tmp = ManualPagination_Tmp & "'>下一页</a></TD>"
		else
			ManualPagination_Tmp = ManualPagination_Tmp & "<TD width=""55""  class=""page_css_1_1"">下一页</TD>"
		end if
		ManualPagination_Tmp = ManualPagination_Tmp & "</tr></table></div>"			
	end if
	ManualPagination = ManualPagination_Tmp
end Function



'=================================================
'过程名：ManualPagination2
'作  用：采用手动分页方式显示文章具体的内容
'参  数：str1,str2,str3
'=================================================
Function ManualPagination2(str1,str2)
	dim ArticleId,strContent,CurrentPage
	dim ContentLen,pages,i
	dim arrContent,ManualPagination_Tmp
	ArticleId = str1
	strContent = str2
	ContentLen=strContent
	CurrentPage=b
	if Instr(strContent,"[xiaowei_page]")<=0 then
		ManualPagination_Tmp = ManualPagination_Tmp & strContent
		ManualPagination_Tmp = ManualPagination_Tmp & "</p>"
	else
		arrContent=split(strContent,"[xiaowei_page]")

		pages=Ubound(arrContent)+1
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		ManualPagination_Tmp = ManualPagination_Tmp & arrContent(CurrentPage-1)

		ManualPagination_Tmp = ManualPagination_Tmp & "</p><div id=""clear""></div><div id=""page""><ul>"
		if CurrentPage>1 then
			ManualPagination_Tmp = ManualPagination_Tmp & "<li><a href='page_" & ArticleId & "_" & CurrentPage-1 & ".html"
			ManualPagination_Tmp = ManualPagination_Tmp & "'>上一页</a></li>"
		else
			ManualPagination_Tmp = ManualPagination_Tmp & "<li><span>上一页</span></li>"
		end if
		for i=1 to pages
			if i=CurrentPage then
				ManualPagination_Tmp = ManualPagination_Tmp & "<li><span>" & cstr(i) & "</span></li>"
			else
				ManualPagination_Tmp = ManualPagination_Tmp & "<li><a href='Article_" & ArticleId & "_" & i & ".html" 
				ManualPagination_Tmp = ManualPagination_Tmp & "'>" & i & "</a></li>"
			end if
			'if (i Mod 10) = 0 then ManualPagination_Tmp = ManualPagination_Tmp & "<br>"
		next
		if CurrentPage<pages then
			ManualPagination_Tmp = ManualPagination_Tmp & "<li><a href='Article_" & ArticleId & "_" & CurrentPage+1 & ".html"
			ManualPagination_Tmp = ManualPagination_Tmp & "'>下一页</a></li>"
		else
			ManualPagination_Tmp = ManualPagination_Tmp & "<li><span>下一页</span></li>"
		end if
		ManualPagination_Tmp = ManualPagination_Tmp & "</ul></div>"				
	end if
	ManualPagination2 = ManualPagination_Tmp
end Function

'=================================================
'过程名：AutoPagination
'作  用：采用自动分页方式显示文章具体的内容
'参  数：str1,str2,str3
'=================================================
Function AutoPagination(str1,str2,str3)
	dim AutoPagination_Tmp
	dim ArticleId,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleId = str1
	strContent = Lcase(str2)
	MaxPerPage = str3
	ContentLen=len(strContent)
	CurrentPage=b
	if ContentLen<=800 then
		AutoPagination_Tmp = AutoPagination_Tmp & strContent
		AutoPagination_Tmp = AutoPagination_Tmp & ""
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\MaxPerPage
		if MaxPerPage*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=ContentLen          '最大误差范围
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=MaxPerPage*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=MaxPerPage*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		
		if EndPoint < BeginPoint then
			'BeginPoint = BeginPoint + str4
			'EndPoint = BeginPoint + str4
		end if

		On Error Resume Next
		AutoPagination_Tmp = AutoPagination_Tmp & mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		If Err Then
			Err.clear
			'response.Write "BeginPoint = "& BeginPoint
			'response.Write "<br>"
			'response.Write "EndPoint = "& EndPoint
			AutoPagination_Tmp = AutoPagination_Tmp & "<br><p align=center style='color:red;'>对不起，自动分页错误，请直接点下一页即可接上页继续。</p><br>"
		End If

		
		AutoPagination_Tmp = AutoPagination_Tmp & "</p><div id=""clear""></div><div id=""vmoviesab""><table border=""0""  cellspacing=""5"" cellpadding=""2"" align=""center""><tr>"
		if CurrentPage>1 then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & CurrentPage-1 
			AutoPagination_Tmp = AutoPagination_Tmp & ".html'>上一页</a></TD>"
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">上一页</TD>"
		end if
		for i=1 to pages
			if i=CurrentPage then
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1"">" & cstr(i) & "</TD>"
			else
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & i 
				AutoPagination_Tmp = AutoPagination_Tmp & ".html'>" & i & "</a></TD>"
			end if
			'if (i Mod 10) = 0 then AutoPagination_Tmp = AutoPagination_Tmp & "<br>"
		next
		if CurrentPage<pages then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1""><a href='"&SitePath&"page/?" & ArticleId & "_" & CurrentPage+1 
			AutoPagination_Tmp = AutoPagination_Tmp & ".html'>下一页</a></TD>"
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">下一页</span>"
		end if
		AutoPagination_Tmp = AutoPagination_Tmp & "</tr></table></div>"
	end if
	AutoPagination = AutoPagination_Tmp
end Function


'=================================================
'过程名：AutoPagination1
'作  用：采用自动分页方式显示文章具体的内容,Asp模式
'参  数：str1,str2,str3
'=================================================
Function AutoPagination1(str1,str2,str3)
	dim AutoPagination_Tmp
	dim ArticleId,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleId = str1
	strContent = Lcase(str2)
	MaxPerPage = str3
	ContentLen=len(strContent)
	CurrentPage=trim(request("Page"))
	if ContentLen<=MaxPerPage then
		AutoPagination_Tmp = AutoPagination_Tmp & strContent
		AutoPagination_Tmp = AutoPagination_Tmp & ""
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\MaxPerPage
		if MaxPerPage*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=ContentLen          '最大误差范围
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=MaxPerPage*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=MaxPerPage*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		
		if EndPoint < BeginPoint then
			'BeginPoint = BeginPoint + str4
			'EndPoint = BeginPoint + str4
		end if

		On Error Resume Next
		AutoPagination_Tmp = AutoPagination_Tmp & mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		If Err Then
			Err.clear
			'response.Write "BeginPoint = "& BeginPoint
			'response.Write "<br>"
			'response.Write "EndPoint = "& EndPoint
			AutoPagination_Tmp = AutoPagination_Tmp & "</p><div id=""clear""></div><p align=center style='color:red;'>对不起，自动分页错误，请直接点下一页即可接上页继续。</p>"
		End If

		
		AutoPagination_Tmp = AutoPagination_Tmp & "</p><div id=""clear""></div><div id=""vmoviesab""><table border=""0""  cellspacing=""5"" cellpadding=""2"" align=""center""><tr>"
		if CurrentPage>1 then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1""><a href='List.asp?ID=" & ArticleId & "&Page=" & CurrentPage-1 
			AutoPagination_Tmp = AutoPagination_Tmp & "'>上一页</a></TD>"
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">上一页</TD>"
		end if
		for i=1 to pages
			if i=CurrentPage then
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1"">" & cstr(i) & "</TD>"
			else
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1""><a href='List.asp?ID=" & ArticleId & "&Page=" & i 
				AutoPagination_Tmp = AutoPagination_Tmp & "'>" & i & "</a></TD>"
			end if
			'if (i Mod 12) = 0 then AutoPagination_Tmp = AutoPagination_Tmp & "</ul><ul>"
		next
		if CurrentPage<pages then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1""><a href='List.asp?ID=" & ArticleId & "&Page=" & CurrentPage+1 
			AutoPagination_Tmp = AutoPagination_Tmp & "'>下一页</a></TD>"
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">下一页</TD>"
		end if
		AutoPagination_Tmp = AutoPagination_Tmp & "</tr></table></div>"
	end if
	AutoPagination1 = AutoPagination_Tmp
end Function


'=================================================
'过程名：AutoPagination2
'作  用：采用自动分页方式显示文章具体的内容,伪静态模式
'参  数：str1,str2,str3
'=================================================
Function AutoPagination2(str1,str2,str3)
	dim AutoPagination_Tmp
	dim ArticleId,strContent,CurrentPage
	dim ContentLen,MaxPerPage,pages,i,lngBound
	dim BeginPoint,EndPoint
	ArticleId = str1
	strContent = Lcase(str2)
	MaxPerPage = str3
	ContentLen=len(strContent)
	CurrentPage=b
	if ContentLen<=MaxPerPage then
		AutoPagination_Tmp = AutoPagination_Tmp & strContent
		AutoPagination_Tmp = AutoPagination_Tmp & ""
	else
		if CurrentPage="" then
			CurrentPage=1
		else
			CurrentPage=Cint(CurrentPage)
		end if
		pages=ContentLen\MaxPerPage
		if MaxPerPage*pages<ContentLen then
			pages=pages+1
		end if
		lngBound=ContentLen          '最大误差范围
		if CurrentPage<1 then CurrentPage=1
		if CurrentPage>pages then CurrentPage=pages

		dim lngTemp
		dim lngTemp1,lngTemp1_1,lngTemp1_2,lngTemp1_1_1,lngTemp1_1_2,lngTemp1_1_3,lngTemp1_2_1,lngTemp1_2_2,lngTemp1_2_3
		dim lngTemp2,lngTemp2_1,lngTemp2_2,lngTemp2_1_1,lngTemp2_1_2,lngTemp2_2_1,lngTemp2_2_2
		dim lngTemp3,lngTemp3_1,lngTemp3_2,lngTemp3_1_1,lngTemp3_1_2,lngTemp3_2_1,lngTemp3_2_2
		dim lngTemp4,lngTemp4_1,lngTemp4_2,lngTemp4_1_1,lngTemp4_1_2,lngTemp4_2_1,lngTemp4_2_2
		dim lngTemp5,lngTemp5_1,lngTemp5_2
		dim lngTemp6,lngTemp6_1,lngTemp6_2
		
		if CurrentPage=1 then
			BeginPoint=1
		else
			BeginPoint=MaxPerPage*(CurrentPage-1)+1
			
			lngTemp1_1_1=instr(BeginPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(BeginPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(BeginPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(BeginPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(BeginPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(BeginPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=BeginPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2
				else
					lngTemp1=lngTemp1_1+8
				end if
			end if

			lngTemp2_1_1=instr(BeginPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(BeginPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(BeginPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(BeginPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=BeginPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngtemp2=lngTemp2_2
				else
					lngTemp2=lngTemp2_1+4
				end if
			end if

			lngTemp3_1_1=instr(BeginPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(BeginPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(BeginPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(BeginPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=BeginPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2
				else
					lngTemp3=lngTemp3_1+5
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2

			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>BeginPoint and lngTemp<=BeginPoint+lngBound then
				BeginPoint=lngTemp
			else
				lngTemp4_1_1=instr(BeginPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(BeginPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(BeginPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(BeginPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=BeginPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2
					else
						lngTemp4=lngTemp4_1+5
					end if
				end if
				
				if lngTemp4>BeginPoint and lngTemp4<=BeginPoint+lngBound then
					BeginPoint=lngTemp4
				else					
					lngTemp5_1=instr(BeginPoint,strContent,"<img",1)
					lngTemp5_2=instr(BeginPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2
					else
						lngTemp5=BeginPoint
					end if
					
					if lngTemp5>BeginPoint and lngTemp5<BeginPoint+lngBound then
						BeginPoint=lngTemp5
					else
						lngTemp6_1=instr(BeginPoint,strContent,"<br>",1)
						lngTemp6_2=instr(BeginPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2
						else
							lngTemp6=0
						end if
					
						if lngTemp6>BeginPoint and lngTemp6<BeginPoint+lngBound then
							BeginPoint=lngTemp6+4
						end if
					end if
				end if
			end if
		end if

		if CurrentPage=pages then
			EndPoint=ContentLen
		else
		  EndPoint=MaxPerPage*CurrentPage
		  if EndPoint>=ContentLen then
			EndPoint=ContentLen
		  else
			lngTemp1_1_1=instr(EndPoint,strContent,"</table>",1)
			lngTemp1_1_2=instr(EndPoint,strContent,"</TABLE>",1)
			lngTemp1_1_3=instr(EndPoint,strContent,"</Table>",1)
			if lngTemp1_1_1>0 then
				lngTemp1_1=lngTemp1_1_1
			elseif lngTemp1_1_2>0 then
				lngTemp1_1=lngTemp1_1_2
			elseif lngTemp1_1_3>0 then
				lngTemp1_1=lngTemp1_1_3
			else
				lngTemp1_1=0
			end if
							
			lngTemp1_2_1=instr(EndPoint,strContent,"<table",1)
			lngTemp1_2_2=instr(EndPoint,strContent,"<TABLE",1)
			lngTemp1_2_3=instr(EndPoint,strContent,"<Table",1)
			if lngTemp1_2_1>0 then
				lngTemp1_2=lngTemp1_2_1
			elseif lngTemp1_2_2>0 then
				lngTemp1_2=lngTemp1_2_2
			elseif lngTemp1_2_3>0 then
				lngTemp1_2=lngTemp1_2_3
			else
				lngTemp1_2=0
			end if
			
			if lngTemp1_1=0 and lngTemp1_2=0 then
				lngTemp1=EndPoint
			else
				if lngTemp1_1>lngTemp1_2 then
					lngtemp1=lngTemp1_2-1
				else
					lngTemp1=lngTemp1_1+7
				end if
			end if

			lngTemp2_1_1=instr(EndPoint,strContent,"</p>",1)
			lngTemp2_1_2=instr(EndPoint,strContent,"</P>",1)
			if lngTemp2_1_1>0 then
				lngTemp2_1=lngTemp2_1_1
			elseif lngTemp2_1_2>0 then
				lngTemp2_1=lngTemp2_1_2
			else
				lngTemp2_1=0
			end if
						
			lngTemp2_2_1=instr(EndPoint,strContent,"<p",1)
			lngTemp2_2_2=instr(EndPoint,strContent,"<P",1)
			if lngTemp2_2_1>0 then
				lngTemp2_2=lngTemp2_2_1
			elseif lngTemp2_2_2>0 then
				lngTemp2_2=lngTemp2_2_2
			else
				lngTemp2_2=0
			end if
			
			if lngTemp2_1=0 and lngTemp2_2=0 then
				lngTemp2=EndPoint
			else
				if lngTemp2_1>lngTemp2_2 then
					lngTemp2=lngTemp2_2-1
				else
					lngTemp2=lngTemp2_1+3
				end if
			end if

			lngTemp3_1_1=instr(EndPoint,strContent,"</ur>",1)
			lngTemp3_1_2=instr(EndPoint,strContent,"</UR>",1)
			if lngTemp3_1_1>0 then
				lngTemp3_1=lngTemp3_1_1
			elseif lngTemp3_1_2>0 then
				lngTemp3_1=lngTemp3_1_2
			else
				lngTemp3_1=0
			end if
			
			lngTemp3_2_1=instr(EndPoint,strContent,"<ur",1)
			lngTemp3_2_2=instr(EndPoint,strContent,"<UR",1)
			if lngTemp3_2_1>0 then
				lngTemp3_2=lngTemp3_2_1
			elseif lngTemp3_2_2>0 then
				lngTemp3_2=lngTemp3_2_2
			else
				lngTemp3_2=0
			end if
					
			if lngTemp3_1=0 and lngTemp3_2=0 then
				lngTemp3=EndPoint
			else
				if lngTemp3_1>lngTemp3_2 then
					lngtemp3=lngTemp3_2-1
				else
					lngTemp3=lngTemp3_1+4
				end if
			end if
			
			if lngTemp1<lngTemp2 then
				lngTemp=lngTemp2
			else
				lngTemp=lngTemp1
			end if
			if lngTemp<lngTemp3 then
				lngTemp=lngTemp3
			end if

			if lngTemp>EndPoint and lngTemp<=EndPoint+lngBound then
				EndPoint=lngTemp
			else
				lngTemp4_1_1=instr(EndPoint,strContent,"</li>",1)
				lngTemp4_1_2=instr(EndPoint,strContent,"</LI>",1)
				if lngTemp4_1_1>0 then
					lngTemp4_1=lngTemp4_1_1
				elseif lngTemp4_1_2>0 then
					lngTemp4_1=lngTemp4_1_2
				else
					lngTemp4_1=0
				end if
				
				lngTemp4_2_1=instr(EndPoint,strContent,"<li",1)
				lngTemp4_2_1=instr(EndPoint,strContent,"<LI",1)
				if lngTemp4_2_1>0 then
					lngTemp4_2=lngTemp4_2_1
				elseif lngTemp4_2_2>0 then
					lngTemp4_2=lngTemp4_2_2
				else
					lngTemp4_2=0
				end if
				
				if lngTemp4_1=0 and lngTemp4_2=0 then
					lngTemp4=EndPoint
				else
					if lngTemp4_1>lngTemp4_2 then
						lngtemp4=lngTemp4_2-1
					else
						lngTemp4=lngTemp4_1+4
					end if
				end if
				
				if lngTemp4>EndPoint and lngTemp4<=EndPoint+lngBound then
					EndPoint=lngTemp4
				else					
					lngTemp5_1=instr(EndPoint,strContent,"<img",1)
					lngTemp5_2=instr(EndPoint,strContent,"<IMG",1)
					if lngTemp5_1>0 then
						lngTemp5=lngTemp5_1-1
					elseif lngTemp5_2>0 then
						lngTemp5=lngTemp5_2-1
					else
						lngTemp5=EndPoint
					end if
					
					if lngTemp5>EndPoint and lngTemp5<EndPoint+lngBound then
						EndPoint=lngTemp5
					else
						lngTemp6_1=instr(EndPoint,strContent,"<br>",1)
						lngTemp6_2=instr(EndPoint,strContent,"<BR>",1)
						if lngTemp6_1>0 then
							lngTemp6=lngTemp6_1+3
						elseif lngTemp6_2>0 then
							lngTemp6=lngTemp6_2+3
						else
							lngTemp6=EndPoint
						end if
					
						if lngTemp6>EndPoint and lngTemp6<EndPoint+lngBound then
							EndPoint=lngTemp6
						end if
					end if
				end if
			end if
		  end if
		end if
		
		if EndPoint < BeginPoint then
			'BeginPoint = BeginPoint + str4
			'EndPoint = BeginPoint + str4
		end if

		On Error Resume Next
		AutoPagination_Tmp = AutoPagination_Tmp & mid(strContent,BeginPoint,EndPoint-BeginPoint)
		
		If Err Then
			Err.clear
			'response.Write "BeginPoint = "& BeginPoint
			'response.Write "<br>"
			'response.Write "EndPoint = "& EndPoint
			AutoPagination_Tmp = AutoPagination_Tmp & "</p><div id=""clear""></div><p align=center style='color:red;'>对不起，自动分页错误，请直接点下一页即可接上页继续。</p>"
		End If

		
		AutoPagination_Tmp = AutoPagination_Tmp & "</p><div id=""clear""></div><div id=""vmoviesab""><table border=""0""  cellspacing=""5"" cellpadding=""2"" align=""center""><tr>" & VbCrLf
		if CurrentPage>1 then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55""  class=""page_css_1_1""><a href='Article_" & ArticleId & "_" & CurrentPage-1 
			AutoPagination_Tmp = AutoPagination_Tmp & ".html'>上一页</a></TD>" & VbCrLf
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55""  class=""page_css_1_1"">上一页</TD>" & VbCrLf
		end if
		for i=1 to pages
			if i=CurrentPage then
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1"">" & cstr(i) & "</TD>" & VbCrLf
			else
				AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""25"" class=""page_css_2_1""><a href='Article_" & ArticleId & "_" & i 
				AutoPagination_Tmp = AutoPagination_Tmp & ".html'>" & i & "</a></TD>" & VbCrLf
			end if
			'if (i Mod 12) = 0 then AutoPagination_Tmp = AutoPagination_Tmp & "</ul><ul>"
		next
		if CurrentPage<pages then
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55""  class=""page_css_1_1""><a href='Article_" & ArticleId & "_" & CurrentPage+1 
			AutoPagination_Tmp = AutoPagination_Tmp & ".html'>下一页</a></TD>" & VbCrLf
		else
			AutoPagination_Tmp = AutoPagination_Tmp & "<TD width=""55"" class=""page_css_1_1"">下一页</TD>" & VbCrLf
		end if
		AutoPagination_Tmp = AutoPagination_Tmp & "</tr></table></div>"& VbCrLf
	end if
	AutoPagination2 = AutoPagination_Tmp
end Function

'=================================================
'过程名：BbbImg
'作  用：鼠标滚轮控制图片大小的函数
'参  数：strText
'=================================================
Function BbbImg(strText)
         Dim s,re
         Set re=New RegExp
         re.IgnoreCase = true
         re.Global = true		 
         s=strText
		 
		'去掉图片中的脚本代码
		re.Pattern="<IMG.[^>]*SRC(=| )(.[^>]*)>"
		s=re.replace(s,"<IMG SRC=$2 onload=""javascript:resizeimg(this,600,450)"">")
		
		 BbbImg = ChkBadWords(s)
	     Set re=Nothing
End Function

'脏话过滤
Function ChkBadWords(Str)
		If IsNull(Str) Then Exit Function
		Dim i,rBadWord,BadWord
		BadWord	=""&BadWord1&""
		BadWord = Split(BadWord,"|")
		For i = 0 To Ubound(BadWord)
			rBadWord = Split(BadWord(i),"=")
			Str = Replace(Str,rBadWord(0),rBadWord(1))
		Next
		ChkBadWords = Str
End Function




Function xiaowei(block)
	if not isnull(block) then
    block = Replace(block, "{$SitePath}", SitePath)
    xiaowei = block
	end if
End Function

'# IIF
Function IIF(A,B,C)
	If A Then IIF = B Else IIF = C
End Function

'搜索蜘蛛
function spiderbot()
	dim agent
	agent = lcase(request.servervariables("http_user_agent"))
	dim Bot: Bot = ""	
	if instr(agent, "googlebot") > 0 then Bot = "Google"
	if instr(agent, "baiduspider") > 0 then Bot = "Baidu"
	if instr(agent, "sogou") > 0 then Bot = "Sogou"
	if instr(agent, "yahoo") > 0 then Bot = "Yahoo!"
	if instr(agent, "msn") > 0 then Bot = "MSN"
	if instr(agent, "ia_archiver") > 0 then Bot = "Alexa"
	if instr(agent, "iaarchiver") > 0 then Bot = "Alexa"
	if instr(agent, "sqworm") > 0 then Bot = "AOL"
	if instr(agent, "yodaobot") > 0 then Bot = "Yodao"
	if instr(agent, "iaskspider") > 0 then Bot = "Iask"
	
	if len(Bot) > 0 then
		set rs = server.CreateObject ("adodb.recordset")
		sql="select [Botname],[LastDate] From [xiaowei_Bots] Where [Botname]='" & Bot & "'"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
		rs.AddNew 
		rs(0) = Bot
		rs(1) = now()
		else
		rs(1) = now()
		end if
		rs.update
		rs.close: set rs = nothing
	end if
end function



'显示相关文章
'P_ConID:数值型，当前文章ID
'P_Key:字符型，当前文章关健字
'P_Row:数值型，要显示相关文章的条数
'P_ICO:字符型，标题前图标，可以图片也可为字符
'P_Time:数值型，显示时间，０为不显示，否则为时间格式

Function ShowMutualityArticle(P_ConID,P_Key,P_Row,P_ICO,P_Time)
    dim pRs,pSql
	dim i,TempKeyWord
	
	if P_Row > 0 then
		pSql = "Select TOP "& P_Row
	else
		pSql = "Select "
	end if
	pSql = pSql & " ID,Title,ClassId,DateAndTime From [xiaowei_Article] Where ID <> "& P_ConID &" And "	
	if Instr(P_Key,"|") > 0 then
		P_Key = Split(P_Key,"|")
		TempKeyWord = TempKeyWord &"("
		For i = 0 to Ubound(P_Key)
			TempKeyWord = TempKeyWord &" KeyWord like '%"& P_Key(i) &"%' or KeyWord like '%|"& P_Key(i) &"|%' or KeyWord like '%"& P_Key(i) &"|%' or KeyWord like '%|"& P_Key(i) &"%' "
			if i = Ubound(P_Key) then
				TempKeyWord = TempKeyWord &") And "
			else
				TempKeyWord = TempKeyWord &" Or "
			end if
		Next
	else
		TempKeyWord = TempKeyWord &" KeyWord like '%"& P_Key &"%' And "
	end if
	pSql = pSql & TempKeyWord &" yn = 0 Order By Id Desc"
	'Response.Write pSql
	

    Set pRs = Server.CreateObject("Adodb.recordset")
	pRs.open pSql,conn,1,3
	if not(pRs.bof and pRs.eof) then
		Do While Not pRs.eof
			if pRs(0) <> P_ConID Then

		ShowMutualityArticle = ShowMutualityArticle & "<ul>" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "<li style=""float:left;text-align:left;width:100%;line-height:20px;padding:5px 0 0 0;"">" & VbCrLf
		
				ShowMutualityArticle = ShowMutualityArticle & "<a href="""&SitePath&"page/?"&pRs(0)&".html"">"&left(LoseHtml(pRs(1)),18)&"</a>&nbsp;&nbsp;&nbsp;&nbsp; <font color=""#cccccc"">"&FormatDate(pRs(3),5)&"</font>"
		 ShowMutualityArticle = ShowMutualityArticle & "<p><font color=""#Eeeeee"">-------------------------------------------------</font></p>" & VbCrLf
        ShowMutualityArticle = ShowMutualityArticle & "</li>" & VbCrLf
        ShowMutualityArticle = ShowMutualityArticle & "</ul>" & VbCrLf
 
			end if
			pRs.movenext
		Loop
	else
		ShowMutualityArticle = ShowMutualityArticle & "<div class=""vmovies"">" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "<div class=""movernrs"">" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "<ul>" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "<li class=""titles"">" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "没有文章" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "</li>" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "</ul>" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "</div>" & VbCrLf
		ShowMutualityArticle = ShowMutualityArticle & "</div>" & VbCrLf
	end if
	pRs.close:set pRs = nothing

End function












'文章调用_用户发表
'ClassID:数值型，栏目ID
'N:数值型，要显示文章条数
'T:数值型，显示时间，０为不显示，否则为时间格式
'ICO:字符型，标题前图标，可以图片也可为字符
'Z:标题字数
'msql:增强条件
'P:排序方式

Sub ShowuserArticle(N,T,ICO,Z,msql,P)
	set rs1=server.createobject("ADODB.Recordset")
	SQL1="select Top "&N&" ID,Title,ClassID,DateAndTime,username,TitleFontColor,IsHot from xiaowei_Article where username <> ''"
	

	
	If msql<>"no" then
			SQL1=SQL1&" and "&msql&""
	End if
	
	SQL1=SQL1&" Order by "&P&""
	
	rs1.open sql1,conn,1,3
	If Not rs1.Eof Then 
	do while not (rs1.eof or err)
	

	Response.Write("<p style=""float:left;text-align:left;width:100%;line-height:22px;padding:5px 0 3px 0;"">"&ICO&"")
	
	

		Response.Write("<a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank"">")

Response.Write(""&left(rs1("Title"),Z)&"</a>")
Response.Write("-----------")
		Response.Write("<a href=""../xwuser/showuser.asp?username="&rs1("username")&""" target=""_blank"">")
		Response.Write(""&rs1("username")&"</a>")

	If T<>0 then
		Response.Write(" <span style=""color:#ccc;"">["&FormatDate(rs1("DateAndTime"),5)&"]</span>")
	end if
		Response.Write("</p><p align=""left""><font color=""#Eeeeee"">----------------------------------------------------------------------------</font></p>") & VbCrLf



	rs1.movenext
	loop

	end if
	rs1.close
	set rs1=nothing
End Sub

'文章调用
'ClassID:数值型，栏目ID
'N:数值型，要显示文章条数
'T:数值型，显示时间，０为不显示，否则为时间格式
'ICO:字符型，标题前图标，可以图片也可为字符
'Z:标题字数
'msql:增强条件
'P:排序方式

Sub ShowArticle(ClassID,N,T,ICO,Z,msql,P)
	set rs1=server.createobject("ADODB.Recordset")
	SQL1="select Top "&N&" ID,Title,ClassID,DateAndTime,zffy,IsHot from xiaowei_Article where yn = 0"
	
	If ClassID<>0 then
		If xiaowei_MyID(ClassID)="0" then
			SQL1=SQL1&" and ClassID="&ClassID&""
		else
			MyID = Replace(""&xiaowei_MyID(ClassID)&"","|",",")
			SQL1=SQL1&" and ClassID in ("&MyID&")"
		End if
	End if
	
	If msql<>"no" then
			SQL1=SQL1&" and "&msql&""
	End if
	
	SQL1=SQL1&" Order by "&P&""
	
	rs1.open sql1,conn,1,3
	If Not rs1.Eof Then 
	do while not (rs1.eof or err)
	
	Response.Write("<ul>")
	Response.Write("<li style=""float:left;text-align:left;width:50%;line-height:35px;padding:0 0 5px 0;"">"&ICO&"")
	
	

		Response.Write("<a href="""&SitePath&"page/?"&rs1("ID")&".html"" target=""_blank"">")

		

	Response.Write(""&left(rs1("Title"),Z)&"</a>")
		If rs1("zffy")>0 then
	Response.Write("<span style=""background:#FF6600;color:#ffffff;padding:0px 4px;"">费</span>")
	end if
	If T<>0 then
	
      If ""&DateDiff("d",""&FormatDate(rs1("DateAndTime"),5)&"",date())&""=0 then 
           Response.Write("<font color=FF0000> [今天]</FONT>")         
               else
                If ""&DateDiff("d",""&FormatDate(rs1("DateAndTime"),5)&"",date())&""=1 then 
                 Response.Write("<font color=cccccc> [1天前]</FONT>") 
                 else
                              If ""&DateDiff("d",""&FormatDate(rs1("DateAndTime"),5)&"",date())&""=2 then 
                 Response.Write("<font color=cccccc> [2天前]</FONT>") 
                 else
		Response.Write(" <span style=""color:#ccc;"">["&FormatDate(rs1("DateAndTime"),5)&"]</span>")
		end if
		end if
		end if
	end if
	Response.Write("</li>") & VbCrLf
	Response.Write("</ul>") & VbCrLf

	rs1.movenext
	loop
	else
	Response.Write("<div class=""vmovies"">")
	Response.Write("<div class=""movernrs"">")
	Response.Write("<ul>")
	Response.Write("<li class=""titles"">"&ICO&"")
	Response.Write("没有")
	Response.Write("</li>") & VbCrLf
	Response.Write("</ul>") & VbCrLf
	Response.Write("</div>") & VbCrLf
	Response.Write("</div>") & VbCrLf
	end if
	rs1.close
	set rs1=nothing
End Sub


'SORT 2图片方式显示
'ClassID:数值型，栏目ID
'N:数值型，要显示文章条数
'T:数值型，显示时间，０为不显示，否则为时间格式
'ICO:字符型，标题前图标，可以图片也可为字符
'Z:标题字数
'msql:增强条件
'P:排序方式

Sub ShowimgsortArticle(ClassID,N,T,ICO,Z,msql,P)
	set rs1=server.createobject("ADODB.Recordset")
	SQL1="select Top "&N&" * from xiaowei_Article where yn = 0 and images <> ''"
	
	If ClassID<>0 then
		If xiaowei_MyID(ClassID)="0" then
			SQL1=SQL1&" and ClassID="&ClassID&""
		else
			MyID = Replace(""&xiaowei_MyID(ClassID)&"","|",",")
			SQL1=SQL1&" and ClassID in ("&MyID&")"
		End if
	End if
	
	If msql<>"no" then
			SQL1=SQL1&" and "&msql&""
	End if
	
	SQL1=SQL1&" Order by "&P&""
	
	rs1.open sql1,conn,1,3
	If Not rs1.Eof Then 
	do while not (rs1.eof or err)
	
	Response.Write("<ul>")
	Response.Write("<li style=""float:left;text-align:left;width:20%;line-height:35px;padding:0 0 5px 0;""><table width=""100%"" cellspacing=""0"" cellpadding=""0"" id=""table1"" style=""border: 2px solid #FFFFFF; padding: 15px"">")
	
	 if left(LoseHtml(rs1("images")),4) <> "http" then 

		Response.Write("<tr><td width=""100%""><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><img src="".."&rs1("Images")&""" alt="""" width=""170"" height=""117""/></a></td></tr>")
		else
			Response.Write("<tr><td width=""100%""><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><img src="""&rs1("Images")&""" alt="""" width=""170"" height=""117""/></a></td></tr>")
			end if 
		Response.Write("<tr><td width=""100%""><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><span><b><font size=""2"">"&left(LoseHtml(rs1("title")),18)&"</font></b></span></a> &nbsp;&nbsp;[<font color=""#ff0000"">"&rs1("hits")&"</font>] ")
				If rs1("zffy")>0 then
	Response.Write("<span style=""background:#FF6600;color:#ffffff;padding:0px 4px;"">费</span>")
	end if
	Response.Write("</td> ")
		Response.Write("</tr></table> ")





	Response.Write("</li>") & VbCrLf
	Response.Write("</ul>") & VbCrLf

	rs1.movenext
	loop

	end if
	rs1.close
	set rs1=nothing
End Sub

'图片文章调用
'ClassID:数值型，栏目ID
'N:数值型，要显示文章条数
'T:数值型，显示时间，０为不显示，否则为时间格式
'ICO:字符型，标题前图标，可以图片也可为字符
'Z:标题字数
'msql:增强条件
'P:排序方式

Sub ShowimgArticle(ClassID,N,T,ICO,Z,msql,P)
	set rs1=server.createobject("ADODB.Recordset")
	SQL1="select Top "&N&" * from xiaowei_Article where yn = 0 and images <> ''"
	
	If ClassID<>0 then
		If xiaowei_MyID(ClassID)="0" then
			SQL1=SQL1&" and ClassID="&ClassID&""
		else
			MyID = Replace(""&xiaowei_MyID(ClassID)&"","|",",")
			SQL1=SQL1&" and ClassID in ("&MyID&")"
		End if
	End if
	
	If msql<>"no" then
			SQL1=SQL1&" and "&msql&""
	End if
	
	SQL1=SQL1&" Order by "&P&""
	
	rs1.open sql1,conn,1,3
	If Not rs1.Eof Then 
	do while not (rs1.eof or err)
	
	Response.Write("<ul>")
	Response.Write("<li style=""float:left;text-align:left;width:100%;line-height:35px;padding:0 0 5px 0;""><table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" id=""table1"" height=""60"">")
	
	 if left(LoseHtml(rs1("images")),4) <> "http" then 
		Response.Write("<tr><td width=""80""><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><img src="".."&rs1("Images")&""" alt="""" width=""73"" height=""50""/></a></td> ")
		else
		Response.Write("<tr><td width=""80""><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><img src="""&rs1("Images")&""" alt="""" width=""73"" height=""50""/></a></td> ")
		end if
		Response.Write("<td><a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""><span><b><font size=""2"">"&left(LoseHtml(rs1("title")),18)&"</font></b></span></a> ")
				If rs1("zffy")>0 then
	Response.Write("<span style=""background:#FF6600;color:#ffffff;padding:0px 4px;"">费</span>")
	end if
	Response.Write("</td> ")
		Response.Write("<td  width=""30"" align=""right""><font color=""#0099CC"">"&rs1("hits")&"</font></td></tr></table> ")

		Response.Write("<p style=""line-height: 180%"" align=""left"">"&left(LoseHtml(rs1("Content")),40)&"…&nbsp;&nbsp;<a href=""../page/index.asp?"&rs1("ID")&""" target=""_blank""> <font color=""#cccccc"">[详情]</font></a> ")

Response.Write("<p><font color=""#Eeeeee"">-----------------------------------------------------</font></p>")


	Response.Write("</li>") & VbCrLf
	Response.Write("</ul>") & VbCrLf

	rs1.movenext
	loop

	end if
	rs1.close
	set rs1=nothing
End Sub

Function xiaowei_MyID(a)
xiaowei_MyID=""
 Dim rs1,sql1
 set rs1=server.createobject("ADODB.Recordset")
 sql1="select * from xiaowei_Class where TopID = "&a&""
 rs1.open sql1,conn,1,3
 If Not rs1.Eof Then 
 do while not (rs1.eof or err)

If xiaowei_MyID = "" then
	xiaowei_MyID = rs1("ID")
else
	xiaowei_MyID = xiaowei_MyID &"|"& rs1("ID")
End if
 rs1.movenext
 loop
 else
xiaowei_MyID = "0"
 end if
 rs1.close
 set rs1=nothing
End Function

Function Classlist(id)
	if id = "" or isnull(id) then
		Classlist = ""
	else
		Sqld = "Select ClassName from xiaowei_Class where ID = " & id
		Set rsd = conn.execute(Sqld)
		if not rsd.eof then
			Classlist = rsd(0)
		else
			Classlist = ""
		end if
		rsd.close
	end if
End function

function checkpost(byval back)
	dim server_v1, server_v2
	server_v1 = cstr(request.servervariables("http_referer"))
	server_v2 = cstr(request.servervariables("server_name"))
	if Mid(server_v1, 8, len(server_v2)) <> server_v2 then
		if not back then
			response.write lang_errorpost : response.end
		else
			checkpost = false
		end if
	else
		checkpost = true
	end if
end function

Function Alert(message,gourl) 
    message = replace(message,"'","\'")
    If gourl="-1" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-1)</script>")
    Else
        Response.Write ("<script language=javascript>alert('" & message & "');location='" & gourl &"'</script>")
    End If
    Response.End()
End Function

Function UserGroup(id)
	if id = "" or isnull(id) then
		UserGroup = ""
	else
		Sqld = "Select dengji from xiaowei_User where ID = " & id
		Set rsd = conn.execute(Sqld)
		if not rsd.eof then
			UserGroup = rsd(0)
		else
			UserGroup = ""
		end if
		rsd.close
	end if
End function

'过滤指定html标签

Function lFilterBadHTML(byval strHTML,byval strTAGs)  
  Dim objRegExp,strOutput  
  Dim arrTAG,i
  arrTAG=Split(strTAGs,",")  
  Set objRegExp = New Regexp   
  strOutput=strHTML   
  objRegExp.IgnoreCase = True  
  objRegExp.Global = True  
  For i=0 to UBound(arrTAG)  
    objRegExp.Pattern = "<"&arrTAG(i)&"[\s\S]+</"&arrTAG(i)&"*>"  
    strOutput = objRegExp.Replace(strOutput, "")   
  Next  
  Set objRegExp = Nothing  
  lFilterBadHTML = strOutput   
End Function 



%>