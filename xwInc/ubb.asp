<%
function UBB_MP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		strContent=re.replace(strContent, chr(1) & "MP=$1,$2" & chr(2))
		re.Pattern="\[\/MP\]"
		Test=re.Test(strContent)
		if Test then
			strContent=re.replace(strContent, chr(1) & "/MP" & chr(2))
				re.Pattern="\x01MP=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/MP\x02"
				strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3 width=$1 height=$2></embed></object>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
		end if
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	set re=Nothing
	UBB_MP=strContent
end function

function UBB_RM(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		strContent=re.replace(strContent, chr(1) & "RM=$1,$2" & chr(2))
		re.Pattern="\[\/RM\]"
		Test=re.Test(strContent)
		if Test then
			strContent=re.replace(strContent, chr(1) & "/RM" & chr(2))
				re.Pattern="\x01RM=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/RM\x02"
				strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
		end if
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	set re=Nothing
	UBB_RM=strContent
end function

function UBB_FLASH(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[FLASH\]"
	Test=re.Test(strContent)
	if Test then
		strContent=re.replace(strContent, chr(1) & "FLASH" & chr(2))
		re.Pattern="\[\/FLASH\]"
		Test=re.Test(strContent)
		if Test then
			strContent=re.replace(strContent, chr(1) & "/FLASH" & chr(2))
				re.Pattern="\x01FLASH\x02(.[^\x01]*)\x01\/FLASH\x02"
				strContent=re.Replace(strContent,"<a href=""$1"" TARGET=_blank><IMG SRC=/images/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$1""><PARAM NAME=quality VALUE=high><embed src=""$1"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$1</embed></OBJECT>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
		end if
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	re.Pattern="\[FLASH=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		strContent=re.replace(strContent, chr(1) & "FLASH=$1,$2" & chr(2))
		re.Pattern="\[\/FLASH\]"
		Test=re.Test(strContent)
		if Test then
			strContent=re.replace(strContent, chr(1) & "/FLASH" & chr(2))
				re.Pattern="\x01FLASH=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/FLASH\x02"
				strContent=re.Replace(strContent,"<a href=""$3"" TARGET=_blank><IMG SRC=/images/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$1 height=$2>$3</embed></OBJECT>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
		end if
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	set re=Nothing
	UBB_FLASH=strContent
end function

'参数：strContent内容
function UBBCode(strContent)

'UbbCode
dim re,ii,po
dim reContent,Test
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

strContent=UBB_FLASH(strContent)
strContent=UBB_MP(strContent)
strContent=Replace(strContent,"http://hiphotos.baidu.com","http://%77%77%77%2e%6c%61%6f%79%38%2e%63%6e/photo.asp?url=http://hiphotos.baidu.com")

	set objRegExp=Nothing
	UBBCode=strContent

If InStr(LCase(strContent),"[code]")>0 Then
			'strContent = Replace(strContent,"<br />",vbLf)
			strContent = Replace(strContent,"<","&lt;")
			strContent = Replace(strContent,">","&gt;")
			'strContent = Replace(strContent,"<p>","")
			'strContent = Replace(strContent,"</p>",vbNewLine)
			strContent = Replace(strContent,"&nbsp;",Chr(9))
			'strContent = Replace(strContent,"&nbsp;","")
			strContent = Replace(strContent,vbLf,"")
			re.Pattern = "\[code\](.*?)\[\/code\]"
			Set strMatchs = re.Execute(strContent)
			For Each strMatch In strMatchs
				Randomize
				CodeNum = CStr(Int(7999 * Rnd + 2000))
				strContent = Replace(strContent,strMatch.Value,"<li>HTML代码</li><div style=""width:500px;float:left;""><textarea name=""runcode0"" rows=""15"" style='font-family:Courier New,Courier,monospace;width:500px;font-size:12px;margin-bottom:5px;'>"&strMatch.SubMatches(0)& "</textarea><br/><input type=""button"" value=""运行代码"" class=""borderall1"" onclick=""runCode(runcode0)""> <input type=""button"" class=""borderall1"" value=""复制代码"" onclick=""copycode(runcode0)""> <input type=""button"" class=""borderall1"" value=""另存代码"" onclick=""saveCode(runcode0)""> &nbsp;提示：您可以先修改部分代码再运行</div>")
			Next
			Set strMatchs = Nothing
			strContent = Replace(strContent,vbCr,vbCrLf)
			'strContent = Replace(strContent,Chr(8)&Chr(11)&Chr(9)&Chr(12),vbCr)
End If

set re=Nothing
UBBCode=BbbImg(strContent)
end function

'——脚本字符处理	
Function JScode(JSstr)
if not isnull(JSstr) then
dim ts
dim re
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="(javascript)"
ts=re.Replace(JSstr,"&#106avascript")
re.Pattern="(jscript:)"
ts=re.Replace(ts,"&#106script:")
re.Pattern="(js:)"
ts=re.Replace(ts,"&#106s:")
re.Pattern="(value)"
ts=re.Replace(ts,"&#118alue")
re.Pattern="(about:)"
ts=re.Replace(ts,"about&#58")
re.Pattern="(file:)"
ts=re.Replace(ts,"file&#58")
re.Pattern="(document.cookie)"
ts=re.Replace(ts,"documents&#46cookie")
re.Pattern="(vbscript:)"
ts=re.Replace(ts,"&#118bscript:")
re.Pattern="(vbs:)"
ts=re.Replace(ts,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
ts=re.Replace(ts,"&#111n$2")
re.Pattern="(&#)"
ts=re.Replace(ts,"＆#")
JScode=ts
set re=nothing
end if
End Function
%>