<!--#include file="../xwinc/config.asp"-->
<!--#include file="Admin_check.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Abiao CMS ϵͳ����</title>
<link href="Images/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<script src="js/admin.js"></script>
<body topmargin="0" leftmargin="0">


<!--#include file="top.asp"-->


<table border="0" width="100%" cellspacing="0" cellpadding="0" height="126" id="table1">
	<tr>
		<td width="200"  valign="top"><!--#include file="left.asp"--></td>
		<td width="1" bgcolor="#006699">��</td>
		<td valign="top"><br>	
			
<%
Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ�������ʹ��FTP�ȹ��ܣ���<font color=Red >inc/config.asp</font>�ļ������滻�ɿ�������"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub

If request("action")="Edit" then
SiteTitle = replace(Trim(Request.Form("SiteTitle")),CHR(34),"'")
SiteUrl = replace(Trim(Request.Form("SiteUrl")),CHR(34),"'")
SitePath = replace(Trim(Request.Form("SitePath")),CHR(34),"'")
DataName = replace(Trim(Request.Form("DataName")),CHR(34),"'")

SiteUp = replace(Trim(Request.Form("SiteUp")),CHR(34),"'")
Sitekeywords = replace(Trim(Request.Form("Sitekeywords")),CHR(34),"'")
Sitedescription = replace(Trim(Request.Form("Sitedescription")),CHR(34),"'")
IsPing = replace(Trim(Request.Form("IsPing")),CHR(34),"'")
pingoff = replace(Trim(Request.Form("pingoff")),CHR(34),"'")
fenlei1 = replace(Trim(Request.Form("fenlei1")),CHR(34),"'")
fenlei2 = replace(Trim(Request.Form("fenlei2")),CHR(34),"'")
fenlei3 = replace(Trim(Request.Form("fenlei3")),CHR(34),"'")
fenlei4 = replace(Trim(Request.Form("fenlei4")),CHR(34),"'")
fenlei5 = replace(Trim(Request.Form("fenlei5")),CHR(34),"'")

SiteTcp = replace(Trim(Request.Form("SiteTcp")),CHR(34),"'")
BadWord1 = replace(Trim(Request.Form("BadWord1")),CHR(34),"'")
FontSize = replace(Trim(Request.Form("FontSize")),CHR(34),"'")
FontFamily = replace(Trim(Request.Form("FontFamily")),CHR(34),"'")
Fonttext = replace(Trim(Request.Form("Fonttext")),CHR(34),"'")
aspjpeg = replace(Trim(Request.Form("aspjpeg")),CHR(34),"'")
Color1 = replace(Trim(Request.Form("Color1")),"#","")
Color2 = replace(Trim(Request.Form("Color2")),"#","")

ad1 = replace(Trim(Request.Form("ad1")),CHR(34),"'")
ad2 = replace(Trim(Request.Form("ad2")),CHR(34),"'")
ad3 = replace(Trim(Request.Form("ad3")),CHR(34),"'")
ad4 = replace(Trim(Request.Form("ad4")),CHR(34),"'")
ad5 = replace(Trim(Request.Form("ad5")),CHR(34),"'")
ad6 = replace(Trim(Request.Form("ad6")),CHR(34),"'")
ad7 = replace(Trim(Request.Form("ad7")),CHR(34),"'")


artlistnum = Request.Form("artlistnum")


linkoff = replace(Trim(Request.Form("linkoff")),CHR(34),"'")
tougaooff = replace(Trim(Request.Form("tougaooff")),CHR(34),"'")
userynoff = replace(Trim(Request.Form("userynoff")),CHR(34),"'")
useraddoff = replace(Trim(Request.Form("useraddoff")),CHR(34),"'")
userWord = replace(Trim(Request.Form("userWord")),CHR(34),"'")
useroff = replace(Trim(Request.Form("useroff")),CHR(34),"'")
money1 = replace(Trim(Request.Form("money1")),CHR(34),"'")
money2 = replace(Trim(Request.Form("money2")),CHR(34),"'")
money3 = replace(Trim(Request.Form("money3")),CHR(34),"'")
money4 = replace(Trim(Request.Form("money4")),CHR(34),"'")
money5 = replace(Trim(Request.Form("money5")),CHR(34),"'")
moneyname = replace(Trim(Request.Form("moneyname")),CHR(34),"'")
yaopostgetime = replace(Trim(Request.Form("yaopostgetime")),CHR(34),"'")

Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim SiteTitle,SiteUrl,SitePath,DataName,skin,SiteUp,fenlei1,fenlei2,fenlei3,fenlei4,fenlei5,Sitekeywords,Sitedescription,SiteAdmin,Htmledit,IsPing,isfagao,rss,Css,SiteTcp,Sitelx,BadWord1,FontSize,Aspjpeg,FontFamily,Fonttext,Color1,Color2,mood,menuimg,indeximg,ad1,ad2,ad3,ad4,ad5,ad6,ad7,seo,artlist,artlistnum,gsname,pingoff,bookoff,linkoff,tougaooff,userynoff,useraddoff,userWord,useroff,money1,money2,money3,money4,money5,moneyname,yaopostgetime" & VbCrLf & VbCrLf
	
	TempStr = TempStr & "'=====��վ����" & VbCrLf & VbCrLf
	TempStr = TempStr & "SiteTitle="& Chr(34) & SiteTitle & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "fenlei1="& Chr(34) & fenlei1 & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "fenlei2="& Chr(34) & fenlei2 & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "fenlei3="& Chr(34) & fenlei3 & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "fenlei4="& Chr(34) & fenlei4 & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "fenlei5="& Chr(34) & fenlei5 & Chr(34) &"" & VbCrLf & VbCrLf

	TempStr = TempStr & "'=====��վ����" & VbCrLf
	TempStr = TempStr & "'=====ע�ⲻҪ��д��ַǰ���http://�������/����www.zjc.com����" & VbCrLf & VbCrLf
	TempStr = TempStr & "SiteUrl="& Chr(34) & SiteUrl & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "'=====�����վĿ¼" & VbCrLf
	TempStr = TempStr & "'=====��Ŀ¼ֱ����/" & VbCrLf
	TempStr = TempStr & "'=====�磺SitePath="& Chr(34) & test2 & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "SitePath="& Chr(34) & SitePath & Chr(34) &"" & VbCrLf & VbCrLf
	TempStr = TempStr & "'==============================" & VbCrLf
	TempStr = TempStr & "DataName="& Chr(34) & DataName & Chr(34) &" '���ݿ�����" & VbCrLf

	TempStr = TempStr & "SiteUp="& Chr(34) & SiteUp & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "Sitekeywords="& Chr(34) & Sitekeywords & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "Sitedescription="& Chr(34) & Sitedescription & Chr(34) &"" & VbCrLf


	TempStr = TempStr & "SiteTcp="& Chr(34) & SiteTcp & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "BadWord1="& Chr(34) & BadWord1 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "'=====��ʾ����=====" & VbCrLf

	TempStr = TempStr & "IsPing="& Chr(34) & IsPing & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "pingoff="& Chr(34) & pingoff & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "ad1="& Chr(34) & ad1 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad2="& Chr(34) & ad2 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad3="& Chr(34) & ad3 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad4="& Chr(34) & ad4 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad5="& Chr(34) & ad5 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad6="& Chr(34) & ad6 & Chr(34) &"" & VbCrLf
    TempStr = TempStr & "ad7="& Chr(34) & ad7 & Chr(34) &"" & VbCrLf








	TempStr = TempStr & "artlistnum="& Chr(34) & artlistnum & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "linkoff="& Chr(34) & linkoff & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "tougaooff="& Chr(34) & tougaooff & Chr(34) &"" & VbCrLf

	TempStr = TempStr & "'=====�ϴ�ͼƬˮӡ=====" & VbCrLf
	TempStr = TempStr & "Aspjpeg="& Chr(34) & aspjpeg & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "FontSize="& Chr(34) & FontSize & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "FontFamily="& Chr(34) & FontFamily & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "Fonttext="& Chr(34) & Fonttext & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "Color1="& Chr(34) & Color1 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "Color2="& Chr(34) & Color2 & Chr(34) &"" & VbCrLf

	TempStr = TempStr & "'=====��Ա���=====" & VbCrLf
	TempStr = TempStr & "useroff=" & useroff & "" & VbCrLf
	TempStr = TempStr & "useraddoff="& Chr(34) & useraddoff & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "userynoff=" & userynoff & "" & VbCrLf
	TempStr = TempStr & "moneyname="& Chr(34) & moneyname & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "userWord="& Chr(34) & userWord & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "yaopostgetime=" & yaopostgetime & "" & VbCrLf
	TempStr = TempStr & "money1="& Chr(34) & money1 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "money2="& Chr(34) & money2 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "money3="& Chr(34) & money3 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "money4="& Chr(34) & money4 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "money5="& Chr(34) & money5 & Chr(34) &"" & VbCrLf
	TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../xwinc/Config.asp"
	If GBL_CHK_TempStr = "" Then
		Response.Write("<script language=javascript>alert('�޸ĳɹ���');this.location.href='admin_setting.asp';</script>")
	Else
		%><table width=""98%"" align=""center"" border=""1"" cellspacing=""0"" cellpadding=""4"" class=lanyubk style=""border-collapse: collapse""><tr><td class=lanyuss>�������ϸ���</td></tr><tr class=lanyuds><td align=""center"" height=""66"">&gt;<%=GBL_CHK_TempStr%>&lt;<br><br>
		<textarea name="fileContent" cols="1" rows="1"><%=Server.htmlencode(TempStr)%></textarea></td></tr></table><%
		GBL_CHK_TempStr = ""
	End If
End if
%><form action="?Action=Edit" method="post">
<table border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
  <tr>
    <td colspan="2" class="admintitle"> �޸���վ��������</td>
  </tr>
  <tr>
    <td width="20%" height="25" bgcolor="f7f7f7" class="tdleft">��վ���ƣ�</td>
    <td height="25" bgcolor="f7f7f7"><input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40"></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">��վ������</td>
    <td height="12" bgcolor="f7f7f7"><input name="SiteUrl"type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40"> <span class="note">�磺www.90wei.com,��Ҫ"http://"</span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">��װĿ¼��</td>
    <td height="-3" bgcolor="f7f7f7"><input name="SitePath"type="text" id="SitePath" value="<%=SitePath%>" size="40">
      <span class="note">��վ��װĿ¼����Ŀ¼����д&quot;/&quot;���ݲ�֧�ֶ���Ŀ¼��</span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">���ݿ����ƣ�</td>
    <td height="0" bgcolor="f7f7f7"><input name="DataName"type="text" id="DataName" value="<%=DataName%>" size="40">
      <span class="note">�����DataĿ¼�µ����ݿ����Ʋ��ڴ���д</span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">ͼƬĿ¼��</td>
    <td height="5" bgcolor="f7f7f7"><input name="SiteUp" type="hidden" id="SiteUp" value="<%=SiteUp%>" size="40">images
      <span class="note">�����޸�</span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">�� �� �֣�</td>
    <td height="25" bgcolor="f7f7f7"><input name="SiteKeywords" type="text" id="SiteKeywords" value="<%=SiteKeywords%>" size="40"> <span class="note">��վ�����������Ĺؼ���</span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">��վ������</td>
    <td bgcolor="f7f7f7"><input name="Sitedescription" type="text" id="Sitedescription" value="<%=Sitedescription%>" size="100"></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">�����ţ�</td>
    <td height="-3" bgcolor="f7f7f7"><input name="SiteTcp" type="text" id="SiteTcp" value="<%=SiteTcp%>" size="40"></td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">�໰���ˣ�</td>
    <td height="0" bgcolor="f7f7f7"><input name="BadWord1" type="text" id="BadWord1" value="<%=BadWord1%>" size="100">
      <br><span class="note">��ע���ʽ������ȷ�ĸ�ʽ���ܵ�����������ҳ�޷���ʾ,ÿ����˴���|����</span></td>
  </tr>
</table>





<table border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
  <tr>
    <td colspan="2" class="admintitle"> ��ϸ����</td>
  </tr>
    <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">�����Ƿ���ʾ���ۣ�</td>
    <td bgcolor="f7f7f7"><input name="IsPing" type="radio" class="noborder" value="1"<%IF ""&IsPing&""=1 then Response.Write("  checked") end if%>>
      ��
      <input name="IsPing" type="radio" class="noborder" value="0"<%IF ""&IsPing&""=0 then Response.Write("  checked") end if%>>
      ��</td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft">�����Ƿ���Ҫ��ˣ�</td>
    <td bgcolor="f7f7f7"><input name="pingoff" type="radio" class="noborder" value="0"<%IF ""&pingoff&""=0 then Response.Write("  checked") end if%>>
��
  <input name="pingoff" type="radio" class="noborder" value="1"<%IF ""&pingoff&""=1 then Response.Write("  checked") end if%>>
��</td>
  </tr>

  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft" width="19%">�����Ƿ���Ҫ��ˣ�</td>
    <td bgcolor="f7f7f7" width="79%"><input name="linkoff" type="radio" class="noborder" value="1"<%IF ""&linkoff&""=1 then Response.Write("  checked") end if%>>
��
  <input name="linkoff" type="radio" class="noborder" value="0"<%IF ""&linkoff&""=0 then Response.Write("  checked") end if%>>
��</td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft" width="19%">Ͷ���Ƿ���Ҫ��ˣ�</td>
    <td bgcolor="f7f7f7" width="79%"><input name="tougaooff" type="radio" class="noborder" value="1"<%IF ""&tougaooff&""=1 then Response.Write("  checked") end if%>>
��
  <input name="tougaooff" type="radio" class="noborder" value="0"<%IF ""&tougaooff&""=0 then Response.Write("  checked") end if%>>
��</td>
  </tr>
  <tr>
    <td height="25" bgcolor="f7f7f7" class="tdleft" width="19%">�����б�ÿҳ��ʾ��¼��</td>
    <td bgcolor="f7f7f7" width="79%"><input name="artlistnum" type="text" id="artlistnum" value="<%=artlistnum%>" size="5" maxlength="3">
      ��<span class="note">���·����б�ÿҳ��ʾ��¼��</span></td>
  </tr>
  </table>
        <table border="0" align="center" cellpadding="3" cellspacing="2" bgcolor="#FFFFFF" class="admintable">
          <tr>
            <td colspan="2" class="admintitle"> �ϴ�ͼƬˮӡ����</td>
          </tr>
          <tr>
            <td height="25" bgcolor="f7f7f7" class="tdleft">ͼƬˮӡ��</td>
            <td height="25" bgcolor="f7f7f7"><select name="AspJpeg" id="AspJpeg">
              <option value="1"<%If AspJpeg=1 then Response.Write(" selected") end if%>>��</option>
              <option value="0"<%If AspJpeg=0 then Response.Write(" selected") end if%>>��</option>
            </select>
            <%If IsObjInstalled("Persits.Jpeg") Then Response.Write "<font color=green><b>��</b>������֧��!</font>" Else Response.Write "<font color=red><b>��</b>��������֧��,��ѡ��رգ�����ᵼ����ҳ�õ�ͼƬ������ͼ�޷���ʾ��</font>" %></td>
          </tr>
          <tr>
            <td width="20%" height="25" bgcolor="f7f7f7" class="tdleft">ˮӡ���ִ�С��</td>
            <td height="25" bgcolor="f7f7f7">
			<SELECT name="FontSize" id="FontSize">
              <option value="<%=FontSize%>" selected><%=FontSize%>px</option>
              <option value="12">12px</option>
              <option value="14">14px</option>
              <option value="16">16px</option>
              <option value="18">18px</option>
			  <option value="22">22px</option>
			  <option value="32">32px</option>
			  <option value="48">48px</option>
			  <option value="56">56px</option>
            </SELECT></td>
          </tr>
          <tr>
            <td height="11" bgcolor="f7f7f7" class="tdleft">ˮӡ�������壺</td>
            <td height="-2" bgcolor="f7f7f7"><SELECT name="FontFamily" id="UploadSetting(4)">
      <option value="<%=FontFamily%>" selected><%=FontFamily%></option>
      <option value="����">����</option>
      <option value="����_GB2312">����</option>
      <option value="������">������</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <OPTION value="Andale Mono">Andale Mono</OPTION>
      <OPTION value=Arial>Arial</OPTION>
      <OPTION value="Arial Black">Arial Black</OPTION>
      <OPTION value="Book Antiqua">Book Antiqua</OPTION>
      <OPTION value="Century Gothic">Century Gothic</OPTION>
      <OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
      <OPTION value="Courier New">Courier New</OPTION>
      <OPTION value=Georgia>Georgia</OPTION>
      <OPTION value=Impact>Impact</OPTION>
      <OPTION value=Tahoma>Tahoma</OPTION>
      <OPTION value="Times New Roman" >Times New Roman</OPTION>
      <OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
      <OPTION value="Script MT Bold">Script MT Bold</OPTION>
      <OPTION value=Stencil>Stencil</OPTION>
      <OPTION value=Verdana>Verdana</OPTION>
      <OPTION value="Lucida Console">Lucida Console</OPTION>
            </SELECT></td>
          </tr>
          <tr>
            <td height="-2" bgcolor="f7f7f7" class="tdleft">ˮӡ������ɫ��</td>
            <td height="-2" bgcolor="f7f7f7">
              <input name="Color1" type="text" id="Color1" value="<%=Color1%>">
              ��</td>
          </tr>
          <tr>
            <td height="12" bgcolor="f7f7f7" class="tdleft">���ֱ�����ɫ��</td>
            <td height="12" bgcolor="f7f7f7">
              <input name="Color2" type="text" id="Color2" value="<%=Color2%>"></td>
          </tr>
          <tr>
            <td height="25" bgcolor="f7f7f7" class="tdleft">ˮӡ�������ݣ�</td>
            <td height="-1" bgcolor="f7f7f7"><input name="Fonttext" type="text" id="Fonttext" value="<%=Fonttext%>" size="40" maxlength="20"></td>
          </tr>
        </table>
  
        
        
   <br>     
	<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input name="Submit" type="submit" id="Submit" value="ȷ���޸�">
	<br>
		</form>
<%
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>

				
		</td>
	</tr>
</table>


</body>

</html>