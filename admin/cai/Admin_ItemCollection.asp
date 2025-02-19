<%@language=vbscript codepage=936 %>
<%
option explicit
Response.Buffer = True 
Server.ScriptTimeOut=10
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/clsCache.asp"-->
<%
Dim Action,ItemID,CollecType
Dim FoundErr,ErrMsg
Dim SqlItem,RsItem
Dim Arr_Item,Arr_Filters,Arr_Histrolys,myCache,CollecTest,Content_View
Dim CacheTemp
FoundErr=False

CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

'检察表单
Call DelNews()
Call CheckForm()
If FoundErr<>True Then
   Call SetCache()
   If FoundErr<>True Then
      If CollecType=0 Then
         ErrMsg= "<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecSteady.asp?ItemNum=1&ListNum=1&ListSuccesNum=0&ListFalseNum=0&NewsNumAll=0"">"
      ElseIf CollecType=1 Then
         ErrMsg="<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecFast.asp?ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0"">"
      ElseIf CollecType=2 Then
         ErrMsg="<meta http-equiv=""refresh"" content=""3;url=Admin_ItemCollecScreen.asp?Action=GetList"">"
      End If
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
			
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="100" colspan="2" align=center>
      <br><br>
      欢迎使用[90站长]采集系统，正在初始化数据，请稍后...
      <br><br>
<%=ErrMsg%>
      </td>
    </tr>
</table>     
				
		</td>
	</tr>
</table>


</body>

</html>
<%End Sub

Sub CheckForm()

   '提取表单
   Action=trim(Request.Form("Action"))
   ItemID=Trim(Request.Form("ItemID"))
   CollecType=Trim(Request.Form("CollecType"))
   CollecTest=Trim(Request.Form("CollecTest"))
   Content_View=Trim(Request.Form("Content_View"))
   '检察表单
   If Action<>"Start" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数不足!</li>"
   End If
   If ItemID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请您选择项目!</li>"
   Else
      If Instr(ItemID,",")>0 Then
         ItemID=Replace(ItemID," ","")
      End If
   End If 
   If CollecType="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请您选择采集模式!</li>"
   Else
      CollecType=Clng(CollecType)
      If CollecType<>0 And CollecType<>1 And CollecType<>2 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>您选择的采集模式无效!</li>"
      End If
   End If
   If CollecTest="yes" Then
      CollecTest=True
   Else
      CollecTest=False
   End If
   If Content_View="yes" Then
      Content_View=True
   Else
      Content_View=False
   End If
End Sub
Sub SetCache()
   '项目信息
   SqlItem ="select * from Item where ItemID in(" & ItemID & ")"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Item=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   Set myCache=new clsCache
   myCache.name=CacheTemp & "items"
   Call myCache.clean()
   If IsArray(Arr_Item)=True Then
      myCache.add Arr_Item,Dateadd("n",1000,now)
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br>发生意外错误！"
   End If

   '过滤信息
   SqlItem ="select * from Filters where Flag=True"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Filters=RsItem.GetRows()
   End If
   RsItem.Close
   Set Rsitem=Nothing

   myCache.name=CacheTemp & "filters"
   Call myCache.clean()
   If IsArray(Arr_Filters)=True Then
      myCache.add Arr_Filters,Dateadd("n",1000,now)
   End If

   '历史记录
   SqlItem ="select NewsUrl,Title,CollecDate,Result from Histroly"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Histrolys=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()
   If IsArray(Arr_Histrolys)=True Then
      myCache.add Arr_Histrolys,Dateadd("n",1000,now)
   End If

   '其它信息
   myCache.name=CacheTemp & "collectest"
   Call myCache.clean()
   myCache.add CollecTest,Dateadd("n",1000,now)

   myCache.name=CacheTemp & "contentview"
   Call myCache.clean()
   myCache.add Content_View,Dateadd("n",1000,now)

   set myCache=nothing
End Sub
Sub DelNews()
   ConnItem.execute("Delete From [NewsList]")
End Sub
%>