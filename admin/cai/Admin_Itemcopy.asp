<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->

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
			
<%
Dim Action,ItemID,Arr_Item,Arr_i,ItemName,Arr_Filter
Action=Trim(Request("Action"))
ItemID=Trim(Request("ItemID"))
FoundErr=False

If Action<>"Copy" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
End if   
If ItemID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If

If FoundErr<>True Then
   Set Rs=ConnItem.execute("Select * from Item Where ItemID=" & ItemID)
   If Not Rs.Eof And Not Rs.Bof Then
      Arr_Item=Rs.GetRows()
   Else
      Arr_Item=""
   End if
   Set Rs=Nothing
   
   If IsArray(Arr_Item)=True Then
      set rs=server.createobject("adodb.recordset")
      sql="select top 1 * from Item" 
      rs.open sql,connItem,1,3
      rs.AddNew
         rs(1)=Arr_Item(1,0)&"复件"
         ItemName=Arr_Item(1,0)
         For Arr_i=2 To Ubound(Arr_Item,1)
            rs(Arr_i)=Arr_Item(Arr_i,0)
         Next
         ItemID=rs("ItemID")
      rs.Update
      rs.close
      set rs=Nothing

      If IsArray(Arr_Filter)=True Then
         set rs=server.createobject("adodb.recordset")
         sql="select top 1 * from Filters" 
         rs.open sql,connItem,1,3
         rs.AddNew
            rs(1)=ItemID
            For Arr_i=2 To Ubound(Arr_Filter,1)
               rs(Arr_i)=Arr_Filter(Arr_i,0)
            Next
         rs.Update
         rs.close
         set rs=Nothing
      End if      
   else
      FoundErr=True
      ErrMsg=ErrMsg & "参数错误，没有找到该项目"
   end if
   
End if

Call Main  
'关闭数据库链接

Sub Main%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1" style="text-align:center;padding:50px;"><%=ItemName%> 项目复制完成，新的项目保存为：<%=ItemName%>复件<br><br><a href="Admin_ItemStart.asp">返回</a></td>
  </tr>
</table>
<%End Sub%>
	
		</td>
	</tr>
</table>


</body>

</html>
