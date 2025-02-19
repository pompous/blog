<head>
<meta charset="GB2312" />
		<title><%=SiteTitle%></title>
<script src="../js/menu.js" type="text/javascript"></script>


<style type="text/css">
*{ margin: 0; padding: 0; }
.header-wrap{ width: 100%;z-index:99; }
.header-hd{ width: 100%; height: 150px; margin: 0 auto; }
.header-bd{ width: 100%; height: 90px; margin: 0 auto; background-color:#cccccc }
</style>
<script type="text/javascript">
	window.onload=function(){
		function adsorption(){
			var headerWrap=document.getElementById('header-wrap');
			var scrollTop=0;
			window.onscroll=function(){
		scrollTop=document.body.scrollTop||document.documentElement.scrollTop;
			if(scrollTop>100){
			headerWrap.className='fixed';
			}else{
			headerWrap.className='header-wrap';
			}
			}
			}
		adsorption();
		}
	</script>



</head>

<body topmargin="0" leftmargin="0"><%call spiderbot()%>
<div id="header-wrap" class="header-wrap">
		<div class="header-hd">
<table border="0" width="100%" cellspacing="0" height="90"  bgcolor="#cccccc">
	<tr>
		<td><center>
		<table border="0" width="980" id="table3" height="150" cellspacing="0" cellpadding="0" bgcolor="#cccccc" >
			<tr>
				<td width="222">
				<img border="0" src="../images/logo.png" width="222" height="54"></td>
				<td width="87">
				<p align="center">　</td>
			</tr>
			</table>
		
		</td>
	</tr>
</table>
</div>
		<div class="header-bd">
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table4" height="40" bgcolor="#333333">
	<tr>
		<td><center>
		<table border="0" width="980" id="table4" height="45" cellspacing="0" cellpadding="0">
			<tr>
				<td><b><%=Menu%></b></td>
			</tr>
		</table></td>
	</tr>
</table>
<center></div></div>