<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>实验对象信息</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function clickform( SYDX,num)
	{
		document.GLform.SYDX.value=SYDX
		document.GLform.sss.value=num
		document.GLform.submit()
	}
-->
</script>
<%
dim SYDX
SYDX =request.form("SHIYDX")
%>
</head>
<body bgcolor="#B0DDFF">
<table class="clean" cellpadding="3" cellspacing="0">
<tr bgcolor="#FFDE5B">
<th>该实验对象的相关联接：</th>
<td><a href='javascript:clickform("<%=SYDX%>",1)'>实验室</a></td>
<td><a href='javascript:clickform("<%=SYDX%>",2)'>实验方法</a></td>
<td><a href='javascript:clickform("<%=SYDX%>",3)'>实验名称</a></td>
</tr>
</table>
<table class="clean">
<tr>
<th>实验对象</th>
<td><%=SYDX%></td>
</tr>
</table>
<form method=post action="ssslist.asp"  name=GLform>
<input type=hidden name=SYDX>
<input type=hidden name=sss>
</form>

</body>
</html>