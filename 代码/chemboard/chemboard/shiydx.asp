<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ʵ�������Ϣ</title>
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
<th>��ʵ������������ӣ�</th>
<td><a href='javascript:clickform("<%=SYDX%>",1)'>ʵ����</a></td>
<td><a href='javascript:clickform("<%=SYDX%>",2)'>ʵ�鷽��</a></td>
<td><a href='javascript:clickform("<%=SYDX%>",3)'>ʵ������</a></td>
</tr>
</table>
<table class="clean">
<tr>
<th>ʵ�����</th>
<td><%=SYDX%></td>
</tr>
</table>
<form method=post action="ssslist.asp"  name=GLform>
<input type=hidden name=SYDX>
<input type=hidden name=sss>
</form>

</body>
</html>