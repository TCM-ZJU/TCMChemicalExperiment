<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function clickFFform(Mkeyword)
	{
		document.RSform.KEYorID.value=Mkeyword
		document.RSform.SQLtype.value=2		
		document.RSform.submit()
	}
	function clickSYSform(SYS_ID)
	{
		document.RSform.KEYorID.value=SYS_ID
		document.RSform.SQLtype.value=3	
		document.RSform.submit()
	}
	function clickSYform(HUAXSY_ID)
	{
	 	document.SYform.HUAXSY_ID.value=HUAXSY_ID
		document.SYform.submit()
	}
-->
</script>
<%
dim SYDX
dim sss
SYDX =request.form("SYDX")
sss =request.form("sss")

dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)

select case sss
case 1
	strSQL="select distinct SHIYS_ID, SHIYSMC from SHIYANSTJ natural inner join  HUAXUESY natural inner join CHOUXHXSY natural inner join HUAXSYBZ where SHIYDX ='"&SYDX&"' or SHIYWMC='"&SYDX&"' or GONGSP='"&SYDX&"'"
case 2
	strSQL="select distinct SHIYFF from HUAXUESY natural inner join CHOUXHXSY natural inner join HUAXSYBZ where SHIYDX='"&SYDX&"' or SHIYWMC='"&SYDX&"' or GONGSP='"&SYDX&"'"
case 3
	strSQL ="select distinct HUAXSYMC, HUAXSY_ID from HUAXUESY natural inner join CHOUXHXSY natural inner join HUAXSYBZ  where SHIYDX='"&SYDX&"' or SHIYWMC='"&SYDX&"' or GONGSP='"&SYDX&"'"
end select	
set oraRS=oraDB.CreateDynaset(strSQL,0)
dim recordnum 
recordnum =oraRS.recordcount

%>
</head>
<body bgcolor="#B0DDFF">
<table class="clean">
<%
select case sss
case 1
%>
	<tr><th nowrap>采用过该实验对象的实验室有<%=recordnum%>个</th></tr>
	<%DO Until oraRS.EOF%>
	<tr>
	<td><a href='javascript:clickSYSform("<%=oraRS.fields(0).value%>")'><%=oraRS.fields(1).value%></a></td>
	</tr>
	<%
	oraRS.movenext
	loop
	%>
<%case 2%>
	<tr><th nowrap>采用过该实验对象的实验方法有<%=recordnum%>种</th></tr>
	<%DO Until oraRS.EOF%>
	<tr>
	<td><a href='javascript:clickFFform("<%=oraRS.fields(0).value%>")'><%=oraRS.fields(0).value%></a></td>
	</tr>
	<%
	oraRS.movenext
	loop
	%>
<%case 3%>
	<tr><th nowrap>采用该实验对象的实验纪录有<%=recordnum%>条</th></tr>
	<%DO Until oraRS.EOF%>
	<tr>
	<td><a href='javascript:clickSYform(<%=oraRS.fields(1)%>)'><%=oraRS.fields(0).value%></a></td>
	</tr>
	<%
	oraRS.movenext
	loop
	%>
<%end select%>
</table>
<form method=post action="resultlist.asp?offset=1"   name=RSform>
<input type=hidden name= KEYorID>
<input type=hidden name= SQLtype>
</form>

<form method=post action="sydetail.asp"   name=SYform>
<input type=hidden name=HUAXSY_ID>
</form>

</body>
</html>