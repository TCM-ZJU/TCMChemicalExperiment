<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function clickDXform(Okeyword)
	{
		document.RSform.KEYorID.value=Okeyword
		document.RSform.SQLtype.value=1		
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
dim GID
dim Tname
dim sos
GID=request.form("GID")
Tname=request.form("Tname")
sos=request.form("sos")
dim DXMC

select case Tname
	case "HUAXCFTQ"
		DXMC="TIQWMC"
	case "HUAXCFFL"
		DXMC="FENLWMC"
	case "HUAXCFJD"
		DXMC="JIANDWMC"
	case "HUAXCFHC"
		DXMC="HECWMC"
	case "HUAXCFCH"
		DXMC="CHUNHWMC"
	case "HANLCD"
		DXMC="JIANCPMC"
end select

dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)

select case sos
case 1
	strSQL="select SHIYS_ID, SHIYSMC from SHIYANSTJ where SHIYS_ID in ( select SHIYS_ID from HUAXUESY natural inner join CHOUXHXSY where TONGYJCFF_ID =(select TONGYJCFF_ID from CHOUXHXSY where GID="&GID&"))"
case 2
	strSQL="select distinct "&DXMC&" from CHOUXHXSY  natural inner join "&Tname&" where TONGYJCFF_ID =(select TONGYJCFF_ID from CHOUXHXSY where GID="&GID&")"
case 3
	strSQL ="select distinct HUAXSYMC, HUAXSY_ID from HUAXUESY natural inner join CHOUXHXSY where TONGYJCFF_ID =(select TONGYJCFF_ID from CHOUXHXSY where GID="&GID&")"
end select	
set oraRS=oraDB.CreateDynaset(strSQL,0)
dim recordnum 
recordnum =oraRS.recordcount

%>
</head>
<body bgcolor="#B0DDFF">
<table class="clean">
<%
select case sos
case 1
%>
	<tr><th nowrap>采用过该实验方法的实验室有<%=recordnum%>个</th></tr>
	<%DO Until oraRS.EOF%>
	<tr>
	<td><a href='javascript:clickSYSform("<%=oraRS.fields(0).value%>")'><%=oraRS.fields(1).value%></a></td>
	</tr>
	<%
	oraRS.movenext
	loop
	%>
<%case 2%>
	<tr><th nowrap>采用过该实验方法的实验对象有<%=recordnum%>种</th></tr>
	<%DO Until oraRS.EOF%>
	<tr>
	<td><a href='javascript:clickDXform("<%=oraRS.fields(0).value%>")'><%=oraRS.fields(0).value%></a></td>
	</tr>
	<%
	oraRS.movenext
	loop
	%>
<%case 3%>
	<tr><th nowrap>采用该实验方法的实验纪录有<%=recordnum%>条</th></tr>
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