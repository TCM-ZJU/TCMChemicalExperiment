<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>实验室信息</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
	function clickform( SYS_ID,num)
	{
		document.GLform.SYS_ID.value=SYS_ID
		document.GLform.oss.value=num
		document.GLform.submit()
	}
</script>

<%
dim SYS_ID
SYS_ID =request.form("SHIYS_ID")
//response.write SYS_ID
dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)
strSQL ="select * from SHIYANSTJ where SHIYS_ID=" &SYS_ID
set oraRS=oraDB.CreateDynaset(strSQL,0)
%>
</head>
<body bgcolor="#B0DDFF">
<table class="clean" cellpadding="3" cellspacing="0">
<tr bgcolor="#FFDE5B">
<th>该实验室的相关联接：</th>
<td><a href='javascript:clickform(<%=SYS_ID%>,1)'>实验对象</a></td>
<td><a href='javascript:clickform(<%=SYS_ID%>,2)'>实验方法</a></td>
<td><a href='javascript:clickform(<%=SYS_ID%>,3)'>实验名称</a></td>
</tr>
</table>
<table class="clean">

<%
dim strtemp
dim oraRST
dim i
dim fnum
fnum =oraRS.fields.count
for i=1 to fnum-1
	if oraRS.fields(i).value <>"" then
		strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='SHIYANSTJ' and COLUMN_NAME='" &oraRS.fields(i).name &"'"
		set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<tr>
<th><%=oraRST.fields(0).value%></th>
<td><%=oraRS.fields(i).value%></td>
</tr>
<%
	end if
next
%>
</table>
<form method=post action="osslist.asp"  name=GLform>
<input type=hidden name=SYS_ID>
<input type=hidden name=oss>
</form>

</body>

</html>