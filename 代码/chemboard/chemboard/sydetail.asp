<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function clickSHIYS( SHIYS_ID)
	{
		document.formSYS.SHIYS_ID.value=SHIYS_ID
		document.formSYS.submit()
	}
	function clickFFBZ( HUAXSY_ID,BUZXLH)
	{
		document.formFFBZ.HUAXSY_ID.value=HUAXSY_ID
		document.formFFBZ.BUZXLH.value=BUZXLH
		document.formFFBZ.submit()
	}
-->
</script>

<%
dim SY_ID
SY_ID=request.form("HUAXSY_ID")
//response.write HUAXSY_ID
dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)
strSQL ="select * from HUAXUESY where HUAXSY_ID='" &SY_ID &"'"
set oraRS=oraDB.CreateDynaset(strSQL,0)
%>
</head>
<body bgcolor="#B0DDFF">
<table class="clean" cellpadding="3" cellspacing="0">
<tr bgcolor="#FFDE5B">
<th>该实验记录的相关联接：</th>
<td><a href='javascript:clickSHIYS(<%=oraRS.fields(1).value%>)'>实验室信息</a></td>
</tr>
</table>
<table class="clean">
<%
dim strtemp
dim i
dim fnum
dim oraRST

//response.write fnum
//实验的名称
if oraRS.fields(2).value <>"" then
	strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='HUAXUESY' and COLUMN_NAME='" &oraRS.fields(2).name &"'"
	set oraRST=oraDB.CreateDynaset(strtemp,0)
end if
%>
<tr>
<th><%=oraRST.fields(0).value%></th>
<td>
<FORM name=viewpic action="view.asp" method="POST" target=_blank>
<%=oraRS.fields(2).value%><INPUT type=submit value="查看图片">
<INPUT type=hidden name="table" value="PIC_HUAXUESY">
<INPUT type=hidden name="field" value="HUAXSY_ID">
<INPUT type=hidden name="value" value="<%=SY_ID%>">
</FORM>
</td>
</tr>
<tr>
<th>实验对象</th>
<td><%=oraRS.fields(5).value%></td>
</tr>
<%
//实验相关的实验室
if oraRS.fields(1).value <>"" then
	strtemp="select SHIYSMC from SHIYANSTJ where SHIYS_ID ='" &oraRS.fields(1).value &"'"
	set oraRST=oraDB.CreateDynaset(strtemp,0)
end if
%>
<tr>
<th>实验室名称</th>
<td>
<a href='javascript:clickSHIYS(<%=oraRS.fields(1).value%>)'>
<%
if oraRST.fields(0).value <>""then
	response.write oraRST.fields(0).value
else
	response.write "无名"
end if
%>
</a>
</td>
</tr>
<%
//实验相关的方法和步骤
strtemp ="SELECT BUZXLH , TONGYJCFF ,GID ,SHIYLX FROM CHOUXHXSY ,TONGYJCFF WHERE HUAXSY_ID ='"&SY_ID &"'"
strtemp =strtemp &"and CHOUXHXSY.TONGYJCFF_ID =TONGYJCFF.TONGYJCFF_ID ORDER BY BUZXLH"
set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<tr>
<th>实验方法和步骤 </th>
<td>
<table class="clean">
<%
dim sylx
Do Until oraRST.EOF
select case oraRST.fields(3).value
	case 1
		sylx="（化学成分提取）"
	case 2
		sylx="（化学成分分离）"
	case 3
		sylx="（化学成分鉴定）"
	case 4
		sylx="（化学成分合成）"
	case 5
		sylx="（化学成分纯化）"
	case 6
		sylx="（含量测定）"
end select
%>
<tr>
<th><%=oraRST.fields(0).value%></th>
<td>
<a href='javascript:clickFFBZ(<%=SY_ID%>, <%=oraRST.fields(0).value%>)'><%=oraRST.fields(1).value%></a>
</td>
<td><%=sylx%></td>
</tr>
<%
	oraRST.movenext
loop
%>
</table>
</td>
</tr>
<%
//实验相关的其他信息
fnum =oraRS.fields.count
for i=3 to fnum-1
	if oraRS.fields(i).value <>"" and i<>5 then
		strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='HUAXUESY' and COLUMN_NAME='" &oraRS.fields(i).name &"'"
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
<form method=post action="shiys.asp"  name=formSYS>
<input type=hidden name=SHIYS_ID>
</form>
<form method=post action="fangfbz.asp"  name=formFFBZ>
<input type=hidden name=HUAXSY_ID>
<input type=hidden name=BUZXLH>
</form>

<div ALIGN=CENTER >
<a href="javascript:history.back()">返回</a><HR color=black>
</div>
</body>
</html>