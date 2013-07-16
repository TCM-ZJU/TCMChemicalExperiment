<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>方法步骤详情</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function jumpbz()
	{
	 	document.jumpform.BUZXLH.value=document.jumpform.bzselect.value
		document.jumpform.submit()
	}
	
	function clickform(GID,Tname,num)
	{
		document.GLform.GID.value=GID
		document.GLform.Tname.value=Tname
		document.GLform.sos.value=num
		document.GLform.submit()
	}
	
	function clickSHIYDX( SHIYDX)
	{
		document.formSYDX.SHIYDX.value=SHIYDX
		document.formSYDX.submit()
	}
-->
</script>
<%
dim HUAXSY_ID
HUAXSY_ID =request.form("HUAXSY_ID")
dim BUZXLH
BUZXLH =request.form("BUZXLH")

dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)

dim strtemp
dim oraRST
strtemp ="select GID, SHIYLX from CHOUXHXSY where HUAXSY_ID="&HUAXSY_ID&" and BUZXLH="&BUZXLH
set oraRST=oraDB.CreateDynaset(strtemp,0)
dim GID
GID= oraRST.fields(0).value
dim SYtype
SYtype=oraRST.fields(1).value
dim tableName
//由于magic和fullproject使用与数据库不一致的对应关系，以fullproject为准
select case SYtype
	case 1
		tableName ="HUAXCFTQ"
	case 2
		tableName ="HUAXCFFL"
	case 3
		tableName ="HUAXCFJD"
	case 4
		tableName ="HUAXCFHC"
	case 5
		tableName ="HUAXCFCH"
	case 6
		tableName ="HANLCD"
end select


strSQL ="select * from " &tableName& " where GID="&GID
//response.write strSQL
set oraRS=oraDB.CreateDynaset(strSQL,0)
%>
</head>

<body bgcolor="#B0DDFF">
<table class="clean" cellpadding="3" cellspacing="0">
<tr bgcolor="#FFDE5B">
<th>该实验方法的相关联接：</th>
<td><a href='javascript:clickform(<%=GID%>,"<%=tableName%>",1)'>实验室</a></td>
<td><a href='javascript:clickform(<%=GID%>,"<%=tableName%>",2)'>实验对象</a></td>
<td><a href='javascript:clickform(<%=GID%>,"<%=tableName%>",3)'>实验名称</a></td>
</tr>
</table>

<table class="clean">
<caption align=top>当前为实验第<%=BUZXLH%>步</caption>
<%
dim i
dim fnum
fnum =oraRS.fields.count
for i=1 to fnum-1
	if oraRS.fields(i).value <>"" then
		strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='"&tableName&"' and COLUMN_NAME='" &oraRS.fields(i).name &"'"
		set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<tr>
<th width=100><%=oraRST.fields(0).value%></th>
<%		if i=1 or i=2 then %>
<td><a href='javascript:clickSHIYDX("<%=oraRS.fields(i).value%>")'><%=oraRS.fields(i).value%></a></td>
<%		else %>
<td><%=oraRS.fields(i).value%></td>
</tr>
<%
		end if
	end if
next
strtemp="select BUZXLH from CHOUXHXSY where HUAXSY_ID="&HUAXSY_ID&" order by BUZXLH"
set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<caption align=bottom>本实验一共有<%=oraRST.RecordCount%>个步骤</caption>
</table>

<table>
<tr>
<form method=post action="fangfbz.asp">
<input type=hidden name=HUAXSY_ID value=<%=HUAXSY_ID%>>
<input type=hidden name=BUZXLH value=<%=(BUZXLH-1)%>>
<input 
 name=submit type=submit class="input" id=submit value=上一步 
<%
	if BUZXLH-1<=0 then
		Response.Write "disabled"
	end if
%>>
</form>
</tr>
<tr>
<form method=post action="fangfbz.asp">
<input type=hidden name=HUAXSY_ID value=<%=HUAXSY_ID%>>
<input type=hidden name=BUZXLH value=<%=(BUZXLH+1)%>>
<input 
 name=submit type=submit class="input" id=submit value=下一步 
<%
	if BUZXLH+1>oraRST.RecordCount then
		Response.Write "disabled"
	end if
%>>
</form>
</tr>
<tr>
<form method=post action="fangfbz.asp" name=jumpform>
<td>第</td>
<td><select name=bzselect>
<%
for i=1 to oraRST.RecordCount
%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select></td>
<td>步</td>
<input type=hidden name=HUAXSY_ID value=<%=HUAXSY_ID%>>
<input type=hidden name=BUZXLH >
<td><input type="button" value="跳转" onclick="jumpbz()" ></td>
</form>
</tr>
</table>

<form method=post action="soslist.asp"  name=GLform>
<input type=hidden name=GID>
<input type=hidden name=Tname>
<input type=hidden name=sos>
</form>
<form method=post action="shiydx.asp"  name=formSYDX>
<input type=hidden name=SHIYDX>
</form>

<div ALIGN=CENTER >
<a href="javascript:history.back()">返回</a><HR color=black>
</div>
</body>
</html>