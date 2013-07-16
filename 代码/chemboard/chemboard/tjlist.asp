<%@ Language=VBScript %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查看结果</title>
<link href="link" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function clickform(HUAXSY_ID)
	{
	 	document.HXSYform.HUAXSY_ID.value=HUAXSY_ID
		document.HXSYform.submit()
	}
	function jumppn()
	{
		document.jumpform.offset.value=(document.jumpform.pnselect.value-1)*10+1
		document.jumpform.submit()
	}
-->	
</script>
<%
dim recordnum

dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)
strSQL =request.form("sqlstring")
strSQL =Replace(strSQL,"范宽","")
//	response.write strSQL
set oraRS=oraDB.CreateDynaset(strSQL,0)
recordnum =oraRS.recordcount

dim row
dim intPage
dim intoffset
dim intUBound
dim pagenum
dim sum

intPage=10
intOffset=Request.QueryString("offset")
pagenum=Cint(recordnum/intPage)
sum =pagenum*intPage
if recordnum>sum then
	pagenum=pagenum+1
end if

%>

</head>
<body bgcolor="#B0DDFF">
<table class="clean" cellpadding="0" cellspacing="0" >
<tr bgcolor="#FFDE5B">
<td width="16%" height="10"><div align="left">:: 查询结果：</div></td>
<td><div align="center">符合条件的记录共有<%=recordnum%>条，本页显示<%=intOffset%>—<%if intOffset+intPage-1 > recordnum then response.write recordnum else response.write intOffset+intPage-1 end if%>条，共<%=pagenum%>页</div></td>
</tr>
</table>
<table class="orange" width=100% cellpadding="0" cellspacing="0" >
<caption align=bottom>第<%=(intOffset-1)/intPage+1%>页</caption>
<%
if recordnum<>0 then
	if intOffset+intPage> recordnum then
		intUBound=recordnum-intOffset
	else
		intUBound=intPage-1
	end if
%>
	<tr bgcolor="#D3D0FD">
	<th width=13% align="center">实验对象</th><th width=12% align="center">实验室</th><th width=25% align="center">实验方法和步骤<br>方法名称&#032方法类型&#032步骤计数</th><th width=30% align="center">实验结果</th><th width=20% align="center">实验名称</th>
	</tr>
<tr bgcolor="#B501DC"><td colspan=5><table class="insidetable" cellpadding="3" cellspacing="1" >
	<%
	dim strT
	dim oraRST
	on error resume next
	oraRS.MoveTo intOffset
	for row = 0 to intUBound
		strT="select SHIYSMC from SHIYANSTJ where SHIYS_ID="&oraRS.fields(2).value
		set oraRST=oraDB.CreateDynaset(strT,0)
	%>
	<tr bgcolor="#B0DDFF">
	<td width=13% align="center"><%=oraRS.fields(3).value%></td>
	<td width=12% align="center"><%if oraRST.recordcount<>0 then response.write oraRST.fields(0).value else response.write "无名" end if%></td>
	<td width=25%>
	<%
	//实验相关的方法和步骤
	strT ="SELECT  TONGYJCFF , SHIYLX , count( TONGYJCFF) FROM CHOUXHXSY ,TONGYJCFF WHERE HUAXSY_ID ="&oraRS.fields(1).value
	strT =strT &"and CHOUXHXSY.TONGYJCFF_ID =TONGYJCFF.TONGYJCFF_ID group by TONGYJCFF , SHIYLX"
	set oraRST=oraDB.CreateDynaset(strT,0)	
	%>
	<table class="clean" width=100%  cellspacing="0" cellpadding="0">
	<%
	dim sylx
	Do Until oraRST.EOF
		select case oraRST.fields(1).value
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
	<td width=45%><%=oraRST.fields(0).value%></td>
	<td width=45%><%=sylx%></td>
	<td width=10% align="center"><%=oraRST.fields(2).value%></td>
	</tr>
	<%
	oraRST.movenext
	loop
	%>
	</table>
	</td>
	<td width=30%><%=oraRS.fields(4).value%></td>
	<td width=20%>
	<a href='javascript:clickform(<%=oraRS.fields(1)%>)'> <%=oraRS.fields(0).value%></a>
	</td>
	</tr>
	<%
	oraRS.movenext
	next
	response.write  errinformation
%>
</table></td></tr>
</table>
<form method=post action="sydetail.asp"   name=HXSYform>
<input type=hidden name=HUAXSY_ID>
</form>

<TABLE width=100%>
<TR><TD align=right width=35%>
<FORM method=post action="tjlist.asp?offset=<% =(intOffset-intPage) %>">
<input type=hidden name=sqlstring value="<%=request.form("sqlstring")%>">
<INPUT
 name=submit type=submit class="input" id=submit value=上一页 
<%
	if intOffset-intPage<=0 then
		Response.Write "disabled"
	end if
%>>
</FORM></td>
<TD align=left width=35%>
<FORM method=post action="tjlist.asp?offset=<% =(intOffset+intPage) %>">
<input type=hidden name=sqlstring value="<%=request.form("sqlstring")%>">
<INPUT
 name=submit type=submit class="input" id=submit value=下一页 
<%
	if intOffset+intPage>recordnum then
		Response.Write "disabled"
	end if
%>>
</FORM></td><td>
<form method=post action="tjlist.asp?isNew=0" name=jumpform>
<td>第</td>
<td><select name=pnselect>
<%
dim i
for i=1 to pagenum
%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select></td>
<td>页</td>
<input type=hidden name=sqlstring value="<%=request.form("sqlstring")%>">
<input type=hidden name=offset >
<td><input type="button" value="跳转" onclick="jumppn()" ></td>
</form>
</td></tr>
<%
end if
%>
</table>

</body>
</html>