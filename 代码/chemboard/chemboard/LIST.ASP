<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
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
dim Otablename
dim Ltablename 
dim Mtablename
dim Rtablename
dim Ocolumn
dim Lcolumn
dim Mcolumn
dim Rcolumn
dim recordnum
recordnum =0

dim oraDB
dim oraRS
dim strSQL
dim qtype
dim per
	
if request.form("QueryType")="0" then
	qtype="="
else 
	qtype=" like "
	per="%"
end if

if request.form("Okeyword")<>"" then
	Ocolumn =request.form("concept")
	select case Ocolumn
	case "SHIYDX"
		Otablename="FUHE"
	case "SHIYDXLYBW"
		Otablename="HUAXUESY"
	case "SHIYDXLYCD"
		Otablename="HUAXUESY"
	case "SHIYDXLYKS"
		Otablename="HUAXUESY"
	end select
end if

if request.form("Lkeyword")<>"" then
	Ltablename="SHIYANSTJ"
	Lcolumn=request.form("Lab")
end if

if request.form("Mkeyword")<>"" then
	Mtablename =request.form("method")
	select case Mtablename
	case "CEDFF"
		Mcolumn =Mtablename
		Mtablename="HUAXUESY"
	case "HUAXCFCH"
		Mcolumn="CHUNHFF"
	case "HUAXCFFL"
		Mcolumn="FENLFF"
	case "HUAXCFTQ"
		Mcolumn="TIQFF"
	case "HUAXCFHC"
		Mcolumn="HECFF"
	case "HUAXCFJD"
		Mcolumn="JIANDFF"
	case "HANLCD"
		Mcolumn="CEDFF"
	end select
end if

if request.form("result")<>"不选" then
	Rcolumn=request.form("result")
	select case Rcolumn
	case "RONGD"
		Rtablename="HUAXCFCH"
	case "XUANGD"
		Rtablename="HUAXCFCH"
	case "CHUNHHX"
		Rtablename="HUAXCFCH"
	case "CHUNHPL"
		Rtablename="HUAXCFCH"
	case "HUISL1"
		Rcolumn="HUISL"
		Rtablename="HUAXCFCH"
	case "CHUND"
		Rtablename="HUAXCFCH"
	case "CHENJPH"
		Rtablename="HUAXCFFL"
	case "CHENJXS"
		Rtablename="HUAXCFFL"
	case "CHENJSD"
		Rtablename="HUAXCFFL"
	case "CHENJSJ"
		Rtablename="HUAXCFFL"
	case "K"
		Rtablename="HUAXCFFL"
	case "ZHONGL"
		Rtablename="HUAXCFTQ"
	case "SHOUL"
		Rtablename="HUAXCFTQ"
	case "JIANCPHL"
		Rtablename="HANLCD"
	case "JINGMDCD"
		Rtablename="HANLCD"
	case "XIANXGX"
		Rtablename="HANLCD"
	case "WUNDXSY"
		Rtablename="HANLCD"
	case "CHONGXXSY"
		Rtablename="HANLCD"
	case "HUISL2"
		Rcolumn="HUISL"
		Rtablename="HANLCD"
	case "BIAOZQX"
		Rtablename="HANLCD"
	case "XIANGDBZPC"
		Rtablename="HANLCD"	
	end select
end if

dim OstrSQL
dim LstrSQL
dim MstrSQL
dim RstrSQL
dim startstr
startstr=0
//dim sqlnum=0
if Otablename ="FUHE" then
	startstr=1
	OstrSQL ="select HUAXSY_ID from CHOUXHXSY where GID in("
	OstrSQL =OstrSQL&" (select GID from HANLCD where JIANCPMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"')"
	OstrSQL =OstrSQL&" union (select GID from HUAXCFCH where CHUNHWMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"')"
	OstrSQL =OstrSQL&" union (select GID from HUAXCFFL where FENLWMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"')"
	OstrSQL =OstrSQL&" union (select GID from HUAXCFHC where HECWMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"')"
	OstrSQL =OstrSQL&" union (select GID from HUAXCFJD where JIANDWMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"')"
	OstrSQL =OstrSQL&" union (select GID from HUAXCFTQ where TIQWMC "&qtype&"'"&per&request.form("Okeyword")&per&"' or GONGSP "&qtype&"'"&per&request.form("Okeyword")&per&"'))"	
else
	if Otablename <>"" then
		startstr=1 
		OstrSQL ="select HUAXSY_ID from HUAXUESY where "&Ocolumn&qtype&"'"&per&request.form("Okeyword")&per&"'"
	end if
end if

if Ltablename <>"" then
	if startstr=0 then
		startstr=2
	end if  
	LstrSQL ="select HUAXSY_ID from HUAXUESY where SHIYS_ID in (select SHIYS_ID from SHIYANSTJ where "&Lcolumn&qtype&"'"&per&request.form("Lkeyword")&per&"')"
end if

if Mtablename ="HUAXUESY" then
	if startstr=0 then
		startstr=3
	end if 
	MstrSQL ="select HUAXSY_ID from HUAXUESY where CEDFF"&qtype&"'"&per&request.form("Mkeyword")&per&"'"
else
	if Mtablename<>"" then 
		if startstr=0 then
			startstr=3
		end if
		MstrSQL ="select HUAXSY_ID from CHOUXHXSY where GID in (select GID from "&Mtablename&" where "&Mcolumn&qtype&"'"&per&request.form("Mkeyword")&per&"')"
	end if
end if

if Rtablename <>"" then
	if startstr=0 then
		startstr=4
	end if
	RstrSQL ="select HUAXSY_ID from CHOUXHXSY where GID in (select GID from "&Rtablename&" where "&Rcolumn&" is not null )"
end if

set oraDB=OraSession.GetDatabaseFromPool(10)		
strSQL ="select HUAXSYMC,HUAXSY_ID,SHIYS_ID,SHIYDX,SHIYJG from HUAXUESY"
select case startstr
case 1
	strSQL=strSQL&" where HUAXSY_ID in(("&OstrSQL&")"
	if LstrSQL<>"" then
		strSQL=strSQL&" intersect("&LstrSQL&")"
	end if
	if MstrSQL<>"" then
		strSQL=strSQL&" intersect("&MstrSQL&")"
	end if
	if RstrSQL<>"" then
		strSQL=strSQL&" intersect("&RstrSQL&")"
	end if
	strSQL=strSQL&")"
case 2
	strSQL=strSQL&" where HUAXSY_ID in(("&LstrSQL&")"
	if MstrSQL<>"" then
		strSQL=strSQL&" intersect("&MstrSQL&")"
	end if
	if RstrSQL<>"" then
		strSQL=strSQL&" intersect("&RstrSQL&")"
	end if
	strSQL=strSQL&")"
case 3
	strSQL=strSQL&" where HUAXSY_ID in(("&MstrSQL&")"
	if RstrSQL<>"" then
		strSQL=strSQL&" intersect("&RstrSQL&")"
	end if
	strSQL=strSQL&")"
case 4	
	strSQL=strSQL&" where HUAXSY_ID in("&RstrSQL&")"
end select

set oraRS=oraDB.CreateDynaset(strSQL,0)

recordnum =oraRS.recordcount

dim row
dim intPage
dim intoffset
dim intUBound
dim pagenum
dim sum

intPage=10
intOffset=Cint(Request("offset"))
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
	<table class="clean" width=100% cellpadding="0" cellspacing="0">
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
<form method=post action="sydetail.asp" target=main  name=HXSYform>
<input type=hidden name=HUAXSY_ID>
</form>

<TABLE width=100%>
<TR><TD align=right width=35%>
<FORM method=post action="list.asp?offset=<% =(intOffset-intPage) %>&isNew=0">
<INPUT type=hidden name=concept value=<% =request.form("concept") %>>
<INPUT type=hidden name=Okeyword value=<% =request.form("Okeyword") %>>
<INPUT type=hidden name=Lab value=<% =request.form("Lab") %>>
<INPUT type=hidden name=Lkeyword value=<% =request.form("Lkeyword") %>>
<INPUT type=hidden name=method value=<% =request.form("method") %>>
<INPUT type=hidden name=Mkeyword value=<% =request.form("Mkeyword") %>>
<INPUT type=hidden name=result value=<% =request.form("result") %>>
<input type=hidden name=QueryType value=<% =request.form("QueryType")%>>
<INPUT
 name=submit type=submit class="input" id=submit value=上一页 
<%
	if intOffset-intPage<=0 then
		Response.Write "disabled"
	end if
%>>
</FORM></td>
<TD align=left width=35%>
<FORM method=post action="list.asp?offset=<% =(intOffset+intPage) %>&isNew=0">
<INPUT type=hidden name=concept value=<% =request.form("concept") %>>
<INPUT type=hidden name=Okeyword value=<% =request.form("Okeyword") %>>
<INPUT type=hidden name=Lab value=<% =request.form("Lab") %>>
<INPUT type=hidden name=Lkeyword value=<% =request.form("Lkeyword") %>>
<INPUT type=hidden name=method value=<% =request.form("method") %>>
<INPUT type=hidden name=Mkeyword value=<% =request.form("Mkeyword") %>>
<INPUT type=hidden name=result value=<% =request.form("result") %>>
<input type=hidden name=QueryType value=<% =request.form("QueryType")%>>
<INPUT
 name=submit type=submit class="input" id=submit value=下一页 
<%
	if intOffset+intPage>recordnum then
		Response.Write "disabled"
	end if
%>>
</FORM></td><td>
<form method=post action="list.asp?isNew=0" name=jumpform>
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
<INPUT type=hidden name=concept value=<% =request.form("concept") %>>
<INPUT type=hidden name=Okeyword value=<% =request.form("Okeyword") %>>
<INPUT type=hidden name=Lab value=<% =request.form("Lab") %>>
<INPUT type=hidden name=Lkeyword value=<% =request.form("Lkeyword") %>>
<INPUT type=hidden name=method value=<% =request.form("method") %>>
<INPUT type=hidden name=Mkeyword value=<% =request.form("Mkeyword") %>>
<INPUT type=hidden name=result value=<% =request.form("result") %>>
<input type=hidden name=QueryType value=<% =request.form("QueryType")%>>
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