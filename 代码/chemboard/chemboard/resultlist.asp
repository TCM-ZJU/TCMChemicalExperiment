<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>resultlist</title>
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
dim KEYorID
dim SQLtype

KEYorID =request.form("KEYorID")
SQLtype =request.form("SQLtype")

dim oraDB
dim oraRS
dim strSQL
set oraDB=OraSession.GetDatabaseFromPool(10)

select case SQLtype
	case 1
		strSQL ="select distinct HUAXSYMC,HUAXSY_ID,SHIYS_ID,SHIYDX,SHIYJG from HUAXUESY natural inner join CHOUXHXSY  where SHIYDX ='"&KEYorID&"' or GID in("
		strSQL =strSQL&" (select GID from HANLCD where JIANCPMC='"&KEYorID&"')"
		strSQL =strSQL&" union (select GID from HUAXCFCH where CHUNHWMC='"&KEYorID&"')"
		strSQL =strSQL&" union (select GID from HUAXCFFL where FENLWMC='"&KEYorID&"')"
		strSQL =strSQL&" union (select GID from HUAXCFHC where HECWMC='"&KEYorID&"')"
		strSQL =strSQL&" union (select GID from HUAXCFJD where JIANDWMC='"&KEYorID&"')"
		strSQL =strSQL&" union (select GID from HUAXCFTQ where TIQWMC='"&KEYorID&"'))"
	case 2
		strSQL ="select distinct HUAXSYMC,HUAXSY_ID,SHIYS_ID,SHIYDX,SHIYJG from HUAXUESY natural inner join CHOUXHXSY  natural inner join TONGYJCFF where TONGYJCFF ='"&KEYorID&"'"
	case 3
		strSQL ="select distinct HUAXSYMC,HUAXSY_ID,SHIYS_ID,SHIYDX,SHIYJG from HUAXUESY  where SHIYS_ID ="&KEYorID
end select

set oraRS=oraDB.CreateDynaset(strSQL,0)
dim recordnum

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
<td width="16%" height="10"><div align="left">:: ��ѯ�����</div></td>
<td><div align="center">���������ļ�¼����<%=recordnum%>������ҳ��ʾ<%=intOffset%>��<%if intOffset+intPage-1 > recordnum then response.write recordnum else response.write intOffset+intPage-1 end if%>������<%=pagenum%>ҳ</div></td>
</tr>
</table>
<table class="orange" width=100% cellpadding="0" cellspacing="0" >
<caption align=bottom>��<%=(intOffset-1)/intPage+1%>ҳ</caption>
<%
if recordnum<>0 then
	if intOffset+intPage> recordnum then
		intUBound=recordnum-intOffset
	else
		intUBound=intPage-1
	end if
%>
	<tr>
	<th width=13% align="center">ʵ�����</th><th width=12% align="center">ʵ����</th><th width=25% align="center">ʵ�鷽���Ͳ���<br>��������&#032��������&#032�������</th><th width=30% align="center">ʵ����</th><th width=12% align="center">ʵ������</th>
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
	<td width=12% align="center"><%if oraRST.recordcount<>0 then response.write oraRST.fields(0).value else response.write "����" end if%></td>
	<td width=25%>
	<%
	//ʵ����صķ����Ͳ���
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
				sylx="����ѧ�ɷ���ȡ��"
			case 2
				sylx="����ѧ�ɷַ��룩"
			case 3
				sylx="����ѧ�ɷּ�����"
			case 4
				sylx="����ѧ�ɷֺϳɣ�"
			case 5
				sylx="����ѧ�ɷִ�����"
			case 6
				sylx="�������ⶨ��"
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
<FORM method=post action="resultlist.asp?offset=<% =(intOffset-intPage) %>">
<INPUT type=hidden name=KEYorID value=<% =request.form("KEYorID") %>>
<INPUT type=hidden name=SQLtype value=<% =request.form("SQLtype") %>>

<INPUT
 name=submit type=submit class="input" id=submit value=��һҳ 
<%
	if intOffset-intPage<=0 then
		Response.Write "disabled"
	end if
%>>
</FORM></td>
<TD align=left width=35%>
<FORM method=post action="resultlist.asp?offset=<% =(intOffset+intPage) %>">
<INPUT type=hidden name=KEYorID value=<% =request.form("KEYorID") %>>
<INPUT type=hidden name=SQLtype value=<% =request.form("SQLtype") %>>

<INPUT
 name=submit type=submit class="input" id=submit value=��һҳ 
<%
	if intOffset+intPage>recordnum then
		Response.Write "disabled"
	end if
%>>
</FORM></td><td>
<form method=post action="resultlist.asp?isNew=0" name=jumpform>
<td>��</td>
<td><select name=pnselect>
<%
dim i
for i=1 to pagenum
%>
<option value=<%=i%>><%=i%></option>
<%next%>
</select></td>
<td>ҳ</td>
<INPUT type=hidden name=KEYorID value=<% =request.form("KEYorID") %>>
<INPUT type=hidden name=SQLtype value=<% =request.form("SQLtype") %>>

<input type=hidden name=offset >
<td><input type="button" value="��ת" onclick="jumppn()" ></td>
</form>
</td></tr>
<%
end if
%>
</table>

</body>
</html>