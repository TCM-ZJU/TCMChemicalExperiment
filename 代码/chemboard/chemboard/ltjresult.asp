<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type="text/css">
<!--
input { border: 1 double black }
table { background-color:#f2e7db; border-collapse: collapse; border: 1px double brown }
td { font-size:9 pt; border: 1px double brown }
-->
</STYLE>
<title>ͳ�ƽ��</title>
<LINK href="common.css" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
	function viewlist(sqlstr)
	{
		document.sqlform.sqlstring.value=sqlstr;
		document.sqlform.submit();
	}
-->
</script>
<%
    dim i
	dim num
	dim fieldnum
    
    dim selectStr
    dim fromStr
    dim whereStr
    dim groupStr
    
    dim qtype
    dim per
    if request.form("QueryType")="0" then
	qtype="="
	else 
	qtype=" like "
	per="%"
	end if
///////////////////////	
'����
	dim qttab(6)
	
	qttab(0)="SHIYWMC"
	qttab(1)="SHIYFF"
    qttab(2)="HUAXSYMC"
    qttab(3)="SHIYDXLYCD"
    qttab(4)="SHIYDXLYBW"
    qttab(5)="SHIYDXLYKS"
	
'ʵ����
    dim jgtab(21,2)
    
    jgtab(0,0)="HUAXCFCH"
    jgtab(0,1)="RONGD"
    jgtab(1,0)="HUAXCFCH"
    jgtab(1,1)="CHUNHHX"
    jgtab(2,0)="HUAXCFCH"
    jgtab(2,1)="CHUNHPL"
    jgtab(3,0)="HUAXCFCH"
    jgtab(3,1)="XUANGD"
    jgtab(4,0)="HUAXCFCH"
    jgtab(4,1)="HUISL"
    jgtab(5,0)="HUAXCFCH"
    jgtab(5,1)="CHUND"
    jgtab(6,0)="HUAXCFFL"
    jgtab(6,1)="CHENJPH"
    jgtab(7,0)="HUAXCFFL"
    jgtab(7,1)="CHENJXS"
    jgtab(8,0)="HUAXCFFL"
    jgtab(8,1)="CHENJSD"
    jgtab(9,0)="HUAXCFFL"
    jgtab(9,1)="CHENJSJ"
    jgtab(10,0)="HUAXCFFL"
    jgtab(10,1)="K"
    jgtab(11,0)="HUAXCFTQ"
    jgtab(11,1)="ZHONGL"
    jgtab(12,0)="HUAXCFTQ"
    jgtab(12,1)="SHOUL"
    jgtab(13,0)="HANLCD"
    jgtab(13,1)="JIANCPHL"
    jgtab(14,0)="HANLCD"
    jgtab(14,1)="JINGMDCD"
    jgtab(15,0)="HANLCD"
    jgtab(15,1)="XIANXGX"
    jgtab(16,0)="HANLCD"
    jgtab(16,1)="WUNDXSY"
    jgtab(17,0)="HANLCD"
    jgtab(17,1)="CHONGXXSY"
    jgtab(18,0)="HANLCD"
    jgtab(18,1)="XIANGDBZPC"
    jgtab(19,0)="HANLCD"
    jgtab(19,1)="CDHUISL"
    jgtab(20,0)="HANLCD"
    jgtab(20,1)="BIAOZQX"
	
    //���ҳ�����ؼ���������sql���
    dim lkeyword
    lkeyword=request.form("lkeyword")
    
//    selectSYstr="select HUAXSY_ID from HUAXUESY natural inner join SHIYANSTJ where SHIYSMC "&qtype&"'"&per&lkeyword&per&"'"
    
    //������sql���
    selectStr ="select SHIYSMC"
    fromStr =" from SHIYANSTJ natural inner join  HUAXUESY natural inner join CHOUXHXSY natural inner join HUAXSYBZ "
    whereStr =" where SHIYSMC "&qtype&"'"&per&lkeyword&per&"'"
    groupStr =" group by SHIYSMC"

    '����
    if request.form("syqtselect").count<>0 then
    	for i=1 to request.form("syqtselect").count
    		num =request.form("syqtselect")(i)-1
    		selectStr =selectStr&" , "&qttab(num)
    		groupStr =groupStr&" , "&qttab(num)
    	next
    end if
    'ʵ����
    if request.form("syjgselect").count<>0 then
    	for i=1 to request.form("syjgselect").count
    		num =request.form("syjgselect")(i)-1
    		selectStr =selectStr&" , "&jgtab(num,1)
    		groupStr =groupStr&" , "&jgtab(num,1)
    	next
    end if

    selectStr =selectStr&" , count(*) "

    dim recordnum
    dim oraDB
	dim oraRS
	dim strSQL
	set oraDB=OraSession.GetDatabaseFromPool(10)	
	strSQL =selectStr &fromStr &whereStr &groupStr &" order by count(*) desc "
	
	set oraRS=oraDB.CreateDynaset(strSQL,0)
	
	fieldnum=request.form("syqtselect").count+request.form("syjgselect").count+1
//	response.write strSQL
	recordnum =oraRS.recordcount
%>
</head>

<body>
<%
dim strtemp
dim oraRST

%>
<CENTER><H2><FONT face="����_GB2312" color="brown">ͳ�ƵĽ������<%=recordnum%>��</FONT></H2><HR color=BRown><BR>
<table class="black" border=1>
<tr bgcolor=#cc9966>
<td nowrap><font color=white><b>ʵ��������</b></font></td>
<%

//ʵ������
qttab(1)="CEDFF"
if request.form("syqtselect").count<>0 then
	for i=1 to request.form("syqtselect").count
		num =request.form("syqtselect")(i)-1
		strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='HUAXUESY' and COLUMN_NAME='" &qttab(num)&"'"
		set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<td><font color=white><b><%=oraRST.fields(0).value%></b></font></td>	
<%
	next	
end if
qttab(1)="SHIYFF"

//ʵ����
jgtab(19,1)="HUISL"
if request.form("syjgselect").count<>0 then
	for i=1 to request.form("syjgselect").count
		num =request.form("syjgselect")(i)-1
		strtemp="select COMMENTS from ALL_COL_COMMENTS where OWNER ='TCM_BASIC' and TABLE_NAME='"&jgtab(num,0)&"' and COLUMN_NAME='" &jgtab(num,1)&"'"
		set oraRST=oraDB.CreateDynaset(strtemp,0)
%>
<td><font color=white><b><%=oraRST.fields(0).value%></b></font></td>	
<%
	next	
end if
jgtab(19,1)="CDHUISL"
%>
<td nowrap><font color=white><b>ͳ�ƴ���</b></font></td>
<td align=center width=30 nowrap><font color=white><b>ԭ��</b></font></td>
</tr>
<%
if oraRS.recordcount<>0 then

oraRS.movefirst
Do Until oraRS.EOF
%>
<tr>
<%
	for i=1 to fieldnum
		num=i-1
%>
<td><%=oraRS.fields(num).value%></td>
<%	next

//ƴ��sql���
dim strT

strT= "select distinct HUAXSYMC,HUAXSY_ID,SHIYS_ID,SHIYDX,SHIYJG from SHIYANSTJ natural inner join HUAXUESY natural inner join CHOUXHXSY natural inner join HUAXSYBZ "
//ʵ����Լ������
   
strT= strT&" where SHIYSMC = '"&oraRS.fields(0)&"' "

//ʵ������Լ������
if request.form("syqtselect").count<>0 then
	for i=1 to request.form("syqtselect").count
		num =request.form("syqtselect")(i)-1
		strT= strT&" and "&qttab(num)
		num =i
		if IsNull(oraRS.fields(num).value) then
			strT=strT&" is null "
		else	
			strT= strT&" = '����"&oraRS.fields(num).value&"'"
		end if
	next
end if
//ʵ����Լ������
if request.form("syjgselect").count<>0 then
	for i=1 to request.form("syjgselect").count
		num =request.form("syjgselect")(i)-1
		strT= strT&" and "&jgtab(num,1)
		num =request.form("syqtselect").count +i
		if IsNull(oraRS.fields(num).value) then
			strT=strT&" is null "
		else	
			strT= strT&" = '����"&oraRS.fields(num).value&"'"
		end if
	next
end if

strT =Replace(strT,"'","&#039")
%>
<td><%=oraRS.fields(fieldnum).value%></td>
<td><a href='javascript:viewlist("<%=strT%>")'>�鿴 </a></td>
</tr>
<%
//response.write strT
oraRS.movenext
loop
end if//�Ƿ������ݼ�¼
%>
</table><BR>

<form method="post" action="tjlist.asp?offset=1" target=self name=sqlform >
<input type=hidden name="sqlstring">
</form>
<a href="javascript:history.back()">����</a><HR color=BRown>
<FONT color=BRown size=2>*****��ҽҩ�Ƽ���Ϣ���ݿ�*****</FONT>
</body>

</html>