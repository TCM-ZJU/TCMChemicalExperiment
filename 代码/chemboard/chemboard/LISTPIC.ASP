<%@ Language=VBScript %>
<html>
<title>图片序号显示</title>
<style type="text/css">
<!--
a:link {
	text-decoration: none;
	color: blue;
}
a:visited {
	text-decoration: none;
	color: blue;
}
a:hover {
	text-decoration: none;
	color: red;
}
-->
</style>
<LINK href="script/table.css" rel="stylesheet" type="text/css">
<%
	Dim OraDatabase
	set OraDatabase = OraSession.getDatabaseFromPool(10)

'	response.write "select ID from " & Request("table") & " where " & Request("field") & "='" & Request("value") & "' order by ID"
	Dim PartImage
	set part = OraDatabase.CreateDynaset("select ID from " & Request("table") & " where " & Request("field") & "='" & Request("value") & "' order by ID",&0)
%>
</head>
<body bgcolor=#f4faed>
<%
        if part.RecordCount=0 then
                Response.Write "<font color=red><b>该文题没<br>有图片可<br>以显示！</b></font>"
        else
%>
<table width=100% border=1 cellPadding=3 cellSpacing=0 class="itable">
<tr><td align=center class="itd">图片列表</td></tr>
<%
                for i=0 to part.recordcount-1
%>
<tr><td align=center class="itd">
<a href="viewpic.asp?table=<%=Request("table")%>&ID=<% =part.Fields("ID").Value %>" target="content">第<% =i+1 %>页</a></center>
</td></tr>
<%
			part.MoveNext
        	next
%>
</table>
<%
        end if
%>
<p align=center><a href="javascript:history.back()">返回</a></p>
</body> 
</html>