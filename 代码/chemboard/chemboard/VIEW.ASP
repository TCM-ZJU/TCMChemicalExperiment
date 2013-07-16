<%@ Language=VBScript %>
<html>
<title>Õº∆¨œ‘ æ</title>
<style type="text/css">
<!--
a:link {
	text-decoration: none;
	color: black;
}
a:visited {
	text-decoration: none;
	color: black;
}
a:hover {
	text-decoration: none;
	color: red;
}
-->
</style>
<LINK href="script/table.css" rel="stylesheet" type="text/css">
<frameset cols="100,*" frameborder="no">
<frame src="listpic.asp?table=<%=Request("table")%>&field=<%=Request("field")%>&value=<%=Server.URLEncode(Request("value"))%>" name="menu" scrolling="no" noresize>
<frame src="showpic.asp?strPICNO=<%=strPICNO%>" name="content" scrolling="yes">
</frame>
</html>