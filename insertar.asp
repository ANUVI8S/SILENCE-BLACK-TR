<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments

nom=Request.Form("nombre")
email=request.Form("correo")
coments=Request.Form("coms")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bryan Lopez\Desktop\Proyecto Carrillo\Personal.mdb"
Conn.execute "INSERT INTO comentarios(nombre,correo,coments) values('"& nom & "','"& email & "','"& coments & "')"
Conn.close
set conn=nothing
response.redirect("gracias.html")

%>
</body>
</html>