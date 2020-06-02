<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,rfc,nom,direc,tel,ciud,codPost,NomTarj

rfc=Request.Form("RFC")
nom=Request.Form("nombre")
direc=Request.Form("direccion")
tel=Request.Form("telefono")
ciud=Request.Form("ciudad")
codPost=Request.Form("codigopostal")
NomTarj=Request.Form("NoTarjeta")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bryan Lopez\Desktop\Proyecto Carrillo\Personal.mdb"
Conn.execute "INSERT INTO Clientes(RFC,Nombre,Direccion,Telefono,Ciudad,Codigo_Postal,No_Tarjeta) values('"& rfc & "','"& nom & "','"& direc & "','"& tel & "','"& ciud & "','"& codPost & "','"& NomTarj & "')"
Conn.close
set conn=nothing
response.redirect("gracias.html")

%>
</body>
</html>