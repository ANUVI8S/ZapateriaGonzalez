<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,rfc,nom,dir,tel,cd,cp,ntarj

rfc=Request.Form("RFC")
nom=request.Form("Nombre")
dir=Request.Form("Direccion")
tel=Request.Form("Telefono")
cd=request.Form("Ciudad")
cp=Request.Form("Codigo_Postal")
ntarj=Request.Form("No_Tarjeta")

set conn=Server.CreateObject("ADODB.connection")
Conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Zapateria\Personal.mdb"
Conn.execute "INSERT INTO Clientes(RFC,Nombre,Direccion,Telefono,Ciudad,Codigo_Postal,No_Tarjeta) values('"& rfc & "','"& nom & "','"& dir & "','"& tel & "','"& cd & "','"& cp & "','"& ntarj & "')"
Conn.close
set conn=nothing
response.redirect("gracias.html")

%>
</body>
</html>