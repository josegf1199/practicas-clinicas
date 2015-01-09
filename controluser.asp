<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta http-equiv="Content-Language" content="es">


<%
Response.ContentType="text/html"
Response.Charset="UTF-8"
Session.CodePage=65001
%>

<!--  #INCLUDE FILE="adovbs.inc"-->

<title>CONTROL DE AUTETIFICACIÓN DEL ADMINISTRADOR</title>

</head>

<%
	Dim ObjConn
	Dim ObjRs
	Dim BusquedaUserSql
	%>

<body>
	
<%
	BusquedaUserSql="SELECT * FROM tblAdministradores WHERE LOGIN_USUARIO='" & Request.Form("nameuser") & "' AND PASSWORD_USUARIO='" & Request.Form("codepassword") & "'"
	
	set ObjConn = Server.CreateObject("ADODB.Connection") 
	set ObjRs = Server.CreateObject("ADODB.Recordset")
	
	
	ObjConn.Open "dsndbgpc"
	
	Set ObjRs=ObjConn.Execute(BusquedaUserSql)
	
	If (not ObjRs.EOF) then 
			Session("autentificado")="si"
			Response.Redirect("altapracticaclinica.asp")
	Else
	
			Response.Redirect("index.asp?errorusuario=si")
	End if
	
	ObjRs.Close
	Set ObjRs = Nothing

	ObjConn.Close
	Set ObjConn = Nothing 
%>
    
    
</body>

</html>
