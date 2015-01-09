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


<title>AUTENTIFICACIÓN DE ADMINISTRADOR A LA WEB_APPLICATION GPC v1.0</title>


<link rel="stylesheet" type="text/css" href="stiloscss/estilos.css">


<script src="scripts/AC_RunActiveContent.js" type="text/javascript"></script>



<script type="text/javascript">

function stopRKey(evt) 
					
									{
												var evt = (evt) ? evt : ((event) ? event : null);
												var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
												
												if ((evt.keyCode == 13) && (node.type=="text")) 
												
															{
																		return false;
															}
									}
					
									document.onkeypress = stopRKey;


function fecha()
{
	fecha = new Date()
	mes = fecha.getMonth()
	diaMes = fecha.getDate()
	diaSemana = fecha.getDay()
	anio = fecha.getFullYear()
	dias = new Array('Domingo','Lunes','Martes','Miercoles','Jueves','Viernes','Sábado')
	meses = new Array('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre')
	document.write('<span id="fecha">')
	document.write (dias[diaSemana] + ", " + diaMes + " de " + meses[mes] + " de " + anio)
	document.write ('</span>')
	
}


function hora()
{
	var fecha = new Date()
	var hora = fecha.getHours()
	var minuto = fecha.getMinutes()
	var segundo = fecha.getSeconds()
	if (hora < 10) {hora = "0" + hora}
	if (minuto < 10) {minuto = "0" + minuto}
	if (segundo < 10) {segundo = "0" + segundo}
	var hora_completa = hora + ":" + minuto + ":" + segundo
	document.getElementById('hora').firstChild.nodeValue = hora_completa
	tiempo = setTimeout('hora()',1000)
}


function inicio()
{
	document.write('<span id="hora">')
	document.write ('000000</span>')
	hora()
}






window.onload = function()

{

	document.formulariocontroluser.FechaSistema.value = (dias[diaSemana] + ", " + diaMes + " de " + meses[mes] + " de " + anio);

	var fecha = new Date()
	var hora = fecha.getHours()
	var minuto = fecha.getMinutes()
	var segundo = fecha.getSeconds()
	if (hora < 10) {hora = "0" + hora}
	if (minuto < 10) {minuto = "0" + minuto}
	if (segundo < 10) {segundo = "0" + segundo}
	var hora_completa = hora + ":" + minuto + ":" + segundo
	document.getElementById('hora').firstChild.nodeValue = hora_completa
	tiempo = setTimeout('hora()',1000)

	document.formulariocontroluser.HoraSistema.value = hora_completa;
	
	document.getElementById("idnameuser").focus();
}



</script>


</head>



<body leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0"  bgcolor="#FFFFFF">


<table width="400" align="center" border="0">
    <tr>
      <td width="243" class="campotextofechahora"><div align="left"><script>fecha()</script></div></td>
      <td width="17" ><input type="hidden" name="FechaSistema" id="idfechasys" size=75  value=""/></td>
      
      <td width="105" class="campotextofechahora"><div align="left"><script>inicio()</script></div></td>
      <td width="17" ><input type="hidden" name="HoraSistema" id="idhorasys" value=""/></td>
  </tr>
</table>


<table width="1036" align="center" bgcolor="#CCCCCC" >
<tr>
<td><center><h1>Autentificación del Administrador para el Acceso a la Web_Application GPC v1.0</h1></center></td>
</tr>
</table>


    
<table align="center">
<tr>
<td><script type="text/javascript">

AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','791','height','569','src','flash/f_fmg_2','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','flash/f_fmg_2' ); 

</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="791" height="569">
      <param name="movie" value="flash/f_fmg_2.swf"/>
      <param name="quality" value="high"/>
      <embed src="flash/f_fmg_2.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="791" height="569"></embed>
    </object></noscript></td>
</tr>
</table>
	


<form name="formulariocontroluser"  method="post" action="controluser.asp" autocomplete="off">

		
        
<table align="center" width="400" cellspacing="2" cellpadding="2" border="1">


<tr>
	  		<%
            If Request.QueryString("errorusuario")<>"si" then
            %>
            			<td colspan="2" align="center" bgcolor="#CCCCCC">INTRODUZCA SU LOGIN Y PASSWORD DE ACCESO</td>
            <%
            Else
            %>
            			<td colspan="2" align="center" bgcolor="#FF0000"><span style="color:#FFFFFF"><b>ACCESO DENEGADO</b></span></td>
            <%
            End if
            %>
</tr>
            
                        
<tr>
      <td align="right">LOGIN USUARIO: </td>
      <td><input type="text" name="nameuser" id="idnameuser" size="8" maxlength="8" /></td>
</tr>
                        
<tr>
    <td align="right">PASSWORD: </td>
    <td><input type="password" name="codepassword" id="idcodepassword" size="8" maxlength="8" /></td>
</tr>
			
<tr>
	<td colspan="2" align="center"><input type="submit" value="Aceptar" /></td>
</tr>
			
</table>


<br/><br/>

<table border="0" align="center" width="900">
            <tr>
                        <td align="center" class="powered1">Powered by José Emilio Salvador Concepción</td>
            </tr>
           
            <tr>
                        <td align="center" class="powered2">© Universidad de Granada. Granada, 2014</td>
            </tr>
            <tr>
                        <td align="center" class="powered2">© SAS. Granada, 2014</td>
            </tr>
            <tr>
                        <td align="center" class="powered2">© José Gutiérrez-Fernández y José Emilio Salvador Concepción. Granada, 2014.</td>
            </tr>



</table>



<table width="513" align="center" border="0">
<tr>
			<td width="267"  align="center"><a href="http://www.ugr.es/" target="_blank"><img src="pngs/logo_ugr_21_mini.png"/></a></td>
            <td width="236"  align="center"><a href="http://www.ugr.es/~facmed/" target="_blank"><img src="gif/logoMedicina_21_mini.gif"/></a></td>
</tr>
</table>        




</form>

</body>
</html>
