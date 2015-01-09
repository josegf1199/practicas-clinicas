<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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


<title>PÁGINA AVISO POR TRABAJOS EN CURSO DEL DPTO. DE DESARROLLO</title>

<style type="text/css">

.Estilo1 
{
	font-family: Arial, Helvetica, sans-serif;
	color: #990000;
	font-weight: normal;
}

.Estilo2 
{
color: #990000; 
font-family: Arial, Helvetica, sans-serif;
}

body 
{
	background-color: #CCCCCC;
}


.powered1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color: #CCCCCC;
	color: #666666 ;
	font-weight: normal;

}

.powered2{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	background-color: #FFFFFF;
	color: #999999 ;
	font-weight: normal;

}

</style>


</head>

<body>



<table align="center" width="820" border="0">
  <tr>
    <td width="92"><img src="gif/construccion_1.gif" width="90" height="85" /></td>
    <td width="718"><p class="Estilo1">Esta p&aacute;gina esta temporalmente bloqueada por estar realizando trabajos de actualizaci&oacute;n y/o desarrollo en estos momentos.</p>
    <p><span class="Estilo1">Para cualquier aviso escriban a </span><span class="Estilo2"><a href="mailto:direcciondeemail@ugr.es">direcciondeemail@ugr.es</a></span></p></td>
  </tr>
  
  <tr>
    <td>&nbsp;</td>
    <td><p>&nbsp;</p>
    <p class="Estilo1">Realizaremos dichos trabajos en la mayor brevedad posible para así poder reanudar el servicio.<br /><br /><br /><br /><b>Disculpen las molestias...</b></td>
  </tr>
</table>


<br /><br />

<table border="0" align="center" width="266">
<tr>
<td></td>
<td width="232"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton2').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton2').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo2').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo2').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo2" style="visibility:hidden;"/></a></td>
<td width="232"  align="center" class="powered1" id="idtextoregresarboton2" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>


</body>
</html>
