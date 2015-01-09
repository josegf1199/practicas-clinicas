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

<link href="stiloscss/stilo1.css" rel="stylesheet" type="text/css" />

<style type="text/css">

.textocabecera1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #006666 ;
	*border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}

.textomenuinicio1 {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	*background-color:#CCCCCC;
	*border: 2px solid #505027;
	color: #000000;
	font-weight:normal;
}

.textotitulocabecera {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color: #CCCC99;
	border: 1px solid #000000;
	font-weight: normal;
}

.textotitulocampo1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #95C98B;
	border: 1px solid #000000;
	font-weight: bold;
}

.textotitulocampo1espfree {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #99CC66;
	border: 1px solid #000000;
	font-weight: bold;
}

. textotituloactivarbotongrabar{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	*background-color: #CCCC99;
	*border: 1px solid #000000;
	font-weight: bold;
}

.textotitulocampo2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 13px;
	background-color: #FFFFFF; 
	border: none;
}

.textotitulocampo3 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color:  #99CC66;
 	color: #000000;
  	font-weight:bold; 
 	*border: 2px solid #505027;
	
 }

.textotitulosubcabeceras1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	background-color: #009999;
 	color: #FFFFFF;
  	font-weight: normal; 
 	*border: 2px solid #505027;
 }
 
 
 .textotitulosubcabecerasinfo{
 
 	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
 	color:#333333;
  	font-weight: normal; 
 	border: 1px solid #999999; 
  
 }
 
 .textotitulosubcabecerasinfo11{
 
 	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
 	color: #FF0000;
  	font-weight: bold; 
 	*border: 1px solid #999999; 
  
 }


textotituloaccionborrarcompactar{
 
 	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
 	color:#000000;
  	font-weight: bold; 
  
 }

.campotexto1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 19px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform: uppercase;
}

.campotextofechahora

{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:inherit;
}

.campotextoestado1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #FF0000;
	font-weight:bold;
	text-transform: uppercase;
}


.campotextoestado2{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #009900;
	font-weight:bold;
	text-transform: uppercase;
}


.campoemail{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform: lowercase;
}

.stilocombobox {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color:#666666;
	border: 1px solid #000000;
	color: #FFFFFF;
	font-weight: normal;
}


.cabeceraciclos
{
		font-family:Arial, Helvetica, sans-serif;
		font-size: 14px;
		font-weight:bold;
		color:#006699;
}


.textocirugia 
{
	color: #000000;
    font-family:Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #CCCCCC;
	border: 1px solid #000000;
	*font-weight: normal;
	font-weight:bold;
}


#bloquedatos
{
		font-family:Arial, Helvetica, sans-serif;
		font-size: 14px;
		font-weight:bold;
		
		color:#006699;
		border: 1px #solid #ff0000;
		display:block;
		
		*display: none;
}


.powered1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color: #FFFFFF;
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

.powered3{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #FFFFFF;
	color: #000000 ;
	font-weight: normal;

}



#fecha
 {
	
/* width:80px;
	font-family: Tahoma, Verdana, Arial, sans-serif;
	font-size: .8em;
	color: #996600;
	background : #FFE3D7;
	text-align: center; 
*/
	
	font-family: Arial, Helvetica, sans-serif;
	font-size: 17px;
	background-color:#CCCCCC;
	* border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:inherit;
	text-align: center; 
	
	}

#hora 
{
/*	width:100px;
	font-family: Tahoma, Verdana, Arial, sans-serif;
	font-size: .9em;
	color: #996600;
	background : #FFE3D7;
	text-align: center;
*/	

	font-family: Arial, Helvetica, sans-serif;
	font-size: 17px;
	background-color:#CCCCCC;
	* border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:lowercase;

}


</style>


<script type="text/javascript">


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



//function habilitabotonactualizarficha()
//{
	//if(document.formularioaltapracticaclinica.AceptaFormularioOk.checked == true)            
		//{
			//document.getElementById("idenviardatos").style="visibility:visible";
		//} 
	//else 
		//{
			//document.getElementById("idenviardatos").style="visibility:hidden";
		//}
//}					
				


window.onload = function()

{

	document.formularioaltapracticaclinica.FechaAsignacionPracticas.value = (dias[diaSemana] + ", " + diaMes + " de " + meses[mes] + " de " + anio);

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

	document.formularioaltapracticaclinica.HoraAsignacionPracticas.value = hora_completa;
	
	//document.getElementById("iddnialumnoa").focus();
	
}



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
	


function asignarplazas(NPlazasParaTodasLasEspecialidadesC1C2)

{

var vplz=NPlazasParaTodasLasEspecialidadesC1C2;

document.formadmonplazasesp.NPlazasGeneralDigestivaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasGeneralDigestivaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasCardiacaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasCardiacaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasToracicaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasToracicaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMaxilofacialCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMaxilofacialCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasPlasticaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasPlasticaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasVascularCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasVascularCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeurocirugiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeurocirugiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasTraumatologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasTraumatologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasUrologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasUrologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasPediatricaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasPediatricaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadACiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadACiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadBCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadBCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadCCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadCCiclo2.value=vplz;




document.formadmonplazasesp.NPlazasHematologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasHematologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeumologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeumologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasCardiologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasCardiologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasDigestivoCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasDigestivoCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNefrologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNefrologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMedicinaInternaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMedicinaInternaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEndocrinologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEndocrinologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasReumatologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasReumatologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasOncologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasOncologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeurologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeurologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasInfecciososCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasInfecciososCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMedicinaIntensivaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMedicinaIntensivaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadDCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadDCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadECiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadECiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadFCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadFCiclo2.value=vplz;


}


function asignarplazasA0()

{

var vplz=0;

document.formadmonplazasesp.NPlazasGeneralDigestivaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasGeneralDigestivaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasCardiacaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasCardiacaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasToracicaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasToracicaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMaxilofacialCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMaxilofacialCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasPlasticaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasPlasticaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasVascularCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasVascularCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeurocirugiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeurocirugiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasTraumatologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasTraumatologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasUrologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasUrologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasPediatricaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasPediatricaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadACiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadACiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadBCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadBCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadCCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadCCiclo2.value=vplz;




document.formadmonplazasesp.NPlazasHematologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasHematologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeumologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeumologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasCardiologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasCardiologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasDigestivoCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasDigestivoCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNefrologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNefrologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMedicinaInternaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMedicinaInternaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEndocrinologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEndocrinologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasReumatologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasReumatologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasOncologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasOncologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasNeurologiaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasNeurologiaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasInfecciososCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasInfecciososCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasMedicinaIntensivaCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasMedicinaIntensivaCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadDCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadDCiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadECiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadECiclo2.value=vplz;

document.formadmonplazasesp.NPlazasEspecialidadFCiclo1.value=vplz;
document.formadmonplazasesp.NPlazasEspecialidadFCiclo2.value=vplz;


}












</script>


<title>GESTIÓN ADMINISTRATIVA SOBRE LA BASE DE DATOS</title>

</head>

<body leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0"  bgcolor="#FFFFFF">

<%
	
	Dim ObjConn	
	Dim ObjRs    	
	
if Request.Form="" then 
%>

<form name="formadmonplazasesp" method="post" action="iniplazasespecialidades.asp" autocomplete="off">


<p>


<%
Dim SqlContarRegistros
Dim TotalRegistrosExistentesTabla

SqlContarRegistros = "SELECT * FROM tbl_Alumnos"

Set ObjConn = Server.CreateObject("ADODB.Connection")
Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjConn.Open "dsndbgpc"

ObjRs.CursorType = 1

ObjRs.Open SqlContarRegistros, ObjConn

TotalRegistrosExistentesTabla=ObjRs.RecordCount


ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>

<table width="800" border="0" align="center">
<tr>
    <td width="800" align="center"><img src="jpg/tools_db_109_x_128.jpg" width="109" height="128" /></td>
</tr>
</table>


<br />


<table width="1200" align="center" border="0">
<tr>
<td  class="textocabecera1" align="center">GESTIÓN ADMINISTRATIVA SOBRE LA BASE DE DATOS EN REFERENCIA A LAS PLAZAS DISPONIBLES EN LAS ESPECIALIDADES QUIRÚRGICAS Y MÉDICAS EXISTENTES PARA CADA INICIO DE PERIODO DE SELECCIÓN DE PRÁCTICAS CLÍNICAS.</td>
</tr>
</table>


<br/>


<!--
<table width="900" align="center" border="0">
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1"/></a></td>
<td width="1016" height="12" align="center" class="textotitulosubcabecerasinfo">RELACIÓN DE ESPECIALIDADES DEL PRIMER Y SEGUNDO CUATRIMESTRE, DONDE SE MUESTRA LAS PLAZAS DISPONIBLES POR ESPECIALIDAD Y DONDE PODRÁ PONER A CERO DICHAS ESPECIALIDADES, MODIFICAR EL NÚMERO PLAZAS POR ESPECIALIDAD DE FORMA INDIVIDUAL O GLOBAL.</td>
</tr>
</table>
-->

<br/>

<table align="center" width="900" border="0">
<tr>
    <td width="93"><img src="gif/aviso_1.gif" width="128" height="128" /></td>
    <td width="900" class="textotitulosubcabecerasinfo11" align="justify"><h1>¡¡¡AVISO!!!</h1>CUALQUIERA DE LAS DOS OPCIONES SIGUIENTES QUE ESCOJA Y EJECUTE REALIZARA EL BORRADO DE LOS DATOS DE ALUMNOS/AS DE FORMA IRRECUPERABLE EN LA BASE DE DATOS. ESTE PROCESO SE REALIZA PREVIO A LA ENTRADA DE DATOS DE UNA NUEVA PROMOCIÓN.</td>
</tr>
</table>


<br/>


<table width="1000" align="center" border="0">
<tr>
<td height="12" align="center" class="textocabecera1">QUE PROCESO DESEA SELECCIONAR:</td>
</tr>
</table>


<table width="1043" align="center" border="0">
<tr>
<td width="20" height="93"><input type="radio"  name="opciondelomovecompacrepair" id="idselectiondelandcompacrepair" value="1" checked="checked" onclick="document.getElementById('opciondcr').style='display:block;' ; document.getElementById('opcionmcr').style='display:none;'"/></td><td width="481" class="textotitulosubcabecerasinfo" align="justify"><p>Preparación de la base de datos para un nuevo ciclo de selección de prácticas clínicas y <b>BORRADO DEFINITIVO de datos de alumnos/as ASÍ COMO SUS DOCUMENTOS PDF de selección de prácticas clínicas.</b></p>
  <p><b>(ESTE PROCESO SOLO SE REALIZA SI SE HAN INTRODUCIDO PREVIAMENTE DATOS DE ALUMNOS/AS EN LA BASE DE DATOS).</b></p></td>

<td width="20" height="93"><input type="radio" name="opciondelomovecompacrepair" id="idselectionmoveandcompacrepair" value="2" onclick="document.getElementById('opciondcr').style='display:none;' ; document.getElementById('opcionmcr').style='display:block;'"/></td>
<td width="504" class="textotitulosubcabecerasinfo" align="justify"><p>Preparación de la base de datos para un nuevo ciclo de selección de prácticas clínicas y <b>BORRADO DEFINITIVO de datos de alumnos/as y CONSERVACIÓN HISTORICA DE SUS DOCUMENTOS PDF de selección de prácticas clínicas. </b></p>
  <p><b>(ESTE PROCESO SOLO SE REALIZA SI SE HAN INTRODUCIDO PREVIAMENTE DATOS DE ALUMNOS/AS EN LA BASE DE DATOS).</b></p></td>

<!--<td width="20"><input type="checkbox" name="delcompacrepair" id="idselectiondelandcompac" value="1"/></td>
<td width="20"><input type="checkbox" name="movecompacrepair" id="idselectionmoveandcompac" value="2"/></td>-->
</tr>
</table>



<div id="opciondcr" style="display:block;" >

<table width="551" align="center" border="0">
			
            <tr>

            <td width="545" align="center"><img src="pngs/delcompact_&_repair_256_x_256.png" id="idlogodelcompact&repair"/></td>
            <!--<td width="545" align="center"><img src="pngs/delcompact_&_repair_256_x_256_bn.png" id="idlogodelcompact&repairbn"/></td>-->
            
            <!--<td align="center"><img src="pngs/compact_repair_&_move_256_x_256.png" id="idlogocompactrepairandmove"/></td>-->
            <td align="center"><img src="pngs/compact_repair_&_move_256_x_256_bn.png" id="idlogocompactrepairandmovebn"/></td>
    		</tr>
            
            
            
            <tr>
            
    		<td align="center" ><a href="delcompactandrepairdb.asp" target="_blank"><img src="pngs/button2.png" align="center" id="idbotonejecutarproceso_dcr_on"/></a></td>
            <td align="center" ><img src="pngs/button2_bn.png" align="center" id="idbotonejecutarproceso_dcr_off"/></td>
            
          
     		<!--<td align="center" ><a href="movecompactandrepairdb.asp" target="_blank"><img src="pngs/button2.png" align="center" id="idbotonejecutarproceso_mcr_on"/></a></td>-->
     		<!--<td align="center" ><a href="movecompactandrepairdb.asp" target="_blank"><img src="pngs/button2_bn.png" align="center" id="idbotonejecutarproceso_mcr_off"/></a></td>-->
 			
		
            <!--<td width="728" class="textotituloaccionborrarcompactar" ><a href="delcompactandrepairdb.asp" target="_blank"></a></td>-->
			
            </tr>
</table>
</div>


<div id="opcionmcr"  style="display: none;">

<table width="551" align="center" border="0">
			
            <tr>
            

            <!--<td width="545" align="center"><img src="pngs/delcompact_&_repair_256_x_256.png" id="idlogodelcompact&repair"/></td>-->
            <td width="545" align="center"><img src="pngs/delcompact_&_repair_256_x_256_bn.png" id="idlogodelcompact&repairbn"/></td>
            
            <td align="center"><img src="pngs/compact_repair_&_move_256_x_256.png" id="idlogocompactrepairandmove"/></td>
            <!--<td align="center"><img src="pngs/compact_repair_&_move_256_x_256_bn.png" id="idlogocompactrepairandmovebn"/></td>-->
  
    		</tr>
            
            
            
            <tr>
            
            <td align="center" ><img src="pngs/button2_bn.png" align="center" id="idbotonejecutarproceso_dcr_off"/></td>
     		<td align="center" ><a href="movecompactandrepairdb.asp" target="_blank"><img src="pngs/button2.png" align="center" id="idbotonejecutarproceso_dcr_on"/></a></td>
           
          
     		<!--<td align="center" ><a href="movecompactandrepairdb.asp" target="_blank"><img src="pngs/button2.png" align="center" id="idbotonejecutarproceso_mcr_on"/></a></td>-->
     		<!--<td align="center" ><a href="movecompactandrepairdb.asp" target="_blank"><img src="pngs/button2_bn.png" align="center" id="idbotonejecutarproceso_mcr_off"/></a></td>-->
 			
		
            <!--<td width="728" class="textotituloaccionborrarcompactar" ><a href="delcompactandrepairdb.asp" target="_blank"></a></td>-->
			
            </tr>
</table>
</div>






<br />


<!--
<table width="600" align="center" border="0">
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1"/></a></td>
<td width="1016" height="12" align="center" class="textotitulosubcabecerasinfo">PROCESO DE BORRADO DE TODOS LOS REGISTROS DEL FICHERO DE ALUMNOS/AS, COMPRESIÓN Y REPARACIÓN DE LA BASE DE DATOS PARA SU REUTILIZACIÓN PARA UN NUEVO CICLO DE SELECCIÓN DE ESPECIALIDADES CLÍNICAS POR EL ALUMNADO.</td>
</tr>
</table>

<br/><br/>
-->

<br />


<table width="1100" align="center" border="0">
<tr>
<td width="1016" height="12" align="center" class="textotitulosubcabeceras1">ÁREA DE MODIFICACIÓN Y ACTUALIZACIÓN DE PLAZAS DISPONIBLES EN LAS ESPECIALIDADES CORRESPONDIENTES A LOS SERVICIOS QUIRÚRGICOS.</td>
</tr>
</table>


<table width="1100" align="center" border="0" cellspacing="0">

<tr>
<td width="378" height="12" align="center" class="textotitulocampo3">1º CICLO [PRIMER CUATRIMESTRE]</td>
<td width="412" height="12" align="center" class="textotitulocampo3">2º CICLO [SEGUNDO CUATRIMESTRE]</td>
</tr>



<tr>
<td width="378" height="12" align="center" class="textotitulocampo3"></td>
<td width="412" height="12" align="center" class="textotitulocampo3"></td>
</tr>


</table>


<table width="1100" align="center" border="0" cellspacing="0">

<tr>
<td width="188" height="12" align="center" class="textotitulocampo3"></td>
<td width="102" height="12" align="center" class="textotitulocampo3"></td>
<td width="121" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="86" height="12" align="center" class="textotitulocampo3"></th>
<td width="69" height="12" align="center" class="textotitulocampo3"></td>
<td width="169" height="12" align="left" class="textotitulocampo3"></td>
<td width="129" height="12" align="center" class="textotitulocampo3"></td>
<td width="143" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="75" height="12" align="center" class="textotitulocampo3"></th></tr>

<tr>
<td width="188" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="102" height="12" align="center" class="textotitulocampo3">MODIFICAR PLAZAS</td>
<td width="121" height="12" align="center" class="textotitulocampo3">DISPONIBLES ACTUALES</td>
<td width="86" height="12" align="center" class="textotitulocampo3">
<td width="69" height="12" align="center" class="textotitulocampo3"></td>
<td width="169" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="129" height="12" align="center" class="textotitulocampo3">MODIFICAR PLAZAS</td>
<td width="143" height="12" align="center" class="textotitulocampo3">DISPONIBLES ACTUALES</td>
<td width="75" height="12" align="center" class="textotitulocampo3"></td>
</tr>

</table>


<br />


<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
'ObjConn.Open "dsndbgpc"
ObjConn.Open="driver={Microsoft Access Driver (*.mdb)}; dbq=c:\inetpub\wwwroot\gpc\data\dbgpc.mdb"


Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn, 3, 3

Dim SQLGDIG1
SQLGDIG1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='GDIG1'"

Set ObjRs=ObjConn.Execute(SQLGDIG1)
%>

<table width="1100" align="center" border="0">

<tr>
<td width="185" height="12" align="right" class="textotitulocampo1">General-Digestiva:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasGeneralDigestivaCiclo1" value=""  id="idnplazasgeneraldigestivac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 


Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn, 3, 3

Dim SQLGDIG2
SQLGDIG2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='GDIG2'"

Set ObjRs=ObjConn.Execute(SQLGDIG2)
%>


<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>

<td width="166" height="12" align="right" class="textotitulocampo1">General-Digestiva:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasGeneralDigestivaCiclo2" value=""  id="idnplazasgeneraldigestivac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLCARD1
SQLCARD1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='CARD1'"

Set ObjRs=ObjConn.Execute(SQLCARD1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Cardiaca:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasCardiacaCiclo1" value=""  id="idnplazascardiacac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLCARD2
SQLCARD2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='CARD2'"

Set ObjRs=ObjConn.Execute(SQLCARD2)
%>


<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>

<td width="166" height="12" align="right" class="textotitulocampo1">Cardiaca:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasCardiacaCiclo2" value=""  id="idnplazascardiacac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>
<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>






















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLTORA1
SQLTORA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='TORA1'"

Set ObjRs=ObjConn.Execute(SQLTORA1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Torácica:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasToracicaCiclo1" value=""  id="idnplazastoracicac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLTORA2
SQLTORA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='TORA2'"

Set ObjRs=ObjConn.Execute(SQLTORA2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>

<td height="12" class="textotitulocampo1" align="right">Torácica:</td>
<td width="127" height="12" align="center"><input type='text' name="NPlazasToracicaCiclo2" value=""  id="idnplazastoracicac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMAXI1
SQLMAXI1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MAXI1'"

Set ObjRs=ObjConn.Execute(SQLMAXI1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Maxilofacial:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasMaxilofacialCiclo1" value=""  id="idnplazasmaxilofacialc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMAXI2
SQLMAXI2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MAXI2'"

Set ObjRs=ObjConn.Execute(SQLMAXI2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Maxilofacial:</td>
<td width="127" height="12" align="center"><input type='text' name="NPlazasMaxilofacialCiclo2" value=""  id="idnplazasmaxilofacialc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLPLAS1
SQLPLAS1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='PLAS1'"

Set ObjRs=ObjConn.Execute(SQLPLAS1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Plástica:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasPlasticaCiclo1" value=""  id="idnplazasidplasticac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLPLAS2
SQLPLAS2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='PLAS2'"

Set ObjRs=ObjConn.Execute(SQLPLAS2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>

<td height="12" class="textotitulocampo1" align="right">Plástica:</td>
<td width="127" height="12" align="center"><input type='text' name="NPlazasPlasticaCiclo2" value=""  id="idnplazasplasticac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLVASC1
SQLVASC1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='VASC1'"

Set ObjRs=ObjConn.Execute(SQLVASC1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Vascular:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasVascularCiclo1" value=""  id="idnplazasvascularc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLVASC2
SQLVASC2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='VASC2'"

Set ObjRs=ObjConn.Execute(SQLVASC2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Vascular:</td>
<td width="127" height="12" align="center"><input type='text' name="NPlazasVascularCiclo2" value=""  id="idnplazasvascularc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEUR1
SQLNEUR1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEUR1'"

Set ObjRs=ObjConn.Execute(SQLNEUR1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Neurocirugía:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasNeurocirugiaCiclo1" value=""  id="idnplazasneurocirugiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEUR2
SQLNEUR2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEUR2'"

Set ObjRs=ObjConn.Execute(SQLNEUR2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neurocirugía:</td>
<td width="127" height="12" align="center"><input type='text' name="NPlazasNeurocirugiaCiclo2" value=""  id="idnplazasneurocirugiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLTRAU1
SQLTRAU1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='TRAU1'"

Set ObjRs=ObjConn.Execute(SQLTRAU1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Traumatología:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasTraumatologiaCiclo1" value=""  id="idnplazastraumatologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLTRAU2
SQLTRAU2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='TRAU2'"

Set ObjRs=ObjConn.Execute(SQLTRAU2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Traumatología:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasTraumatologiaCiclo2" value=""  id="idnplazastraumatologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLUROL1
SQLUROL1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='UROL1'"

Set ObjRs=ObjConn.Execute(SQLUROL1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Urología:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasUrologiaCiclo1" value=""  id="idnplazasurologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLUROL2
SQLUROL2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='UROL2'"

Set ObjRs=ObjConn.Execute(SQLUROL2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Urología:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasUrologiaCiclo2" value=""  id="idnplazasurologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLPEDI1
SQLPEDI1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='PEDI1'"

Set ObjRs=ObjConn.Execute(SQLPEDI1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Pediátrica:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasPediatricaCiclo1" value=""  id="idnplazaspediatricac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLPEDI2
SQLPEDI2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='PEDI2'"

Set ObjRs=ObjConn.Execute(SQLPEDI2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Pediátrica:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasPediatricaCiclo2" value=""  id="idnplazaspediatricac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>





<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->

<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPA1
SQLESPA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPA1'"

Set ObjRs=ObjConn.Execute(SQLESPA1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - A:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasEspecialidadACiclo1" value=""  id="idnplazasespecialidadac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPA2
SQLESPA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPA2'"

Set ObjRs=ObjConn.Execute(SQLESPA2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - A:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasEspecialidadACiclo2" value=""  id="idnplazasespecialidadac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPB1
SQLESPB1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPB1'"

Set ObjRs=ObjConn.Execute(SQLESPB1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - B:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasEspecialidadBCiclo1" value=""  id="idnplazasespecialidadbc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPB2
SQLESPB2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPB2'"

Set ObjRs=ObjConn.Execute(SQLESPB2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - B:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasEspecialidadBCiclo2" value=""  id="idnplazasespecialidadbc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPC1
SQLESPC1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPC1'"

Set ObjRs=ObjConn.Execute(SQLESPC1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - C:</td>
<td width="98" height="12" align="center"><input type='text' name="NPlazasEspecialidadCCiclo1" value=""  id="idnplazasespecialidadc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="119" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPC2
SQLESPC2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPC2'"

Set ObjRs=ObjConn.Execute(SQLESPC2)
%>

<td width="86" height="12" align="left"></td>
<td width="67" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - C:</td>

<td width="127" height="12" align="center"><input type='text' name="NPlazasEspecialidadCCiclo2" value=""  id="idnplazasespecialidadc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="143" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>

</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>






</table>




<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->


<br />



<table width="1100" align="center" border="0">
<tr>
<td width="1016" height="12" align="center" class="textotitulosubcabeceras1">ÁREA DE MODIFICACIÓN Y ACTUALIZACIÓN DE PLAZAS DISPONIBLES EN LAS ESPECIALIDADES CORRESPONDIENTES A LOS SERVICIOS MÉDICOS.</td>
</tr>
</table>


<table width="1100" align="center" border="0" cellspacing="0">

<tr>
<td width="378" height="12" align="center" class="textotitulocampo3">1º CICLO [PRIMER CUATRIMESTRE]</td>
<td width="412" height="12" align="center" class="textotitulocampo3">2º CICLO [SEGUNDO CUATRIMESTRE]</td>
</tr>



<tr>
<td width="378" height="12" align="center" class="textotitulocampo3"></td>
<td width="412" height="12" align="center" class="textotitulocampo3"></td>
</tr>


</table>


<table width="1100" align="center" border="0" cellspacing="0">

<tr>
<td width="188" height="12" align="center" class="textotitulocampo3"></td>
<td width="102" height="12" align="center" class="textotitulocampo3"></td>
<td width="121" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="86" height="12" align="center" class="textotitulocampo3"></th>
<td width="69" height="12" align="center" class="textotitulocampo3"></td>
<td width="169" height="12" align="left" class="textotitulocampo3"></td>
<td width="129" height="12" align="center" class="textotitulocampo3"></td>
<td width="143" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="75" height="12" align="center" class="textotitulocampo3"></th></tr>

<tr>
<td width="188" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="102" height="12" align="center" class="textotitulocampo3">MODIFICAR PLAZAS</td>
<td width="121" height="12" align="center" class="textotitulocampo3">DISPONIBLES ACTUALES</td>
<td width="86" height="12" align="center" class="textotitulocampo3">
<td width="69" height="12" align="center" class="textotitulocampo3"></td>
<td width="169" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="129" height="12" align="center" class="textotitulocampo3">MODIFICAR PLAZAS</td>
<td width="143" height="12" align="center" class="textotitulocampo3">DISPONIBLES ACTUALES</td>
<td width="75" height="12" align="center" class="textotitulocampo3"></td>
</tr>

</table>


<br />



<table width="1100" align="center" border="0">


<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLHEMA1
SQLHEMA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='HEMA1'"

Set ObjRs=ObjConn.Execute(SQLHEMA1)
%>

<tr>
<td width="184" height="12" align="right" class="textotitulocampo1">Hematología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasHematologiaCiclo1" value=""  id="idnplazashematologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLHEMA2
SQLHEMA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='HEMA2'"

Set ObjRs=ObjConn.Execute(SQLHEMA2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td width="165" height="12" align="right" class="textotitulocampo1">Hematología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasHematologiaCiclo2" value=""  id="idnplazashematologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>








<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEUM1
SQLNEUM1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEUM1'"

Set ObjRs=ObjConn.Execute(SQLNEUM1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Neumología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasNeumologiaCiclo1" value=""  id="idnplazasneumologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEUM2
SQLNEUM2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEUM2'"

Set ObjRs=ObjConn.Execute(SQLNEUM2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neumología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasNeumologiaCiclo2" value=""  id="idnplazasneumologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>

















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLCGIA1
SQLCGIA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='CGIA1'"

Set ObjRs=ObjConn.Execute(SQLCGIA1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Cardiología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasCardiologiaCiclo1" value=""  id="idnplazascardiologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLCGIA2
SQLCGIA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='CGIA2'"

Set ObjRs=ObjConn.Execute(SQLCGIA2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Cardiología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasCardiologiaCiclo2" value=""  id="idnplazascardiologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLDIGE1
SQLDIGE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='DIGE1'"

Set ObjRs=ObjConn.Execute(SQLDIGE1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Digestivo:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasDigestivoCiclo1" value=""  id="idnplazasdigestivoc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLDIGE2
SQLDIGE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='DIGE2'"

Set ObjRs=ObjConn.Execute(SQLDIGE2)
%>


<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Digestivo:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasDigestivoCiclo2" value=""  id="idnplazasdigestivoc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEFR1
SQLNEFR1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEFR1'"

Set ObjRs=ObjConn.Execute(SQLNEFR1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Nefrología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasNefrologiaCiclo1" value=""  id="idnplazasnefrologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEFR2
SQLNEFR2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEFR2'"

Set ObjRs=ObjConn.Execute(SQLNEFR2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Nefrología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasNefrologiaCiclo2" value=""  id="idnplazasnefrologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMINT1
SQLMINT1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MINT1'"

Set ObjRs=ObjConn.Execute(SQLMINT1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Medicina Interna:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasMedicinaInternaCiclo1" value=""  id="idnplazasmedicinainternac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>
<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMINT2
SQLMINT2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MINT2'"

Set ObjRs=ObjConn.Execute(SQLMINT2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Medicina Interna:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasMedicinaInternaCiclo2" value=""  id="idnplazasmedicinainternac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLENDO1
SQLENDO1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ENDO1'"

Set ObjRs=ObjConn.Execute(SQLENDO1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Endocrinología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasEndocrinologiaCiclo1" value=""  id="idnplazasendocrinologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLENDO2
SQLENDO2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ENDO2'"

Set ObjRs=ObjConn.Execute(SQLENDO2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Endocrinología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasEndocrinologiaCiclo2" value=""  id="idnplazasendocrinologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLREUM1
SQLREUM1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='REUM1'"

Set ObjRs=ObjConn.Execute(SQLREUM1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Reumatología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasReumatologiaCiclo1" value=""  id="idnplazaseumatologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLREUM2
SQLREUM2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='REUM2'"

Set ObjRs=ObjConn.Execute(SQLREUM2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Reumatología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasReumatologiaCiclo2" value=""  id="idnplazasreumatologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLONCO1
SQLONCO1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ONCO1'"

Set ObjRs=ObjConn.Execute(SQLONCO1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Oncología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasOncologiaCiclo1" value=""  id="idnplazasoncologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLONCO2
SQLONCO2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ONCO2'"

Set ObjRs=ObjConn.Execute(SQLONCO2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Oncología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasOncologiaCiclo2" value=""  id="idnplazasoncologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNGIA1
SQLNGIA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NGIA1'"

Set ObjRs=ObjConn.Execute(SQLNGIA1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Neurología:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasNeurologiaCiclo1" value=""  id="idnplazasneurologiac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNGIA2
SQLNGIA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NGIA2'"

Set ObjRs=ObjConn.Execute(SQLNGIA2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neurología:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasNeurologiaCiclo2" value=""  id="idnplazasneurologiac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLINFE1
SQLINFE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='INFE1'"

Set ObjRs=ObjConn.Execute(SQLINFE1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Infecciosos:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasInfecciososCiclo1" value=""  id="idnplazasinfecciososc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLINFE2
SQLINFE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='INFE2'"

Set ObjRs=ObjConn.Execute(SQLINFE2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Infecciosos:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasInfecciososCiclo2" value=""  id="idnplazasinfecciososc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMSIV1
SQLMSIV1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MSIV1'"

Set ObjRs=ObjConn.Execute(SQLMSIV1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Medicina Intensiva:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasMedicinaIntensivaCiclo1" value=""  id="idnplazasmedicinaintensivac1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>





<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMSIV2
SQLMSIV2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MSIV2'"

Set ObjRs=ObjConn.Execute(SQLMSIV2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Medicina Intensiva:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasMedicinaIntensivaCiclo2" value=""  id="idnplazasmedicinaintensivac2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->

<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPD1
SQLESPD1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPD1'"

Set ObjRs=ObjConn.Execute(SQLESPD1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - D:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasEspecialidadDCiclo1" value=""  id="idnplazasespecialidaddc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPD2
SQLESPD2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPD2'"

Set ObjRs=ObjConn.Execute(SQLESPD2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - D:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasEspecialidadDCiclo2" value=""  id="idnplazasespecialidaddc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPE1
SQLESPE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPE1'"

Set ObjRs=ObjConn.Execute(SQLESPE1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - E:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasEspecialidadECiclo1" value=""  id="idnplazasespecialidadec1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPE2
SQLESPE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPE2'"

Set ObjRs=ObjConn.Execute(SQLESPE2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - E:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasEspecialidadECiclo2" value=""  id="idnplazasespecialidadec2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>


















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPF1
SQLESPF1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPF1'"

Set ObjRs=ObjConn.Execute(SQLESPF1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - F:</td>
<td width="101" height="12" align="center"><input type='text' name="NPlazasEspecialidadFCiclo1" value=""  id="idnplazasespecialidadfc1" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="116" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPF2
SQLESPF2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPF2'"

Set ObjRs=ObjConn.Execute(SQLESPF2)
%>

<td width="85" height="12" align="left"></td>
<td width="70" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - F:</td>

<td width="131" height="12" align="center"><input type='text' name="NPlazasEspecialidadFCiclo2" value=""  id="idnplazasespecialidadfc2" size="10" maxlength="3" onfocus='document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>

<td width="139" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<td width="71" height="12" align="left"></td>
</tr>
</table>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->


















<br /><br />


<div id="mensajeaceptacionformulario" style="display: none">
    <table align="center" width="800" border="0">
        <tr>
            <td align="center"><input type="checkbox" value="0" id="idaceptaformulario" name="AceptaFormularioOk" onclick="habilitabotonactualizarficha()"/><label for="idaceptaformulario"></label></td>
            <td align="left">Si esta conforme con todos los datos introducidos y son correctos en el formulario que acaba de rellenar, marca esta casilla y pulse el siguiente botón  <b>'GRABAR ESTE FORMULARIO DEL ALUMNO/A EN LA BASE DE DATOS E IMPRIMIR DOCUMENTO PDF' </b>y generar su informe de asignación de prácticas clínicas PDF para imprimir.</td>
      </tr>
    </table>
</div>   




<table width="1100" align="center" border="0">
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1_fijo"/></a></td>
<td width="1016" height="12" align="center" class="powered3"><b>USTED PODRÁ: </b><img src="pngs/sfera_1_16_x_16.png" id="idlogoinfo1_apartado"/>  Definir en cada especialidad el número de plazas que desee otorgar.  <img src="pngs/sfera_1_16_x_16.png" id="idlogoinfo1_apartado"/>  Asignar una misma cantidad de plazas para todas las especialidades de forma general. Poner a '0' todas las plazas de una forma general. <img src="pngs/sfera_1_16_x_16.png" id="idlogoinfo1_apartado"/>  Y una vez definidas la plazas como usted desee pulsar el botón '<b>ACTUALIZAR LAS PLAZAS DESIGNADAS EN LOS CAMPOS CORRESPONDIENTES A CADA ESPECILIDAD EN LA BASE DE DATOS</b>'.</td>
</tr>
</table>



<br />



<table width="500" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>
      <td width="500" height="12" align="left" class="powered3">Introduzca un número de plazas general a todas las especialidades: </td> 
     </tr>
</table> 
      

<table width="500" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>      
      <td width="132" height="12" align="left"><input type='text' name="NPlazasParaTodasLasEspecialidadesC1C2" value=""  id="idnplazasparatodaslasespecialidadesc1c2" size="5" maxlength="3"/></td>
      <td width="368" align="right" valign="top"><input type="button" name="enviardatos" id="idenviardatos" value="ASIGNACIÓN DE 'n' PLAZAS A TODAS LAS ESPECIALIDADES"  onclick='asignarplazas(document.formadmonplazasesp.NPlazasParaTodasLasEspecialidadesC1C2.value); document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>
    </tr> 
</table>

<br />

<table width="500" border="0" cellpadding="0" cellspacing="0" align="center">
        
     <tr>
     <td width="500"></td>
      <td width="500" align="right" valign="top"><input type="button" name="enviardatos" id="idenviardatos" value="PONER A  '0' PLAZAS A TODAS LAS ESPECIALIDADES"  onclick='asignarplazasA0(); document.getElementById("idbotonactualizarplazas").style.display = "block";'/></td>
    </tr>
</table>


<br />


<div id="idbotonactualizarplazas" style="display: none">
<table width="800" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>
    <td width="800" align="center" valign="top"><input type="Submit" name="enviardatos" id="idenviardatos" value="ACTUALIZAR LAS PLAZAS DESIGNADAS EN LOS CAMPOS CORRESPONDIENTES A CADA ESPECILIDAD EN LA BASE DE DATOS" /></td>
	</tr>
</table>
</div>



<br /><br />


<table border="0" align="center" width="397">
<tr>
<td></td>
<td width="361"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton1').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton1').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo1').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo1').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="48"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1" style="visibility:hidden;"/></a></td>
<td width="361"  align="center" class="powered1" id="idtextoregresarboton1" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>



<br/>

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


<br /><br /><br />



</form>

<%
End if
%>



</body>
</html>