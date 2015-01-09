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

<title>MOVIMIENTO DE FICHEROS A HISTÓRICO, COMPACTACIÓN Y REPARACIÓN DE LA BASE DE DATOS</title>


<!--
<link rel="stylesheet" type="text/css" href="css/style.css" />
<script type='text/javascript' src='js/select.js'></script>
-->


<script type="text/javascript">


</script>



<style type="text/css">

.SelectStyle 

{
width: 220px;
position: relative;
}
 
 
 
select 

{
width: 100%;

background: #F3F3F3;
* background: #9ACDDE;

color: #585757;
padding: 5px;
font-size: 15px;
font:Verdana, Arial, Helvetica, sans-serif;
font-weight: bold;

line-height: 100%;
border: 1px solid #C1C1C1;
border-radius: 0;
height: 30px;
-webkit-appearance: none;
}


 
option 
{
padding: 10px;
} 






.SelectStyle:after 

{

width: 80px;
height: 30px;
display: block;

content: '';

position: absolute;
top: 0;right: 0;
pointer-events:  none;
border: 1px solid #C1C1C1;

background:#ebebeb;


background-image: url(/gpc/gif/arrowdown.gif);

*background-image: url(/gpc/gif/arrowdown.gif), -moz-linear-gradient(top,#dfdfdf 0%,#f6f6f6 100%);

*background-image: url(/gpc/gif/arrowdown.gif), -webkit-gradient(linear,left top,left bottom,color-stop(0%,#dfdfdf),color-stop(100%,#f6f6f6));

*background-image: url(/gpc/gif/arrowdown.gif), -webkit-linear-gradient(top,#dfdfdf 0%,#f6f6f6 100%);

*background-image: url(/gpc/gif/arrowdown.gif), -o-linear-gradient(top,#dfdfdf 0%,#f6f6f6 100%);

*background-image: url(/gpc/gif/arrowdown.gif), -ms-linear-gradient(top,#dfdfdf 0%,#f6f6f6 100%);

*background-image: url(/gpc/gifarrowdown.gif), linear-gradient(top,#dfdfdf 0%,#f6f6f6 100%);



background-repeat: no-repeat;

background-position: center ;

-webkit-box-sizing: border-box; /* Safari/Chrome, other WebKit */

-moz-box-sizing: border-box; /* Firefox, other Gecko */

box-sizing: border-box; /* Opera/IE 8+ */

} 










.powered11{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #FFFFFF;
	color: #666666 ;
	font-weight: bold;

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
	color:  #999999;
	font-weight: normal;
}


.textocabecera1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}


</style>






</head>

<body>


<%
	Dim DirName2
	DirName2="/gpc/historico/"
	
	Dim Fso2
	Dim FoldersExistentes2
	Dim FolderCount2
	
	Dim x
	
	Dim SubcarpetasDetectadas
	
	Dim SubcarpetaDetectada
	
	Dim NombreSubcarpeta
	
	
	Set Fso2 = CreateObject("Scripting.FileSystemObject")
	Set FoldersExistentes2 = Fso2.GetFolder(server.mappath(DirName2))
	
	%>


<form name="formularioselechistorico" method="post" action="verfolderfileshistoricopdf.asp" autocomplete="off">





<table width="1200" align="center" border="0">
<tr>
<td width="1200" class="textocabecera1" align="center">SISTEMA DE ACCESO AL HISTORICO DE PROMOCIONES ANTERIORES AL AÑO EN CURSO <%=Right(date,4)%> DE ALUMNOS/AS, DONDE CONSULTAR LOS INFORMES PDF DE LAS ASIGNACIONES EMITIDAS EN EL PROCESO DE ALTA DE PRÁCTICAS CLÍNICAS DE DICHOS PERIODOS.</td>
</tr>
</table>

<br />


<table width="800" border="0" align="center">
<tr>
    <td width="800" align="center"><img src="gif/az.gif" width="400" height="324" /></td>
</tr>
</table>

<br />


<table width="500"  align="center">  
<tr>
<td width="500"  align="left" class="powered11">Seleccione el Periodo Académico que desea Consultar:</td>
</tr>
</table>
   


<table width="500"  align="center">  

<td width="264" align="left"><div align="center" class="SelectStyle"> 

<!--
<select class="select" name="SeleccionPromocion" size="1" onchange="javascript:document.getElementById('idenviarperiodo').style='visibility:visible'; document.getElementById('idbuscar').style='visibility:visible';">
-->

<select class="select" name="SeleccionPromocion" size="1">


<%
For Each x in FoldersExistentes2.SubFolders
				
				NombreSubcarpeta = x.Name
				
				FolderCount2 = FolderCount2+1
				%>
				
                <option value="<%=NombreSubcarpeta%>"><%=NombreSubcarpeta%></option>
<%
Next
%>    
  
</select>
 </div>

 </td> 


 

 <!--
 <td width="121"  align="right"><img src="pngs/buscar_1_48_x_48.png" align="right" id="idbuscar" style="visibility: hidden"/></td>
 
 <td width="86" align="center" ><input type="Submit" name="enviarperiodo" id="idenviarperiodo" value="Buscar..."  style="visibility: hidden"/></td>
 -->  
 
 
  <td width="121"  align="right"><img src="pngs/buscar_1_48_x_48.png" align="right" id="idbuscar"  style=" visibility: visible"/></td>
 
 <td width="86" align="center" ><input type="Submit" name="enviarperiodo" id="idenviarperiodo" value="Buscar..."  style="visibility: visible"/></td>

 
   
   
   
</table>





























<br /><br /><br /><br /><br /><br />

<table border="0" align="center" width="355">
<tr>
<td></td>
<td width="297"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton3').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton3').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo3').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo3').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="48"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo3" style="visibility:hidden;"/></a></td>
<td width="297"  align="center" class="powered1" id="idtextoregresarboton3" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>


<br />


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

<br />	


</form>


</body>
</html>
