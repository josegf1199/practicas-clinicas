<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="es">

<%
Response.ContentType="text/html"
Response.Charset="UTF-8"
Session.CodePage=65001
%>


<title></title>

<style>
	
	.menu 
	{
		font:Verdana, Arial, Helvetica, sans-serif;
		font-family:Verdana, Arial, Helvetica, sans-serif;
		font-size: 14px;
		font-weight: bold;
		


				
		width: 25%;
		height: 50px;
		
		float: left;
		padding: 10px;	
		text-align: center;
		background: #fff;
		color: #000;
	}
	
	
	.menu:hover
	{
		*background: #000;
		color: #fff;	
		
		background-color: #8E8E46;
		border: 2px solid #505027;

	}
	
	
	#content 
	{
		clear: both;
		background: #e5e5e5;
		padding: 0;
		overflow-y: scroll;
		width: 100%;
		height: 600px;
		border: 0;
	}
	
	
	.textocabecera1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
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
	color: #999999 ;
	font-weight: normal;

}
	
	
</style>

<!--<script src="http://ajax.googleapis.com/ajax/libs/jquery/2.0.0/jquery.min.js"></script>-->

<script type="text/JavaScript" src="jquery/jquery.min.js"></script>

</head>

<body>

<table width="939" align="center" border="0">
<tr>
<td width="933" scope="col" class="textocabecera1"><div align="center">ÁREA DE VISUALIZACIÓN E IMPRESIÓN DE INFORME DE SELECCIÓN DE ESPECIALIDADES CLÍNICAS</div></td>
</tr>
</table>


<%
						Dim RefClaveSeleccionada
						RefClaveSeleccionada=Request.QueryString("RefNumDocAsignacionPracticas")
%>

<script>
						$(document).ready(function(e) 
								{
										$('#area_pdf').on('click', function()
														{
																			$('#content').attr('src', "prnpdfseleccion.asp?RefNum=" +  '<%Response.Write(RefClaveSeleccionada)%>');
														});
								});
</script>

<div id="main" >
	
    <div id="nav">

	  <div id="area_pdf" class="menu">
        
        <table width="445">
        <tr>
   		<td width="78" align="center"><img src="pngs/entorno_pdf_48_x_48_2.png"/></td>
        <td width="292" align="center">Pulse para ver el informe PDF e imprimir documento...</td>
        </tr>
        </table>
        
      </div>
      
      
     </div>
		
        <!--<div id="area_salir" class="menu"><a href="http://localhost/gpc/altapracticaclinica.asp">Salir...</a></div>-->
		<div id="area_salir" class="menu">

        <table width="460">
        <tr>
   		<td width="54"  align="center"><img src="pngs/back_to_home_48_x_48.png"/></td>
        <td width="382" align="center"><a href="/gpc/altapracticaclinica.asp">Regresar a la página principal de Gestión de Prácticas Clínicas...</a></td>
        </tr>
        </table>


    
    </div>

</div>

<iframe id="content" src="http://google.com">iFrames not supported </iframe>


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



</body>
</html>
