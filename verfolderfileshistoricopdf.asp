<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

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

<title>CARPETA CON LA RELACIÓN DE ARCHIVOS PDF GENERADO POR LOS ALUMNOS/AS</title>


<style type="text/css">

.textocabecera1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}

.textocabecera11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 35px;
	background-color: #999966;
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



</head>

<body>

<%
Dim ValorFolderDestino
ValorFolderDestino=Request.Form("SeleccionPromocion")
%>

<form name="formularioverficherospdf" method="post" action="verfolderfilespdf.asp" autocomplete="off">

<table width="1000" align="center" border="0">
<tr>
<td width="850" class="textocabecera1" align="center">RELACIÓN DE INFORMES DE ASIGNACIÓN EMITIDOS POR ALUMNOS/AS EN EL PROCESO DE ALTA DE PRÁCTICAS CLÍNICAS EN EL PERIODO <%=ValorFolderDestino%> DEL HISTORICO</td>
</tr>
</table>

<br />



<table width="858" border="0" align="center">
<tr>
    <td width="184" align="center"><img src="pngs/carpeta_destino_1.png" width="128" height="128" /></td>
	<td width="664" scope="col" class="textocabecera11"><div align="center">PERIODO <%=ValorFolderDestino%> DEL HISTORICO SELECCIONADO</div></td>

</tr>
</table>

<br />


<%
    Set MyFolder=Server.CreateObject("Scripting.FileSystemObject")
    Set MyFiles=MyFolder.GetFolder(Server.MapPath("/gpc/historico/" & ValorFolderDestino))
%>
   
	<%For each FoundFile in MyFiles.files%>
	
            <table width="767" border="0" align="center"> 
                    <tr>
                          <td width="73"><img src="pngs/pdf_icon_32_x_32.png"/></td>
                          <td width="684"><a href="/gpc/historico/<%=ValorFolderDestino & "/" & FoundFile.Name%>" target="_blank"><%=FoundFile.Name%></a></td>
                    </tr>
            </table>

	<%
    Next
    %>
	  
<br /><br /><br />


<table border="0" align="center" width="369">
<tr>
<td></td>
<td width="311"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton1').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton1').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo1').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo1').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="48"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1" style="visibility:hidden;"/></a></td>
<td width="311"  align="center" class="powered1" id="idtextoregresarboton1" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>


<br /> <br /><br /> 


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
