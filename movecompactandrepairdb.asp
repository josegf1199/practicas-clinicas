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


<style type="text/css">

.textotitulosubcabecerasinfo{
 
 	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
 	color:#333333;
  	font-weight: normal; 
 	border: 1px solid #999999; 
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

	
.textonombrecarpeta1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 30px;
	background-color: #FFFFFF;
	color: #666666 ;
	font-weight: bold;
}


.textomensajes1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 20px;
	background-color: #FFFFFF;
	color: #666666 ;
	font-weight: bold ;
}

.textomensajes2{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #FFFFFF;
	color: #000000 ;
	font-weight: bold ;
}


.textoregistrosleidos1normal{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
	color: #666666 ;
	font-weight: normal;
}

.textoregistrosleidos1negrita{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #FFFFFF;
	color: #666666 ;
	font-weight: bold ;
}



.textomensajeserror1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 20px;
	background-color: #FFFFFF;
	color: #FF0000;
	font-weight: bold;
}

</style>


</head>

<body>


<table align="center" width="1064" border="0">
<tr>
    <td width="90"><img src="pngs/acepta_128_x_128.png" width="90" height="85" /></td>
    <td width="962" class="textotitulosubcabecerasinfo" align="center"><b><h1>¡¡¡PROCESOS DE INICIALIZACIÓN REALIZADOS SATISFACTORIAMENTE!!!</h1></b></td>
</tr>
</table>


<%

			'#####################################
			'CONSTRUCCIÓN DEL NOMBRE DE CARPETA DEL HISTORICO
			'#####################################
			Dim CadenaYearActual
			CadenaYearActual=right(date,2)
			
			Dim CadenaYearActualMasUno
			CadenaYearActualMasUno=CInt(CadenaYearActual)+1
			
			Dim NombreCarpetaNueva
			NombreCarpetaNueva="20"+CadenaYearActual+"_20"+CStr(CadenaYearActualMasUno)
			



			'###################################
			'PROCESO DE CREACIÓN DE CARPETA HISTORICA PARA 
			'EL ALMACENAMIENTO DE LOS DOCUMENTOS PDF DE LA
			'EDICIÓN  DE ALUMNOS A ELIMINAR ANTES DEL INICIO 
			'DEL NUEVO PROCESO DE SELECCIÓN DE PRÁCTICAS 
			'CLÍNICAS DE LA NUEVA EDICIÓN DE ALUMNOS/AS. 
			'###################################
			
			Dim fso
			Dim folder
			
			Set fso = CreateObject("Scripting.FileSystemObject")
			
			Dim Ruta1
			
			Ruta1=Server.MapPath("/gpc")
			
			'If (Not fso.FolderExists(NombreCarpetaNueva)) then
    				Set folder = fso.CreateFolder(Ruta1 & "/historico/" & NombreCarpetaNueva)
    		'End if

			set folder=nothing
			set fso=nothing




			'##################################
			'PROCESO DE BORRADO COMPLETO DE REGSITROS DE 
			'LA TABLA DE ALUMNOS/AS. 
			'##################################

			Dim ObjConn
			Dim ObjRs
			
			Set ObjConn = Server.CreateObject("ADODB.Connection") 
			ObjConn.Open "dsndbgpc" 
			
			Set ObjRs = Server.CreateObject("ADODB.Recordset")
			ObjRs.Open "tbl_Alumnos", ObjConn,3,3

			Dim BorrarSQL
			
			BorrarSQL = "DELETE * FROM tbl_Alumnos" 
			ObjConn.Execute(BorrarSQL)
			%>
			
            
<table width="1064" height="28" border="0" align="center">
            
            <tr>
            <td width="167" align="right"><img src="pngs/acepta_32_x_32.png" width="32" height="32" /></td>
            <td width="881" align="center" class="textomensajes1">Se han eliminado todos los registros del fichero de alumnos/as satisfactoriamente...</td>
            </tr>
            
            
            
			
            <%
            
            
            'Response.write ("<br/><center><h3>Se han eliminado todos los registros del fichero de alumnos/as satisfactoriamente...</h3></center>") 
			
			ObjRs.Close
			Set ObjRs = Nothing
			
			ObjConn.Close
			Set ObjConn = Nothing

		
			
			
			'##########################
			'PROCESO DE COMPACTADO Y REPARADO 
			'DE LA BASE DE DATOS. 
			'##########################
			Dim oldFiledb
			Dim newFiledb
			
			Dim strConn
			Dim strConnBak
			
			Dim objJRO
			Dim objFSO
			

			oldFiledb = Server.MapPath("data/dbgpc.mdb")
			newFiledb = Server.MapPath("data/dbgpcbak.mdb")
			
			strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& oldFiledb
			strConnBak = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & newFiledb
			
			  
			Set objJRO = Server.CreateObject("JRO.JetEngine")
			objJRO.CompactDatabase strConn, strConnBak
		
			Set objJRO = Nothing
			
			
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			  
			If objFSO.FileExists(newFiledb) And objFSO.FileExists(oldFiledb) Then
				
						objFSO.DeleteFile(oldFiledb)
						objFSO.MoveFile newFiledb, oldFiledb
							
						'Response.Write "<br/><center><h3>Base de datos compactada y reparada satisfactoriamente para un nuevo uso...</h3></center>"
						%>
						<tr>
                        <td width="167" align="right"><img src="pngs/acepta_32_x_32.png" width="32" height="32" /></td>
                        <td align="center" class="textomensajes1" >Base de datos compactada y preparada satisfactoriamente...</td>
                        </tr>
						<%
			
			
			Else
						
			  
			  			%>
						<tr>
                        <td width="167" align="right"><img src="pngs/error_32_x_32.png" width="32" height="32" /></td>
                        <td align="center" class="textomensajeserror1">Compactación y reparación de la base de datos fallida...</td>
                        </tr>
						<%
						'Response.Write "<br/><center><h3>Compactación y reparación de la base de datos fallida...</h3></center>"
			  
			End If
			
			Set objFSO = Nothing
			
		%>
		                            
</table>


<br /><br />
		
		<%
			
			'########################################
			'PROCESO DE MOVIMIENTO DE FICHEROS PDF. (JUSTIFICANTES
			'DE SELECCIÓN DE PRÁTICAS CLÍNICAS). 
			'########################################
			
			Dim obj3FSO
			Dim Ruta2

			Set obj3FSO=Server.CreateObject("Scripting.FileSystemObject")
				
			Ruta2=Server.MapPath("\gpc")

			obj3FSO.MoveFile Ruta2 & "\pdfs\*.pdf", Ruta2 & "\historico\" & NombreCarpetaNueva

			
			%>
            
            
            
<table width="900
" border="0" align="center">
<tr>
            <td width="948" align="left" class="powered1">Ruta donde se han transferido a nivel histórico los justificantes emitidos por los alumnos/as en el ciclo de la promoción finalizada.</td>
  </tr>

            
</table>
			<table width="900" border="0" align="center">
			<tr>
            <td width="36" align="right"><img src="pngs/acepta_32_x_32.png" width="32" height="32" /></td>
            <td width="862" align="center" class="textomensajes1">Realizado el movimiento de ficheros pdf  de la carpeta '\gpc\pdfs\' a la carpeta del histórico denominada '\gpc\historico\<%Response.Write(NombreCarpetaNueva) &"'"%>...</td>
            </tr>
 </table>                      


<br /><br /><br />


</table>
			<table width="618" border="0" align="center">
			<tr>
            <td width="306" align="right"><img src="pngs/acepta_32_x_32.png" width="32" height="32" /></td>
            <td width="684" align="center" class="textomensajes1">Proceso completado satisfactoriamente....</td>
            </tr>
</table>   


			
<%		
			
			'Response.Write "<br/><center><h3>Movimiento de ficheros pdf  de la carpeta '\gpc\pdfs\' a la carpeta del historico denominada '\gpc\historico\" & NombreCarpetaNueva  & "</h3></center>"
			'Response.Write "<br/><br/><center><h1>Proceso terminado satisfactoriamente....</h1></center>"
			
			
			
			
			
			
			
			
			
			
			Set obj3FSO = Nothing
	
%>


<br /><br />




<table border="0" align="center" width="429">
<tr>
<td></td>
<td width="381"  align="center"><a href="javascript:opener.location.href='admonplazasesp.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton1').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton1').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo1').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo1').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="93"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1" style="visibility:hidden;"/></a></td>
<td width="381"  align="center" class="powered1" id="idtextoregresarboton1" style="visibility:hidden; ">Regresar a la página de Gestión Administrativa sobre la Base de Datos...</td>
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



</body>
</html>
