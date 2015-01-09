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

<title>Compactación y Reparación de la Base de Datos</title>



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



<br />



<%

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
			
			
			End If
			
			Set objFSO = Nothing
			
		    %>
		                            
</table>
            
            
            
            
            <br /><br />
            
<table width="600" border="0" align="center">
            <tr>
            <td align="left" class="powered1">Ruta donde se encuentran los justificantes emitidos por los alumnos/as</td>
            </tr>
</table>
            
            
            
            <table width="600" height="25" border="1" align="center">
            
			<%
			
			'#####################################
			'PROCESO DE BORRADO DE FICHEROS PDF. (JUSTIFICANTES
			'DE SELECCIÓN DE PRÁTICAS CLÍNICAS). 
			'#####################################
			
			Dim obj2FSO
			
			Dim nombre_carpeta
			Dim carpeta
			Dim nombre_archivo
			Dim archivos
			
			Dim fileaux
			
			
			'obtengo el directorio físico de la carpeta donde está este script
			nombre_carpeta = Server.MapPath("/gpc/pdfs/ ")
			
			%>
			
            <tr>
            <td align="center" class="textonombrecarpeta1"><%=nombre_carpeta%> </td>
            </tr>
            
            <%
            
			'Conecto con el sistema de archivos
			set obj2FSO = server.createObject("Scripting.FileSystemObject")
			
			'creo el objeto carpeta
			Set carpeta = obj2FSO.GetFolder(nombre_carpeta)
			
			'traigo los archivos de la carpeta
			Set archivos = carpeta.Files
			
	%>
</table>
            
            <br /><br />
    
    
			
            
            
<table width="1064" border="1" align="center">
    
    
			<%
					
			'para cada archivo, muestro su nombre.
			for each nombre_archivo in archivos
			%>
				
                      <tr>
                      <td width="188" class="textoregistrosleidos1normal"><%Response.Write "fichero existente: "%></td><td width="849" class="textoregistrosleidos1normal"><%Response.Write(nombre_archivo & "<br/>")%></td>
                      </tr>
					  
					  <%
					  fileaux=nombre_archivo
						
					  obj2FSO.DeleteFile nombre_archivo, True
				      %>		
				      
					  <tr>
                      <td width="188" class="textoregistrosleidos1negrita"><%Response.Write "<b>fichero borrado: </b>"%></td><td width="849" class="textoregistrosleidos1negrita"><%Response.Write("<b>" & fileaux & "</b>")%></td>
      				</tr>
     		<%
			next
   			%>
    
</table>
    
 <br /><br /><br /> 
    			
			
			
<table width="1064" border="0" align="center">
            <tr>
            <td width="90"></td>
            <td align="center" class="textomensajes2">Borrado de ficheros pdf. Justificantes de selección de prácticas médicas terminado satisfactoriamente....</td>
            </tr>
</table>			
			
			<%	
			Set obj2FSO = Nothing
			%>


<br /><br />



<table border="0" align="center" width="349">
<tr>
<td></td>
<td width="291"  align="center"><a href="javascript:opener.location.href='admonplazasesp.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton1').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton1').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo1').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo1').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="48"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1" style="visibility:hidden;"/></a></td>
<td width="291"  align="center" class="powered1" id="idtextoregresarboton1" style="visibility:hidden; ">Regresar a la página de Gestión Administrativa sobre la Base de Datos...</td>
</tr>
</table>



<br/><br /><br />



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



</body>
</html>
