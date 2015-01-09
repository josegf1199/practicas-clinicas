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


.textocabecera7espfree {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #99CC66;
	border: 1px solid #000000;
	font-weight: bold;
}

.avisoceroalumnoas1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #FF0000;
	border: 2px solid #000000;
	color: #FFFFFF;
	font-weight: bold;
}


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
	font-size: 14px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}


.textocabecera111{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}


.textocabecera1111 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #CCCCCC;
	*border: 2px solid #000000;
	color: #000000;
	font-weight: normal;
}

.textotitulocampo3 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #505027;
 	color: #FFFFFF;
  	font-weight:bold; 
 	border: 2px solid #505027;
 }


.textotitulocampo33 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	background-color: #505027;
 	color: #FFFFFF;
  	font-weight:bold; 
 	border: 2px solid #505027;
	
 }

.textotitulocampo333 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #505027;
 	color: #FFFFFF;
  	font-weight:bold; 
 	border: 2px solid #505027;
	
 }


.textodetallevacio6
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #CCCCCC;
	color: #FF0000;
	font-weight: bold;
}

.textodetallevacio66
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #999999;
	color: #003399 ;
	font-weight: bold;
}

.textodetallevacio66moresize
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	background-color: #999999;
	color: #003399 ;
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


.textocabecera7 {
	
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	
	background-color: #999966;
	*background-color: #8E8E46;
	
	border: 2px solid #505027;
	color: #000000;
	font-weight: bold;

}

</style>


<script type="text/javascript">


</script>


<title>RELACIÓN DE ALUMNOS/AS DISTRIBUIDOS POR CUATRIMESTRES, SERVICIOS Y ESPECIALIDADES</title>

</head>

<body leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0"  bgcolor="#FFFFFF">

<form name="paginalstalumnoascse" method="post" action="lstalumnoascse.asp">

<%

	Dim ObjConn
	Dim ObjRs
	
	Dim RegDescubierto1
	Dim RegDescubierto2
	
%>

<table width="900" align="center" border="0">
    <td align="center" class="textocabecera1">RELACIÓN DE ALUMNOS/AS DISTRIBUIDOS POR CUATRIMESTRES, SERVICIOS Y ESPECIALIDADES. <I>(por pantalla e impresora).</I></td>
</table>

<br /> 


<%
Dim SqlContar
Dim TotalRegistrosTabla

Dim CadenaNumRegistroActual

SqlContar = "SELECT * FROM tbl_Alumnos"

Set ObjConn = Server.CreateObject("ADODB.Connection")
Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjConn.Open "dsndbgpc"

ObjRs.CursorType = 1

ObjRs.Open SqlContar, ObjConn

TotalRegistrosTabla=ObjRs.RecordCount

ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>


<table border="0" align="center" width="266">
<tr>
<td></td>
<td width="232"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton1').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton1').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo1').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo1').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo1" style="visibility:hidden;"/></a></td>
<td width="232"  align="center" class="powered1" id="idtextoregresarboton1" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>



<%
If TotalRegistrosTabla=0 then 
%>
<table width="900" align="center" border="0">
<td align="center" class="avisoceroalumnoas1">FICHERO DE ALUMNOS/AS DE LA BASE DE DATOS ESTA VACIO...</td>
</table>

<%
Else
%>

<br /> 

<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo333">CICLO 1. (Primer Cuatrimestre)</td>
</tr>
</table>


<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo33">SERVICIOS QUIRÚRGICOS</td>
</tr>
</table>

<%
			Call ContructorListaEspecialidad1("GENERAL DIGESTIVA","GDIG1")
			Call ContructorListaEspecialidad1("CARDIACA","CARD1")
			Call ContructorListaEspecialidad1("TORÁCICA","TORA1")
			Call ContructorListaEspecialidad1("MAXILOFACIAL","MAXI1")
			Call ContructorListaEspecialidad1("PLÁSTICA","PLAS1")
			Call ContructorListaEspecialidad1("VASCULAR","VASC1")
			Call ContructorListaEspecialidad1("NEUROCIRUGÍA","NEUR1")
			Call ContructorListaEspecialidad1("TRAUMATOLOGÍA","TRAU1")
			Call ContructorListaEspecialidad1("UROLOGÍA","UROL1")
			Call ContructorListaEspecialidad1("PEDIÁTRICA","PEDI1")

			Call ContructorListaEspecialidad1("ESPECIALIDAD-A","ESPA1")
			Call ContructorListaEspecialidad1("ESPECIALIDAD-B","ESPB1")
			Call ContructorListaEspecialidad1("ESPECIALIDAD-C","ESPC1")
		
%>


<br /> <br /> <br /> 


<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo33">SERVICIOS MÉDICOS</td>
</tr>
</table>

<%
			Call ContructorListaEspecialidad1("HEMATOLOGÍA","HEMA1")
			Call ContructorListaEspecialidad1("NEUMOLOGÍA","NEUM1")
			Call ContructorListaEspecialidad1("CARDIOLOGÍA","CGIA1")
			Call ContructorListaEspecialidad1("DIGESTIVO","DIGE1")
			Call ContructorListaEspecialidad1("NEFROLOGÍA","NEFR1")
			Call ContructorListaEspecialidad1("MEDICINA INTERNA","MINT1")
			Call ContructorListaEspecialidad1("ENDOCRINOLOGÍA","ENDO1")
			Call ContructorListaEspecialidad1("REUMATOLOGÍA","REUM1")
			Call ContructorListaEspecialidad1("ONCOLOGÍA","ONCO1")
			Call ContructorListaEspecialidad1("NEUROLOGÍA","NGIA1")
			Call ContructorListaEspecialidad1("INFECCIOSOS","INFE1")
			Call ContructorListaEspecialidad1("MEDICINA INTENSIVA","MSIV1")
			
			Call ContructorListaEspecialidad1("ESPECIALIDAD-D","ESPD1")
			Call ContructorListaEspecialidad1("ESPECIALIDAD-E","ESPE1")
			Call ContructorListaEspecialidad1("ESPECIALIDAD-F","ESPF1")

%>




<br /> <br /> <br /> 



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





<br />  <br /> 

<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo333">CICLO 2. (Segundo Cuatrimestre)</td>
</tr>
</table>


<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo33">SERVICIOS QUIRÚRGICOS</td>
</tr>
</table>

<%
			Call ContructorListaEspecialidad2("GENERAL DIGESTIVA","GDIG2")
			Call ContructorListaEspecialidad2("CARDIACA","CARD2")
			Call ContructorListaEspecialidad2("TORÁCICA","TORA2")
			Call ContructorListaEspecialidad2("MAXILOFACIAL","MAXI2")
			Call ContructorListaEspecialidad2("PLÁSTICA","PLAS2")
			Call ContructorListaEspecialidad2("VASCULAR","VASC2")
			Call ContructorListaEspecialidad2("NEUROCIRUGÍA","NEUR2")
			Call ContructorListaEspecialidad2("TRAUMATOLOGÍA","TRAU2")
			Call ContructorListaEspecialidad2("UROLOGÍA","UROL2")
			Call ContructorListaEspecialidad2("PEDIÁTRICA","PEDI2")
			
			Call ContructorListaEspecialidad2("ESPECIALIDAD-A","ESPA2")
			Call ContructorListaEspecialidad2("ESPECIALIDAD-B","ESPB2")
			Call ContructorListaEspecialidad2("ESPECIALIDAD-C","ESPC2")

%>



<br /> <br /> <br /> 


<table width="900" align="center" border="0">
<tr>
<td align="center" class="textotitulocampo33">SERVICIOS MÉDICOS</td>
</tr>
</table>

<%
			Call ContructorListaEspecialidad2("HEMATOLOGÍA","HEMA2")
			Call ContructorListaEspecialidad2("NEUMOLOGÍA","NEUM2")
			Call ContructorListaEspecialidad2("CARDIOLOGÍA","CGIA2")
			Call ContructorListaEspecialidad2("DIGESTIVO","DIGE2")
			Call ContructorListaEspecialidad2("NEFROLOGÍA","NEFR2")
			Call ContructorListaEspecialidad2("MEDICINA INTERNA","MINT2")
			Call ContructorListaEspecialidad2("ENDOCRINOLOGÍA","ENDO2")
			Call ContructorListaEspecialidad2("REUMATOLOGÍA","REUM2")
			Call ContructorListaEspecialidad2("ONCOLOGÍA","ONCO2")
			Call ContructorListaEspecialidad2("NEUROLOGÍA","NGIA2")
			Call ContructorListaEspecialidad2("INFECCIOSOS","INFE2")
			Call ContructorListaEspecialidad2("MEDICINA INTENSIVA","MSIV2")
			
			Call ContructorListaEspecialidad2("ESPECIALIDAD-D","ESPD2")
			Call ContructorListaEspecialidad2("ESPECIALIDAD-E","ESPE2")
			Call ContructorListaEspecialidad2("ESPECIALIDAD-F","ESPF2")
			
%>




<%
Sub ContructorListaEspecialidad1(ValorServicio, ValorEspecialidad)
%>	
            
            <br /> 

			<table width="1254" align="center" border="0">
                        <tr>
                                    
                                    <%
                                    If ValorServicio="ESPECIALIDAD-A" OR ValorServicio="ESPECIALIDAD-B" OR ValorServicio="ESPECIALIDAD-C" OR ValorServicio="ESPECIALIDAD-D" OR ValorServicio="ESPECIALIDAD-E" OR ValorServicio="ESPECIALIDAD-F"  then
                                    
												%>
												<td width="533" align="center" class="textocabecera7espfree">ESPECIALIDAD '<%=ValorServicio%>'.</td>
    <%
									Else
									
												%>
												<td width="533" align="center" class="textocabecera7">ESPECIALIDAD '<%=ValorServicio%>'.</td>
			    <%
									End if			
                                    %>
                                    
                                    <td width="270" valign="bottom"><I>(Listado Ordenado por el Campo 'APELLIDOS').</I></td>
                        
                          <%
                                     RegDescubierto1=0
                                        
                                     Set ObjConn = Server.CreateObject("ADODB.Connection")
                                     Set ObjRs = Server.CreateObject("ADODB.Recordset")
                                    
                                     ObjConn.Open "dsndbgpc"
                                        
                                     Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
                                        
                                     ObjRs.MoveFirst
                                        
                                     
									 Do While Not ObjRs.EOF 
                                                    If ObjRS.Fields("REF_1_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=ValorEspecialidad OR  ObjRS.Fields("REF_2_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=ValorEspecialidad then
                                                            RegDescubierto1=RegDescubierto1+1
                                                    End If
                                                    ObjRs.MoveNext
                                     Loop
                                        
                                     
									 
									If RegDescubierto1=0 Then
                                                    %> 
                                                                              <td width="48" align="right"><img src="pngs/no_prn_1_48_x_48.png" width="48" height="48"/></td>
                                                                              <td width="48" align="right"><img src="pngs/sinalumnoas.png" width="48" height="48" /></td>
													<%
									
                                   Else
								   
								   
								   	
									
									
									If ValorEspecialidad="GDIG1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=GDIG1&RefEspecialidad=General Digestiva" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
                                   End if  
									 
									
									If ValorEspecialidad="CARD1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=CARD1&RefEspecialidad=Cardiaca" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
									End If

									
									If ValorEspecialidad="TORA1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=TORA1&RefEspecialidad=Torácica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
									End If
									
									
									If ValorEspecialidad="MAXI1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=MAXI1&RefEspecialidad=Maxilofacial" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If
									
									
									If ValorEspecialidad="PLAS1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=PLAS1&RefEspecialidad=Plástica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If
									
                             
									If ValorEspecialidad="VASC1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=VASC1&RefEspecialidad=Vascular" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="NEUR1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=NEUR1&RefEspecialidad=Neurocirugía" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
							 
									If ValorEspecialidad="TRAU1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=TRAU1&RefEspecialidad=Traumatología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="UROL1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=UROL1&RefEspecialidad=Urología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
							 
									If ValorEspecialidad="PEDI1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=PEDI1&RefEspecialidad=Pediátrica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="HEMA1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=HEMA1&RefEspecialidad=Hematología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 


									If ValorEspecialidad="NEUM1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=NEUM1&RefEspecialidad=Neumología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="CGIA1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=CGIA1&RefEspecialidad=Cardiología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="DIGE1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=DIGE1&RefEspecialidad=Digestivo" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="NEFR1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=NEFR1&RefEspecialidad=Nefrología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 


	
									If ValorEspecialidad="MINT1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=MINT1&RefEspecialidad=Medicina Interna" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ENDO1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ENDO1&RefEspecialidad=Endocrinología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="REUM1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=REUM1&RefEspecialidad=Reumatología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ONCO1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ONCO1&RefEspecialidad=Oncología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="NGIA1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=NGIA1&RefEspecialidad=Neurología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="INFE1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=INFE1&RefEspecialidad=Infecciosos" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="MSIV1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=MSIV1&RefEspecialidad=Medicina Intensiva" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPA1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPA1&RefEspecialidad=Especialidad - A" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPB1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPB1&RefEspecialidad=Especialidad - B" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPC1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPC1&RefEspecialidad=Especialidad - C" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPD1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPD1&RefEspecialidad=Especialidad - D" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPE1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPE1&RefEspecialidad=Especialidad - E" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPF1" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad1.asp?RefGrupo=ESPF1&RefEspecialidad=Especialidad - F" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
							 End if
                                    
									
									 ObjRs.Close
                                     Set ObjRs=Nothing
                                    
                                     ObjConn.Close
                                     Set ObjConn=Nothing
                                    %>
           
           
                        </tr>
            </table>
            
            
            
            <%
            Call  TitulosCampos()
            
            
            
			RegDescubierto1=0
				
            Set ObjConn = Server.CreateObject("ADODB.Connection")
            Set ObjRs = Server.CreateObject("ADODB.Recordset")
            
            ObjConn.Open "dsndbgpc"
                
            Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
                
            ObjRs.MoveFirst
                
            Do While Not ObjRs.EOF 
                 
                            If ObjRS.Fields("REF_1_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=ValorEspecialidad OR  ObjRS.Fields("REF_2_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=ValorEspecialidad then
            
                                    RegDescubierto1=RegDescubierto1+1
									Call DetalleCampos()
                    
                            End If
                                    
                            ObjRs.MoveNext
                    
            				Loop
							
							
			
			
				
			If RegDescubierto1=0 Then
							%> 
                                        <table width="1254" align="center" border="0">
                                                <tr>
                                                    <td width="134" align="center" class="textodetallevacio6">ESPECIALIDAD SIN ALUMNOS/AS AGRUPADOS DENTRO DE ESTA...</td>
                                                </tr>
                                        </table>
  							<%
			End If
			%>

			<table width="1254" align="center" border="0">
                        <tr>
                                    <td width="660" align="center" class="textodetallevacio66">RECUENTO DE ALUMNOS/AS EN ESTA ESPECIALIDAD HASTA AHORA ES DE: </td>
                          <td width="175" align="center" class="textodetallevacio66moresize"><%=RegDescubierto1%></td>
                                    <td width="405" align="left" class="textodetallevacio66">ALUMNOS/AS</td>
                        </tr>
            </table>


            <%
            ObjRs.Close
            Set ObjRs=Nothing
            
			ObjConn.Close
            Set ObjConn=Nothing
			



End Sub
%>







<%
Sub ContructorListaEspecialidad2(ValorServicio, ValorEspecialidad)
%>	

            <br /> 
            
            
            <table width="1254" align="center" border="0">
            <tr>

                                    <%
                                    if ValorServicio="ESPECIALIDAD-A" OR ValorServicio="ESPECIALIDAD-B" OR ValorServicio="ESPECIALIDAD-C" OR ValorServicio="ESPECIALIDAD-D" OR ValorServicio="ESPECIALIDAD-E" OR ValorServicio="ESPECIALIDAD-F" then
                                    
												%>
												<td width="533" align="center" class="textocabecera7espfree">ESPECIALIDAD '<%=ValorServicio%>'.</td>
												<%
									Else
									
												%>
												<td width="533" align="center" class="textocabecera7">ESPECIALIDAD '<%=ValorServicio%>'.</td>
												<%
									End if			
                                    %>





            <td width="270" valign="bottom"><I>(Listado Ordenado por el Campo 'APELLIDOS').</I></td>
            
			<%
				RegDescubierto2=0
				
                Set ObjConn = Server.CreateObject("ADODB.Connection")
                Set ObjRs = Server.CreateObject("ADODB.Recordset")
            
                ObjConn.Open "dsndbgpc"
                
                Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
                
                ObjRs.MoveFirst
                
                Do While Not ObjRs.EOF 
							If ObjRS.Fields("REF_3_ESPECIALIDAD_SELECCIONADA_CICLO_2").Value=ValorEspecialidad  then                                    
										RegDescubierto2=RegDescubierto2+1
                            End If
                            ObjRs.MoveNext
                Loop
				
				
									If RegDescubierto2=0 Then
                                                    %> 
                                                                              <td width="48" align="right"><img src="pngs/no_prn_1_48_x_48.png" width="48" height="48"/></td>
                                                                              <td width="48" align="right"><img src="pngs/sinalumnoas.png" width="48" height="48" /></td>
													<%
									
                                   Else
								   
								   
								   	
									
									
									If ValorEspecialidad="GDIG2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=GDIG2&RefEspecialidad=General Digestiva" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
                                   End if  
									 
									
									If ValorEspecialidad="CARD2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=CARD2&RefEspecialidad=Cardiaca" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
									End If

									
									If ValorEspecialidad="TORA2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=TORA2&RefEspecialidad=Torácica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="48" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
													<%
									End If
									
									
									If ValorEspecialidad="MAXI2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=MAXI2&RefEspecialidad=Maxilofacial" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If
									
									
									If ValorEspecialidad="PLAS2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=PLAS2&RefEspecialidad=Plástica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If
									
                             
									If ValorEspecialidad="VASC2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=VASC2&RefEspecialidad=Vascular" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="NEUR2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=NEUR2&RefEspecialidad=Neurocirugía" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
							 
									If ValorEspecialidad="TRAU2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=TRAU2&RefEspecialidad=Traumatología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="UROL2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=UROL2&RefEspecialidad=Urología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
							 
									If ValorEspecialidad="PEDI2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=PEDI2&RefEspecialidad=Pediátrica" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
							 
							 
									If ValorEspecialidad="HEMA2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=HEMA2&RefEspecialidad=Hematología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 


									If ValorEspecialidad="NEUM2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=NEUM2&RefEspecialidad=Neumología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="CGIA2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=CGIA2&RefEspecialidad=Cardiología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="DIGE2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=DIGE2&RefEspecialidad=Digestivo" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="NEFR2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=NEFR2&RefEspecialidad=Nefrología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 


	
									If ValorEspecialidad="MINT2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=MINT2&RefEspecialidad=Medicina Interna" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ENDO2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ENDO2&RefEspecialidad=Endocrinología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="REUM2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=REUM2&RefEspecialidad=Reumatología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ONCO2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ONCO2&RefEspecialidad=Oncología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="NGIA2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=NGIA2&RefEspecialidad=Neurología" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="INFE2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=INFE2&RefEspecialidad=Infecciosos" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="MSIV2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=MSIV2&RefEspecialidad=Medicina Intensiva" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPA2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPA2&RefEspecialidad=Especialidad - A" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPB2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPB2&RefEspecialidad=Especialidad - B" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPC2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPC2&RefEspecialidad=Especialidad - C" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPD2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPD2&RefEspecialidad=Especialidad - D" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPE2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPE2&RefEspecialidad=Especialidad - E" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 



									If ValorEspecialidad="ESPF2" then
                                                    %>                       
                                                                             <td width="52" align="right"><a href="prnpdfgrupoespecialidad2.asp?RefGrupo=ESPF2&RefEspecialidad=Especialidad - F" target="_blank"><img src="pngs/prn_1_48_x_48.png" width="48" height="48"/></a></td>
                                                                             <td width="50" align="right"><img src="pngs/conalumnoas.png" width="48" height="48"/></td>
                          							<%
									End If							 
				
							End If
	
				ObjRs.Close
                Set ObjRs=Nothing
            
                ObjConn.Close
                Set ObjConn=Nothing
			%>
           </tr>
           </table>


            
            <%
            Call  TitulosCampos()
            %>
            
            <%
				RegDescubierto2=0
				
                Set ObjConn = Server.CreateObject("ADODB.Connection")
                Set ObjRs = Server.CreateObject("ADODB.Recordset")
            
                ObjConn.Open "dsndbgpc"
                
                Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
                
                ObjRs.MoveFirst
                
                Do While Not ObjRs.EOF 
                 
                            If ObjRS.Fields("REF_3_ESPECIALIDAD_SELECCIONADA_CICLO_2").Value=ValorEspecialidad  then
            
                                    RegDescubierto2=RegDescubierto2+1
									Call DetalleCampos()
                    
                            End If
                                    
                            ObjRs.MoveNext
                    
                Loop
				
				If RegDescubierto2=0 Then
				
							%> 
                                        <table width="1254" align="center" border="0">
                                                <tr>
                                                    <td width="134" align="center" class="textodetallevacio6">ESPECIALIDAD SIN ALUMNOS/AS AGRUPADOS DENTRO DE ESTA...</td>
                                                </tr>
                                        </table>
  <%
				End If
                
				%>
                
<table width="1254" align="center" border="0">
                            <tr>
                                        <td width="660" align="center" class="textodetallevacio66">RECUENTO DE ALUMNOS/AS EN ESTA ESPECIALIDAD HASTA AHORA ES DE: </td>
                                        <td width="175" align="center" class="textodetallevacio66moresize"><%=RegDescubierto2%></td>
                              <td width="405" align="left" class="textodetallevacio66">ALUMNOS/AS</td>
                            </tr>
                </table>
     
            	<%

				


                ObjRs.Close
                Set ObjRs=Nothing
            
                ObjConn.Close
                Set ObjConn=Nothing

End Sub
 %>







<%
Sub TitulosCampos()
%>
            <table width="1254" align="center" border="0">
                        <tr>
                                    <td width="134" align="center" class="textocabecera11">DNI</td>
                                    <td width="394" align="center" class="textocabecera11">APELLIDOS</td>
                                    <td width="304" align="center" class="textocabecera11">NOMBRE</td>
                                    <td width="404" align="center" class="textocabecera11">E-MAIL</td>
                        </tr>
            </table>

<%
End Sub
%>





<%
Sub DetalleCampos()
%>

                        <table width="1254" align="center" border="0">
                                <tr>
                                  <td width="134" align="left" class="textocabecera1111"><%=ObjRs.Fields("DNI_ALUMNOA").Value%></td>
                                  <td width="394" align="left" class="textocabecera1111"><%=ObjRs.Fields("APELLIDOS_ALUMNOA").Value%></td>
                                  <td width="304" align="left" class="textocabecera1111"><%=ObjRs.Fields("NOMBRE_ALUMNOA").Value%></td>
                                  <td width="404" align="left" class="textocabecera1111"><%=ObjRs.Fields("E_MAIL_ALUMNOA").Value%></td>
                          </tr>
	                    </table>
<%
End Sub
%>



<br /><br /><br /><br />	




<table border="0" align="center" width="266">
<tr>
<td></td>
<td width="232"  align="center"><a href="javascript:opener.location.href='altapracticaclinica.asp'; self.close();" onmouseover="document.getElementById('idtextoregresarboton3').style='visibility: visible; font-weight:  bold;'" onmouseout="document.getElementById('idtextoregresarboton3').style='visibility: hidden;'"><img src="pngs/volvermainpagelamitad.png" onmouseover="document.getElementById('idlogoinfo3').style='visibility: visible;'" onmouseout=	"document.getElementById('idlogoinfo3').style='visibility: hidden;'"/></a></td>
</tr>
        		
<tr>
<td width="24"  align="center"><img src="pngs/logoinfo_48_x_48.png" id="idlogoinfo3" style="visibility:hidden;"/></a></td>
<td width="232"  align="center" class="powered1" id="idtextoregresarboton3" style="visibility:hidden; ">Regresar a la página principal de Alta de Prácticas Clínicas...</td>
</tr>
</table>


<br /><br /><br /><br /><br />


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

<%
End If
%>

</form>
	
</body>
</html>
