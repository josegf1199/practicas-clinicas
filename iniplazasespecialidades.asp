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


<!--#INCLUDE FILE="adovbs.inc"-->

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

</style>


<script type="text/javascript">

function gotoLink(url)
				{
						location.href = url;
				}
				
</script>


<title>DEFINICIÓN PLAZAS ESPECILIDADES</title>


<body>

<form name="forminiplazasespecialidades" method="post" action="iniplazasespecialidades.asp" autocomplete="off">


<%
	Dim NuevasPlazasGeneralDigestivaCiclo1
	Dim NuevasPlazasCardiacaCiclo1
	Dim NuevasPlazasToracicaCiclo1
	Dim NuevasPlazasMaxilofacialCiclo1
	Dim NuevasPlazasPlasticaCiclo1
	Dim NuevasPlazasVascularCiclo1
	Dim NuevasPlazasNeurocirugiaCiclo1
	Dim NuevasPlazasTraumatologiaCiclo1
	Dim NuevasPlazasUrologiaCiclo1
	Dim NuevasPlazasPediatriaCiclo1
	
	Dim NuevasPlazasHematologiaCiclo1
	Dim NuevasPlazasNeumologiaCiclo1
	Dim NuevasPlazasCardiologiaCiclo1
	Dim NuevasPlazasDigestivoCiclo1
	Dim NuevasPlazasNefrologiaCiclo1
	Dim NuevasPlazasMedicinaInternaCiclo1
	Dim NuevasPlazasEndocrinologiaCiclo1
	Dim NuevasPlazasReumatologiaCiclo1
	Dim NuevasPlazasOncologiaCiclo1
	Dim NuevasPlazasNeurologiaCiclo1
	Dim NuevasPlazasInfecciososCiclo1
	Dim NuevasPlazasMedicinaIntensivaCiclo1
	
	Dim NuevasPlazasEspecialidadACiclo1
	Dim NuevasPlazasEspecialidadBCiclo1
	Dim NuevasPlazasEspecialidadCCiclo1

	Dim NuevasPlazasEspecialidadDCiclo1
	Dim NuevasPlazasEspecialidadECiclo1
	Dim NuevasPlazasEspecialidadFCiclo1
	
		
	
	Dim NuevasPlazasGeneralDigestivaCiclo2
	Dim NuevasPlazasCardiacaCiclo2
	Dim NuevasPlazasToracicaCiclo2
	Dim NuevasPlazasMaxilofacialCiclo2
	Dim NuevasPlazasPlasticaCiclo2
	Dim NuevasPlazasVascularCiclo2
	Dim NuevasPlazasNeurocirugiaCiclo2
	Dim NuevasPlazasTraumatologiaCiclo2
	Dim NuevasPlazasUrologiaCiclo2
	Dim NuevasPlazasPediatriaCiclo2
	
	Dim NuevasPlazasHematologiaCiclo2
	Dim NuevasPlazasNeumologiaCiclo2
	Dim NuevasPlazasCardiologiaCiclo2
	Dim NuevasPlazasDigestivoCiclo2
	Dim NuevasPlazasNefrologiaCiclo2
	Dim NuevasPlazasMedicinaInternaCiclo2
	Dim NuevasPlazasEndocrinologiaCiclo2
	Dim NuevasPlazasReumatologiaCiclo2
	Dim NuevasPlazasOncologiaCiclo2
	Dim NuevasPlazasNeurologiaCiclo2
	Dim NuevasPlazasInfecciososCiclo2
	Dim NuevasPlazasMedicinaIntensivaCiclo2
	
	Dim NuevasPlazasEspecialidadACiclo2
	Dim NuevasPlazasEspecialidadBCiclo2
	Dim NuevasPlazasEspecialidadCCiclo2

	Dim NuevasPlazasEspecialidadDCiclo2
	Dim NuevasPlazasEspecialidadECiclo2
	Dim NuevasPlazasEspecialidadFCiclo2
%>	



<%

Call ModificarNumPlazasTablaServiciosC1("GDIG1","NuevasPlazasGeneralDigestivaCiclo1","NPlazasGeneralDigestivaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("GDIG2","NuevasPlazasGeneralDigestivaCiclo2","NPlazasGeneralDigestivaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("CARD1","NuevasPlazasCardiacaCiclo1","NPlazasCardiacaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("CARD2","NuevasPlazasCardiacaCiclo2","NPlazasCardiacaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("TORA1","NuevasPlazasToracicaCiclo1","NPlazasToracicaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("TORA2","NuevasPlazasToracicaCiclo2","NPlazasToracicaCiclo2")
	
Call ModificarNumPlazasTablaServiciosC1("MAXI1","NuevasPlazasMaxilofacialCiclo1","NPlazasMaxilofacialCiclo1")
Call ModificarNumPlazasTablaServiciosC2("MAXI2","NuevasPlazasMaxilofacialCiclo2","NPlazasMaxilofacialCiclo2")

Call ModificarNumPlazasTablaServiciosC1("PLAS1","NuevasPlazasPlasticaCiclo1","NPlazasPlasticaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("PLAS2","NuevasPlazasPlasticaCiclo2","NPlazasPlasticaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("VASC1","NuevasPlazasVascularCiclo1","NPlazasVascularCiclo1")
Call ModificarNumPlazasTablaServiciosC2("VASC2","NuevasPlazasVascularCiclo2","NPlazasVascularCiclo2")

Call ModificarNumPlazasTablaServiciosC1("NEUR1","NuevasPlazasNeurocirugiaCiclo1","NPlazasNeurocirugiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("NEUR2","NuevasPlazasNeurocirugiaCiclo2","NPlazasNeurocirugiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("TRAU1","NuevasPlazasTraumatologiaCiclo1","NPlazasTraumatologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("TRAU2","NuevasPlazasTraumatologiaCiclo2","NPlazasTraumatologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("UROL1","NuevasPlazasUrologiaCiclo1","NPlazasUrologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("UROL2","NuevasPlazasUrologiaCiclo2","NPlazasUrologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("PEDI1","NuevasPlazasPediatricaCiclo1","NPlazasPediatricaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("PEDI2","NuevasPlazasPediatricaCiclo2","NPlazasPediatricaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPA1","NuevasPlazasEspecialidadACiclo1","NPlazasEspecialidadACiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPA2","NuevasPlazasEspecialidadACiclo2","NPlazasEspecialidadACiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPB1","NuevasPlazasEspecialidadBCiclo1","NPlazasEspecialidadBCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPB2","NuevasPlazasEspecialidadBCiclo2","NPlazasEspecialidadBCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPC1","NuevasPlazasEspecialidadCCiclo1","NPlazasEspecialidadCCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPC2","NuevasPlazasEspecialidadCCiclo2","NPlazasEspecialidadCCiclo2")

Call ModificarNumPlazasTablaServiciosC1("HEMA1","NuevasPlazasHematologiaCiclo1","NPlazasHematologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("HEMA2","NuevasPlazasHematologiaCiclo2","NPlazasHematologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("NEUM1","NuevasPlazasNeumologiaCiclo1","NPlazasNeumologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("NEUM2","NuevasPlazasNeumologiaCiclo2","NPlazasNeumologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("CGIA1","NuevasPlazasCardiologiaCiclo1","NPlazasCardiologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("CGIA2","NuevasPlazasCardiologiaCiclo2","NPlazasCardiologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("DIGE1","NuevasPlazasDigestivoCiclo1","NPlazasDigestivoCiclo1")
Call ModificarNumPlazasTablaServiciosC2("DIGE2","NuevasPlazasDigestivoCiclo2","NPlazasDigestivoCiclo2")

Call ModificarNumPlazasTablaServiciosC1("NEFR1","NuevasPlazasNefrologiaCiclo1","NPlazasNefrologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("NEFR2","NuevasPlazasNefrologiaCiclo2","NPlazasNefrologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("MINT1","NuevasPlazasMedicinaInternaCiclo1","NPlazasMedicinaInternaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("MINT2","NuevasPlazasMedicinaInternaCiclo2","NPlazasMedicinaInternaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ENDO1","NuevasPlazasEndocrinologiaCiclo1","NPlazasEndocrinologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ENDO2","NuevasPlazasEndocrinologiaCiclo2","NPlazasEndocrinologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("REUM1","NuevasPlazasReumatologiaCiclo1","NPlazasReumatologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("REUM2","NuevasPlazasReumatologiaCiclo2","NPlazasReumatologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ONCO1","NuevasPlazasOncologiaCiclo1","NPlazasOncologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ONCO2","NuevasPlazasOncologiaCiclo2","NPlazasOncologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("NGIA1","NuevasPlazasNeurologiaCiclo1","NPlazasNeurologiaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("NGIA2","NuevasPlazasNeurologiaCiclo2","NPlazasNeurologiaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("INFE1","NuevasPlazasInfecciososCiclo1","NPlazasInfecciososCiclo1")
Call ModificarNumPlazasTablaServiciosC2("INFE2","NuevasPlazasInfecciososCiclo2","NPlazasInfecciososCiclo2")

Call ModificarNumPlazasTablaServiciosC1("MSIV1","NuevasPlazasMedicinaIntensivaCiclo1","NPlazasMedicinaIntensivaCiclo1")
Call ModificarNumPlazasTablaServiciosC2("MSIV2","NuevasPlazasMedicinaIntensivaCiclo2","NPlazasMedicinaIntensivaCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPD1","NuevasPlazasEspecialidadDCiclo1","NPlazasEspecialidadDCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPD2","NuevasPlazasEspecialidadDCiclo2","NPlazasEspecialidadDCiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPE1","NuevasPlazasEspecialidadECiclo1","NPlazasEspecialidadECiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPE2","NuevasPlazasEspecialidadECiclo2","NPlazasEspecialidadECiclo2")

Call ModificarNumPlazasTablaServiciosC1("ESPF1","NuevasPlazasEspecialidadFCiclo1","NPlazasEspecialidadFCiclo1")
Call ModificarNumPlazasTablaServiciosC2("ESPF2","NuevasPlazasEspecialidadFCiclo2","NPlazasEspecialidadFCiclo2")

%>







<%
Sub  ModificarNumPlazasTablaServiciosC1(ParametroEspecialidad1, VariableAcogeNuevasPlazas1, CampoTextoFromadmonNplazas1)


					
					Dim ObjConn1
					Set ObjConn1 = Server.CreateObject("ADODB.Connection")
					ObjConn1.Open "dsndbgpc"
					
					Dim SQL_SELECCION_GDIG1_TABLA_C1
					SQL_SELECCION_GDIG1_TABLA_C1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='" & ParametroEspecialidad1 & "'" 
					
					Dim ObjRs1
					Set ObjRs1=Server.CreateObject("ADODB.Recordset")
					
					ObjRs1.CursorType = 3
					ObjRs1.LockType = 3
					
					
					ObjRS1.Open SQL_SELECCION_GDIG1_TABLA_C1, ObjConn1
					
					VariableAcogeNuevasPlazas1=Request.Form(CampoTextoFromadmonNplazas1)
					
					
					
     				'Response.Write("<b><u><h5>" & ParametroEspecialidad1 & "</u></b></h5>")
				
					
					'If VariableAcogeNuevasPlazas1="" Then
					
					'			Response.Write("Nuevo valor: " & "Sin Valor (Null)" & " |--->| ")
					'Else
					'			Response.Write("Nuevo valor: " & VariableAcogeNuevasPlazas1 & " |--->| ")
					'End if
					
					'Response.Write("Valor Anterior: " & ObjRs1.Fields("PLAZAS_DISPONIBLES_C1") & " |--->| ")
					
					
					Dim MismoValor1
				    MismoValor1=ObjRs1.Fields("PLAZAS_DISPONIBLES_C1")
					
					
					If VariableAcogeNuevasPlazas1<>"" AND VariableAcogeNuevasPlazas1<>ObjRs1.Fields("PLAZAS_DISPONIBLES_C1") Then
					
								ObjRs1.Fields("PLAZAS_DISPONIBLES_C1")=VariableAcogeNuevasPlazas1
					'			Response.Write("Si modifica: " & VariableAcogeNuevasPlazas1)
					
					
					End If
					
					If VariableAcogeNuevasPlazas1="" Then			
					
								ObjRs1.Fields("PLAZAS_DISPONIBLES_C1")=MismoValor1
					'			Response.Write("No modifica: " & MismoValor1)
					
					End If
					
					
		
					
					'Response.Write("<br/>")
					
					
					
					
					ObjRs1.Update
					
					
					
					ObjRs1.Close
					Set ObjRs1 = Nothing
						
					ObjConn1.Close
					Set ObjConn1 = Nothing


End Sub
%>





<%
Sub  ModificarNumPlazasTablaServiciosC2(ParametroEspecialidad2, VariableAcogeNuevasPlazas2, CampoTextoFromadmonNplazas2)

					
					
					
					Dim ObjConn2
					Set ObjConn2 = Server.CreateObject("ADODB.Connection")
					ObjConn2.Open "dsndbgpc"
					
					Dim SQL_SELECCION_GDIG1_TABLA_C2
					SQL_SELECCION_GDIG1_TABLA_C2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='" & ParametroEspecialidad2 & "'" 
					
					Dim ObjRs2
					Set ObjRs2=Server.CreateObject("ADODB.Recordset")
					
					ObjRs2.CursorType = 3
					ObjRs2.LockType = 3
					
					
					ObjRS2.Open SQL_SELECCION_GDIG1_TABLA_C2, ObjConn2
					
					VariableAcogeNuevasPlazas2=Request.Form(CampoTextoFromadmonNplazas2)
					
					
					'Response.Write("<b><u><h5>" & ParametroEspecialidad2 & "</u></b></h5>")
					
					'If VariableAcogeNuevasPlazas2="" Then
					
					'			Response.Write("Nuevo valor: " & "Sin Valor (Null)" & " |--->| ")
					'Else
					'			Response.Write("Nuevo valor: " & VariableAcogeNuevasPlazas2 & " |--->| ")
					'End if
										
					'Response.Write("Valor Anterior: " & ObjRs2.Fields("PLAZAS_DISPONIBLES_C2") & " |--->| ")
					
					
					
					Dim MismoValor2
				    MismoValor2=ObjRs2.Fields("PLAZAS_DISPONIBLES_C2")
					
					
					If VariableAcogeNuevasPlazas2<>"" AND VariableAcogeNuevasPlazas2<>ObjRs2.Fields("PLAZAS_DISPONIBLES_C2") Then
					
					
								ObjRs2.Fields("PLAZAS_DISPONIBLES_C2")=VariableAcogeNuevasPlazas2
					'			Response.Write("Si modifica: " & VariableAcogeNuevasPlazas2)
					
					End If
					
					
					If VariableAcogeNuevasPlazas2="" Then			
					
								ObjRs2.Fields("PLAZAS_DISPONIBLES_C2")=MismoValor2
					'			Response.Write("No modifica: " & MismoValor2)
					
					End If
		
									
					'Response.Write("<br/><br/><br/>")
					
					
					
					
					
					ObjRs2.Update
					
					
					ObjRs2.Close
					Set ObjRs2 = Nothing
						
					ObjConn2.Close
					Set ObjConn2 = Nothing


End Sub
%>

<br /><br />

<table align="center" width="1026" border="0">
<tr>
    <td width="92"><img src="pngs/acepta_128_x_128.png" width="90" height="85" /></td>
    <td width="924" class="textotitulosubcabecerasinfo" align="center"><b><h1>¡¡¡PROCESO DE ACTUALIZACIÓN DE NÚMERO DE PLAZAS REALIZADO SATISFACTORIAMENTE!!!</h1></b></td>
</tr>
</table>



<br /><br />


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






</form>


</body>
</html>
