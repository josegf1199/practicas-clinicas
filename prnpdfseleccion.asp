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


<title>HOJA SELECCIÓNN PRÁCTICAS CLÍNICAS - FACULTAD DE MEDICINA DE GRANADA</title>


<style></style>


</head>

<body>


<%

Dim PDF

Dim ObjConn
Dim ObjRs

Dim QuerySql
Dim NumPage

%>

<!--#include file="fpdf.asp"-->
			

<%
						Dim RefClaveSeleccionada
						
						'RefClaveSeleccionada=Request.QueryString("RefNumDocAsignacionPracticas")
						
						RefClaveSeleccionada=Request.QueryString("RefNum")
						
						'Response.Write(RefClaveSeleccionada)

						
						NumPage=1
									
						
						Set PDF=CreateJsObject("FPDF")
						
						PDF.CreatePDF "P","mm","A4"
						PDF.SetPath("fpdf/")
						PDF.SetFont "Arial","",16
						
						PDF.Open()
						PDF.AddPage()
						
						
						Set ObjConn = Server.CreateObject("ADODB.Connection")
                        Set ObjRS=Server.CreateObject("ADODB.Recordset")
						
						ObjConn.Open "dsndbgpc"   
						
						ObjRs.CursorType = 3
						ObjRs.LockType = 3

						QuerySql="SELECT * FROM tbl_Alumnos WHERE NUMERO_ALUMNOA='" & RefClaveSeleccionada &"'"						
						ObjRs.Open QuerySql, ObjConn
						

						
						Call Cabecera()	
						Call BloqueNumDocFechaHora()
						Call BloqueDatosAlumnoa()
						Call BloqueDatosSeleccionCliclos1_2()
						Call BloqueNotificacion()
						
						Call BloqueFirmaAlumnoa()
						Call BloqueCopiaParaDepartamento()
						
							
						PDF.Addpage()		
						
						Numpage=Numpage+1														
													
													
						Call Cabecera()	
						Call BloqueNumDocFechaHora()
						Call BloqueDatosAlumnoa()
						Call BloqueDatosSeleccionCliclos1_2()
						Call BloqueNotificacion()
						
						Call BloqueFirmaGestor()
						Call BloqueCopiaParaAlunmo()	
						
						Call GeneraFicheroPdfSobreDisco()
						
						PDF.Output()
						
						
										
						
						PDF.Close()
						Set PDF=Nothing	
						
						ObjRs.Close
						Set ObjRs = Nothing
							
						ObjConn.Close
						Set ObjConn = Nothing   

						

Sub Cabecera()

						PDF.SetFont "Arial","", 8
						PDF.Text 191, 11, "Página: " + CStr(NumPage) + "/2"
						
						
						
						'Dim IconoVolverMainPage
						'IconoVolverMainPage= "jpg/volvermainpageinpdf_2.jpg"

						'PDF.Image IconoVolverMainPage, 198, 2,5,5
						'PDF.Link 198, 2,5,5,"http://localhost/gpc/altapracticaclinica.asp"

						
						
						Dim LogoFacultadMedicinaGranada
						LogoFacultadMedicinaGranada= "jpg/logougrfmed_300_x_185.jpg"

						PDF.Image LogoFacultadMedicinaGranada, 6, 9, 60, 40
						
					
						
						PDF.rect 6, 55, 198, 235
						
						PDF.SetFont "Arial","B",15
						PDF.Text 100, 40, "INFORME DE SELECCIÓN DE"
						PDF.Text 100, 45, "ESPECIALIDADES CLÍNICAS"
						
End Sub					
					
						

Sub BloqueNumDocFechaHora()
						
						PDF.SetFont "Arial","", 10
						PDF.Text 10 , 60 ,  "NÚMERO DE DOCUMENTO: "
						
						PDF.SetFont "Arial","B", 10
						PDF.Text 60 , 60 , (ObjRs.Fields("NUMERO_ALUMNOA")) 
						
						PDF.SetFont "Arial","I", 8
						PDF.Text 70 , 60 ,  "(Número de Orden del Alumno/a). "
						
						PDF.Line  6, 63 , 204 , 63
						
						
						PDF.SetFont "Arial","", 10
						PDF.Text 10 , 68 ,  "FECHA Y HORA  DE LA SELECCIÓN DE LAS PRÁCTICAS CLÍNICAS POR EL ALUMNO/A: "
						
						PDF.SetFont "Arial","B", 10
						PDF.Text 10 , 72 , (ObjRs.Fields("FECHA_SELECCION_ESPECIALIDADES")) 
						PDF.Text 75 , 72 , (ObjRs.Fields("HORA_SELECCION_ESPECIALIDADES")) 
						
						PDF.Line  6, 75 , 204 , 75

End Sub


Sub BloqueDatosAlumnoa()
						
						PDF.SetFont "Arial","", 10
						PDF.Text 10 , 80 ,  "DATOS DEL ALUMNO/A."
						
						PDF.Text 10 , 90 ,    "DNI:"
						PDF.Text 10 , 95 ,    "NOMBRE ALUMNO/A:"
						PDF.Text 10 , 100 ,  "APELLIDOS:"
						PDF.Text 10 , 105 ,  "DIRECCIÓN DE CORREO ELECTRÓNICO:"

						
						PDF.SetFont "times","B",10
						PDF.Text  100 , 90,   (ObjRs.Fields("DNI_ALUMNOA")) 
						PDF.Text  100 , 95,   (ObjRs.Fields("NOMBRE_ALUMNOA")) 
						PDF.Text  100 , 100, (ObjRs.Fields("APELLIDOS_ALUMNOA")) 
						PDF.Text  100 , 105, (ObjRs.Fields("E_MAIL_ALUMNOA")) 
						
						PDF.Line  6, 110 , 204 , 110
						
End Sub		




Sub BloqueDatosSeleccionCliclos1_2()
						
						PDF.SetFont "Arial","", 10
						PDF.Text 10 , 120 ,  "ESPECIALIDADES SELECCIONADAS DEL PRIMER CICLO."
						
						PDF.SetFont "Arial","I", 8
						PDF.Text 150 , 120 ,  "(Primer Cuatrimestre). "
						
						PDF.SetFont "times","B",10
						PDF.Text  100 , 130,   (ObjRs.Fields("DENOMINACION_SERVICIO_SELECCIONADO_1")) 
						PDF.Text  150 , 130,   (ObjRs.Fields("1_ESPECIALIDAD_SELECCIONADA_CICLO_1")) 
						
						PDF.Text  100 , 135, (ObjRs.Fields("DENOMINACION_SERVICIO_SELECCIONADO_2")) 
						PDF.Text  150 , 135, (ObjRs.Fields("2_ESPECIALIDAD_SELECCIONADA_CICLO_1"))
						

						PDF.SetFont "Arial","", 10
						PDF.Text 10 , 145 ,  "ESPECIALIDADES SELECCIONADAS DEL SEGUNDO CICLO."
						
						PDF.SetFont "Arial","I", 8
						PDF.Text 150 , 145 ,  "(Segundo Cuatrimestre). "

						PDF.SetFont "times","B",10
						PDF.Text  100 , 155, (ObjRs.Fields("DENOMINACION_SERVICIO_SELECCIONADO_3")) 
						PDF.Text  150 , 155, (ObjRs.Fields("3_ESPECIALIDAD_SELECCIONADA_CICLO_2"))
						
						PDF.Line  6, 160 , 204 , 160
					
End Sub										
						

						
Sub BloqueNotificacion()	
						
						PDF.SetFont "Arial","", 10
   						PDF.Text 10 , 165 ,  "NOTIFICACIÓN AL ALUMNO/A."
						
						PDF.SetFont "Arial","B", 10
						PDF.Text 10 , 175 ,  "Se le notifica al alumno/a mediante este informe la selección de las especialidades correspondientes al primer"
						PDF.Text 10 , 180 ,  "y segundo cuatrimestre que ha realizado a la fecha y hora arriba indicadas quedando informado al respecto " 
						PDF.Text 10 , 185 ,  "desde este departamento que ha gestionado la asignación de prácticas clínicas al alumno/a."

						PDF.Line  6, 190 , 204 , 190
End Sub						
						
						
						

Sub BloqueFirmaAlumnoa()

						'doc para el departamento
						PDF.SetFont "Arial","", 10
   						PDF.Text 10 , 195 ,  "FIRMA DEL ALUMNO/A."
						
						PDF.rect 20, 200, 170, 30
						PDF.Line 105, 200 , 105 , 230
						
						
						PDF.SetFont "Arial","I", 8
						PDF.Text 40 , 205 ,  "(Recibido, Enterado/a y Conforme)" 
						
						PDF.SetFont "Arial","", 8
						PDF.Text 40 , 210 ,  "Firma del Alumno/a..." 

End Sub						

						
						
						
Sub BloqueFirmaGestor()

						'doc para el alumnoa
						PDF.SetFont "Arial","", 10
   						PDF.Text 10 , 195 ,  "FIRMA Y SELLO DEL DEPARTAMENTO GESTOR."
						
						PDF.rect 20, 200, 170, 30
						PDF.Line 105, 200 , 105 , 230
						
						PDF.SetFont "Arial","I", 8
						PDF.Text 130 , 205 ,  "(Alumno/a Notificado/a)" 
						
						PDF.SetFont "Arial","", 8
						PDF.Text 130 , 210 ,  "Firma y Sello Departamento..." 

End Sub												
						
						

Sub BloqueCopiaParaAlunmo()

						PDF.SetFont "Arial","B",15
						PDF.Text 70, 260, "COPIA PARA EL ALUMNO/A..."
End Sub



Sub BloqueCopiaParaDepartamento()

						PDF.SetFont "Arial","B",15
						PDF.Text 50, 260, "COPIA PARA EL DEPARTAMENTO GESTOR..."
End Sub


Sub GeneraFicheroPdfSobreDisco()

						'####################################
						'CONSTRUCCIÓN DE LA TERMINACIÓN DE LA SEGÚN
						'LA FECHA DE LA PROMOCIÓN.
						'####################################
						Dim CadenaYearActual
						CadenaYearActual=right(date,2)
						
						Dim CadenaYearActualMasUno
						CadenaYearActualMasUno=CInt(CadenaYearActual)+1
						
						Dim DenominacionPromocion
						DenominacionPromocion="20"+CadenaYearActual+"_20"+CStr(CadenaYearActualMasUno)
			


						Dim txtfichero
						
						txtfichero=ObjRs.Fields("APELLIDOS_ALUMNOA") & "_" & ObjRs.Fields("NOMBRE_ALUMNOA") & "_" & DenominacionPromocion & ".pdf"  
						PDF.Output Server.MapPath("/gpc/pdfs/") & "/" & txtfichero,"F"

End Sub


%> 


</body>
</html>
