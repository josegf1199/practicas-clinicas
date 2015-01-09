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


<title>LISTADO DE ALUMNOS/AS CORRESPONDIENTES AL GRUPO DE LA ESPECIALIDAD INDICADA- FACULTAD DE MEDICINA DE GRANADA</title>


<style></style>


</head>

<body>


<%

Dim PDF

Dim ObjConn
Dim ObjRs

Dim QuerySql

Dim NumPage
Dim NumTotalPages

Dim RegDescubierto1

Dim ContarRegistros1

Dim Fila

Fila=66

%>

<!--#include file="fpdf.asp"-->
			

<%
						Dim RefClaveGrupo
						RefClaveGrupo=Request.QueryString("RefGrupo")
						
						Dim ValorCadenaEspecialidad
						ValorCadenaEspecialidad=Request.QueryString("RefEspecialidad")
						
						
						
						
						'##########################################
						'CONTAR REGSITROS QUE SATISFACEN UN CRITERIO Y CALCULAR 
						'EL NÚMERO DE PÁGINAS TOTAL.
						'##########################################
												
						ContarRegistros1=0
							
						Set ObjConn = Server.CreateObject("ADODB.Connection")
						Set ObjRs = Server.CreateObject("ADODB.Recordset")
						
						ObjConn.Open "dsndbgpc"
							
						Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
							
						ObjRs.MoveFirst
							
						Do While Not ObjRs.EOF 
							 
										If ObjRS.Fields("REF_1_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=RefClaveGrupo OR  ObjRS.Fields("REF_2_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=RefClaveGrupo then
												
												ContarRegistros1=ContarRegistros1+1
										
										End If
												
										ObjRs.MoveNext
						Loop
		
						ObjRs.Close
						Set ObjRs = Nothing
							
						ObjConn.Close
						Set ObjConn = Nothing 
						
						If ContarRegistros1<=28 then
								
								NumTotalPages=1
						
						Else
						
						
								
								NumTotalPages= Round(ContarRegistros1/28,2)
						
						
						
								Dim miNumero
								
								miNumero=NumTotalPages
								
								If (NumTotalPages - Int(NumTotalPages) >0) Then
								
										miNumero = Int(NumTotalPages) + 1
										NumTotalPages=miNumero
								
								Else
								
										NumTotalPages=miNumero
								
								End if
								
								
						
						
						End if
						'##########################################
						
						
						
						
						
						
						NumPage=0
						
												
						Set PDF=CreateJsObject("FPDF")
						
						PDF.CreatePDF "L","mm","A4"
						PDF.SetPath("fpdf/")
						PDF.SetFont "Arial","",16
						
						PDF.Open()

						PDF.AddPage()
						
	           			
						Call Cabecera()
						Call TituloCampos()	
						
            
						RegDescubierto1=0
							
						Set ObjConn = Server.CreateObject("ADODB.Connection")
						Set ObjRs = Server.CreateObject("ADODB.Recordset")
						
						ObjConn.Open "dsndbgpc"
							
						Set ObjRs=ObjConn.Execute("SELECT * FROM tbl_Alumnos ORDER BY APELLIDOS_ALUMNOA")
							
						ObjRs.MoveFirst
							
						Do While Not ObjRs.EOF 
							 
										If ObjRS.Fields("REF_1_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=RefClaveGrupo OR  ObjRS.Fields("REF_2_ESPECIALIDAD_SELECCIONADA_CICLO_1").Value=RefClaveGrupo then
						
												
												If RegDescubierto1= 28 then
																
																Fila=66
												
																RegDescubierto1=0 
												
																PDF.AddPage()
						      			
																Call Cabecera()
																Call TituloCampos()	
												End If
												
												Call BloqueDatosAlumnoa()
												
												RegDescubierto1=RegDescubierto1+1
								
										End If
												
										ObjRs.MoveNext
						Loop


					
						PDF.Output()
						
						
						PDF.Close()
						Set PDF=Nothing	
						
						ObjRs.Close
						Set ObjRs = Nothing
							
						ObjConn.Close
						Set ObjConn = Nothing   
						




Sub Cabecera()

						Numpage=Numpage+1

						PDF.SetFont "Arial","", 8
						PDF.Text 260, 11, "Página: " + CStr(NumPage) + "/ " + CStr(NumTotalPages)
					
						PDF.Line  260, 13 , 290 , 13
					
						PDF.Text 260, 18, "Número Registros: " 
						PDF.Text 260, 21, "Encontrados:  " + CStr(ContarRegistros1)
						
						
						Dim LogoFacultadMedicinaGranada
						LogoFacultadMedicinaGranada= "jpg/logougrfmed_300_x_185.jpg"

						PDF.Image LogoFacultadMedicinaGranada, 6, 9, 60, 40
						
					
						PDF.rect 6, 55, 285, 150
						
						
						PDF.SetFont "Arial","B",15
						PDF.Text 100, 30, "LISTADO DE ALUMNOS/AS CORRESPONDIENTES"
						PDF.Text 100, 35, "AL GRUPO DE LA ESPECIALIDAD DE: "
						
						PDF.SetFont "Arial","B",22
						PDF.rect 100, 38, 96, 9
						
						PDF.Text 102, 45, (ValorCadenaEspecialidad)
						
						PDF.SetFont "Arial","b", 12
						PDF.Text 200, 45 ,  "(PRIMER CUATRIMESTRE)"
End Sub					


Sub TituloCampos()

						PDF.SetFont "Arial","b", 10
						PDF.Text 6 , 54 ,  "DATOS DEL ALUMNO/A."
						
						PDF.SetFont "Arial","I", 7
						PDF.Text 238 , 54 ,  "(Listado Ordenado por el Campo 'APELLIDOS')."
						
						PDF.SetFont "Arial","b", 12
						
						PDF.Text     8, 60 ,  "DNI"
						PDF.Text   40, 60 ,  "APELLIDOS ALUMNO/A"
						PDF.Text 100, 60 ,  "NOMBRE ALUMNO/A"
						PDF.Text 200, 60 ,  "DIRECCIÓN DE CORREO ELECTRÓNICO"
						
						PDF.Line  6, 62 , 291 , 62
						
						PDF.Line 38, 55 , 38 , 205
						PDF.Line 98, 55 , 98 , 205
						PDF.Line 198, 55 , 198 , 205
End Sub	

				
						

Sub BloqueDatosAlumnoa()
						
						PDF.SetFont "times","",11
						PDF.Text  8 , Fila, (ObjRs.Fields("DNI_ALUMNOA")) 
						PDF.Text  40 , Fila, (ObjRs.Fields("APELLIDOS_ALUMNOA")) 
						PDF.Text  100 , Fila, (ObjRs.Fields("NOMBRE_ALUMNOA")) 
						PDF.Text  200 , Fila, (ObjRs.Fields("E_MAIL_ALUMNOA")) 
						
						Fila=Fila+5
End Sub		
%> 


</body>
</html>
