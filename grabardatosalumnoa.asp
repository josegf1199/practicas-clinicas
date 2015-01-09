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
</style>


<script type="text/javascript">

function gotoLink(url)
				{
						location.href = url;
				}
				
</script>


<title>GRABACIÓN DE DATOS EN BD</title>

</head>

<body>

<%

	Dim NumDocAsignacionPracticas
	Dim CadenaFechaExtendida
	Dim CadenaHora

	Dim CadenaDNIAlumnoa
	Dim CadenaNombreAlumnoa
	Dim CadenaApellidosAlumnoa
	Dim CadenaEmailAlumnoa
	
	Dim EstadoGeneralDigestivaCiclo1
	Dim EstadoCardiacaCiclo1
	Dim EstadoToracicaCiclo1
	Dim EstadoMaxilofacialCiclo1
	Dim EstadoPlasticaCiclo1
	Dim EstadoVascularCiclo1
	Dim EstadoNeurocirugiaCiclo1
	Dim EstadoTraumatologiaCiclo1
	Dim EstadoUrologiaCiclo1
	Dim EstadoPediatriaCiclo1
	
	Dim EstadoHematologiaCiclo1
	Dim EstadoNeumologiaCiclo1
	Dim EstadoCardiologiaCiclo1
	Dim EstadoDigestivoCiclo1
	Dim EstadoNefrologiaCiclo1
	Dim EstadoMedicinaInternaCiclo1
	Dim EstadoEndocrinologiaCiclo1
	Dim EstadoReumatologiaCiclo1
	Dim EstadoOncologiaCiclo1
	Dim EstadoNeurologiaCiclo1
	Dim EstadoInfecciososCiclo1
	Dim EstadoMedicinaIntensivaCiclo1
	
	Dim EstadoEspecialidadACiclo1
	Dim EstadoEspecialidadBCiclo1
	Dim EstadoEspecialidadCCiclo1

	Dim EstadoEspecialidadDCiclo1
	Dim EstadoEspecialidadECiclo1
	Dim EstadoEspecialidadFCiclo1
	
		
	
	Dim EstadoGeneralDigestivaCiclo2
	Dim EstadoCardiacaCiclo2
	Dim EstadoToracicaCiclo2
	Dim EstadoMaxilofacialCiclo2
	Dim EstadoPlasticaCiclo2
	Dim EstadoVascularCiclo2
	Dim EstadoNeurocirugiaCiclo2
	Dim EstadoTraumatologiaCiclo2
	Dim EstadoUrologiaCiclo2
	Dim EstadoPediatriaCiclo2
	
	Dim EstadoHematologiaCiclo2
	Dim EstadoNeumologiaCiclo2
	Dim EstadoCardiologiaCiclo2
	Dim EstadoDigestivoCiclo2
	Dim EstadoNefrologiaCiclo2
	Dim EstadoMedicinaInternaCiclo2
	Dim EstadoEndocrinologiaCiclo2
	Dim EstadoReumatologiaCiclo2
	Dim EstadoOncologiaCiclo2
	Dim EstadoNeurologiaCiclo2
	Dim EstadoInfecciososCiclo2
	Dim EstadoMedicinaIntensivaCiclo2
	
	Dim EstadoEspecialidadACiclo2
	Dim EstadoEspecialidadBCiclo2
	Dim EstadoEspecialidadCCiclo2

	Dim EstadoEspecialidadDCiclo2
	Dim EstadoEspecialidadECiclo2
	Dim EstadoEspecialidadFCiclo2
	
	
		
	NumDocAsignacionPracticas=Request.Form("ClaveAlumnoPasaraDB")
	CadenaFechaExtendida=Request.Form("FechaAsignacionPracticas")
	CadenaHora=Request.Form("HoraAsignacionPracticas")
	
	CadenaDNIAlumnoa=UCase(Request.Form("DNIAlumnoa"))
	CadenaNombreAlumnoa=UCase(Request.Form("NombreAlumnoa"))
	CadenaApellidosAlumnoa=UCase(Request.Form("ApellidosAlumnoa"))
	CadenaEmailAlumnoa=LCase(Request.Form("EmailAlumnoa"))
	
		
	Dim ListaEstadoEspecialidadesCiclo1(2,27)
	
	
	EstadoGeneralDigestivaCiclo1=Request.Form("GeneralDigestivaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,0)="GDIG1"
	ListaEstadoEspecialidadesCiclo1(1,0)="General-Digestiva"
	ListaEstadoEspecialidadesCiclo1(2,0)=Cbool(EstadoGeneralDigestivaCiclo1)
			
	

	
	EstadoCardiacaCiclo1=Request.Form("CardiacaCiclo1")
		
	ListaEstadoEspecialidadesCiclo1(0,1)="CARD1"
	ListaEstadoEspecialidadesCiclo1(1,1)="Cardiaca"
	ListaEstadoEspecialidadesCiclo1(2,1)=Cbool(EstadoCardiacaCiclo1)
		


		
	EstadoToracicaCiclo1=Request.Form("ToracicaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,2)="TORA1"
	ListaEstadoEspecialidadesCiclo1(1,2)="Torácica"
	ListaEstadoEspecialidadesCiclo1(2,2)=Cbool(EstadoToracicaCiclo1)
	
	
	
	EstadoMaxilofacialCiclo1=Request.Form("MaxilofacialCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,3)="MAXI1"
	ListaEstadoEspecialidadesCiclo1(1,3)="Maxilofacial"
	ListaEstadoEspecialidadesCiclo1(2,3)=Cbool(EstadoMaxilofacialCiclo1)
	
	
	
	EstadoPlasticaCiclo1=Request.Form("PlasticaCiclo1")

	ListaEstadoEspecialidadesCiclo1(0,4)="PLAS1"
	ListaEstadoEspecialidadesCiclo1(1,4)="Plástica"
	ListaEstadoEspecialidadesCiclo1(2,4)=Cbool(EstadoPlasticaCiclo1)



	EstadoVascularCiclo1=Request.Form("VascularCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,5)="VASC1"
	ListaEstadoEspecialidadesCiclo1(1,5)="Vascular"
	ListaEstadoEspecialidadesCiclo1(2,5)=Cbool(EstadoVascularCiclo1)
	
	
	
	EstadoNeurocirugiaCiclo1=Request.Form("NeurocirugiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,6)="NEUR1"
	ListaEstadoEspecialidadesCiclo1(1,6)="Neurocirugía"
	ListaEstadoEspecialidadesCiclo1(2,6)=Cbool(EstadoNeurocirugiaCiclo1)
	
	
		
	EstadoTraumatologiaCiclo1=Request.Form("TraumatologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,7)="TRAU1"
	ListaEstadoEspecialidadesCiclo1(1,7)="Traumatología"
	ListaEstadoEspecialidadesCiclo1(2,7)=Cbool(EstadoTraumatologiaCiclo1)
	
	
		
	EstadoUrologiaCiclo1=Request.Form("UrologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,8)="UROL1"
	ListaEstadoEspecialidadesCiclo1(1,8)="Urología"
	ListaEstadoEspecialidadesCiclo1(2,8)=Cbool(EstadoUrologiaCiclo1)
	
	
		
	EstadoPediatriaCiclo1=Request.Form("PediatricaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,9)="PEDI1"
	ListaEstadoEspecialidadesCiclo1(1,9)="Pediátrica"
	ListaEstadoEspecialidadesCiclo1(2,9)=Cbool(EstadoPediatriaCiclo1)
	
	
	
	
	EstadoHematologiaCiclo1=Request.Form("HematologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,10)="HEMA1"
	ListaEstadoEspecialidadesCiclo1(1,10)="Hematología"
	ListaEstadoEspecialidadesCiclo1(2,10)=Cbool(EstadoHematologiaCiclo1)
	
	
	
	EstadoNeumologiaCiclo1=Request.Form("NeumologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,11)="NEUM1"
	ListaEstadoEspecialidadesCiclo1(1,11)="Neumología"
	ListaEstadoEspecialidadesCiclo1(2,11)=Cbool(EstadoNeumologiaCiclo1)
	
	
	
	EstadoCardiologiaCiclo1=Request.Form("CardiologiaCiclo1")

	ListaEstadoEspecialidadesCiclo1(0,12)="CGIA1"
	ListaEstadoEspecialidadesCiclo1(1,12)="Cardiología"
	ListaEstadoEspecialidadesCiclo1(2,12)=Cbool(EstadoCardiologiaCiclo1)
	
	
		
	EstadoDigestivoCiclo1=Request.Form("DigestivoCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,13)="DIGE1"
	ListaEstadoEspecialidadesCiclo1(1,13)="Digestivo"
	ListaEstadoEspecialidadesCiclo1(2,13)=Cbool(EstadoDigestivoCiclo1)
	
	
		
	EstadoNefrologiaCiclo1=Request.Form("NefrologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,14)="NEFR1"
	ListaEstadoEspecialidadesCiclo1(1,14)="Nefrología"
	ListaEstadoEspecialidadesCiclo1(2,14)=Cbool(EstadoNefrologiaCiclo1)
	
	
	
	EstadoMedicinaInternaCiclo1=Request.Form("MedicinaInternaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,15)="MINT1"
	ListaEstadoEspecialidadesCiclo1(1,15)="Medicina Interna"
	ListaEstadoEspecialidadesCiclo1(2,15)=Cbool(EstadoMedicinaInternaCiclo1)
	
	
		
	EstadoEndocrinologiaCiclo1=Request.Form("EndocrinologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,16)="ENDO1"
	ListaEstadoEspecialidadesCiclo1(1,16)="Endocrinología"
	ListaEstadoEspecialidadesCiclo1(2,16)=Cbool(EstadoEndocrinologiaCiclo1)
	
	
		
	EstadoReumatologiaCiclo1=Request.Form("ReumatologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,17)="REUM1"
	ListaEstadoEspecialidadesCiclo1(1,17)="Reumatología"
	ListaEstadoEspecialidadesCiclo1(2,17)=Cbool(EstadoReumatologiaCiclo1)
	
	
		
	EstadoOncologiaCiclo1=Request.Form("OncologiaCiclo1")

	ListaEstadoEspecialidadesCiclo1(0,18)="ONCO1"
	ListaEstadoEspecialidadesCiclo1(1,18)="Oncología"
	ListaEstadoEspecialidadesCiclo1(2,18)=Cbool(EstadoOncologiaCiclo1)
	
	
		
	EstadoNeurologiaCiclo1=Request.Form("NeurologiaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,19)="NGIA1"
	ListaEstadoEspecialidadesCiclo1(1,19)="Neurología"
	ListaEstadoEspecialidadesCiclo1(2,19)=Cbool(EstadoNeurologiaCiclo1)
	
	
		
	EstadoInfecciososCiclo1=Request.Form("InfecciososCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,20)="INFE1"
	ListaEstadoEspecialidadesCiclo1(1,20)="Infecciosos"
	ListaEstadoEspecialidadesCiclo1(2,20)=Cbool(EstadoInfecciososCiclo1)
	
	
		
	EstadoMedicinaIntensivaCiclo1=Request.Form("MedicinaIntensivaCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,21)="MSIV1"
	ListaEstadoEspecialidadesCiclo1(1,21)="Medicina Intensiva"
	ListaEstadoEspecialidadesCiclo1(2,21)=Cbool(EstadoMedicinaIntensivaCiclo1)
	
	
	
	
	
	EstadoEspecialidadACiclo1=Request.Form("EspecialidadACiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,22)="ESPA1"
	ListaEstadoEspecialidadesCiclo1(1,22)="Especialidad - A"
	ListaEstadoEspecialidadesCiclo1(2,22)=Cbool(EstadoEspecialidadACiclo1)
	
	
	EstadoEspecialidadBCiclo1=Request.Form("EspecialidadBCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,23)="ESPB1"
	ListaEstadoEspecialidadesCiclo1(1,23)="Especialidad - B"
	ListaEstadoEspecialidadesCiclo1(2,23)=Cbool(EstadoEspecialidadBCiclo1)
	
	
	EstadoEspecialidadCCiclo1=Request.Form("EspecialidadCCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,24)="ESPC1"
	ListaEstadoEspecialidadesCiclo1(1,24)="Especialidad - C"
	ListaEstadoEspecialidadesCiclo1(2,24)=Cbool(EstadoEspecialidadCCiclo1)
	




	EstadoEspecialidadDCiclo1=Request.Form("EspecialidadDCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,25)="ESPD1"
	ListaEstadoEspecialidadesCiclo1(1,25)="Especialidad - D"
	ListaEstadoEspecialidadesCiclo1(2,25)=Cbool(EstadoEspecialidadDCiclo1)
	
	
	EstadoEspecialidadECiclo1=Request.Form("EspecialidadECiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,26)="ESPE1"
	ListaEstadoEspecialidadesCiclo1(1,26)="Especialidad - E"
	ListaEstadoEspecialidadesCiclo1(2,26)=Cbool(EstadoEspecialidadECiclo1)
	
	
	EstadoEspecialidadFCiclo1=Request.Form("EspecialidadFCiclo1")
	
	ListaEstadoEspecialidadesCiclo1(0,27)="ESPF1"
	ListaEstadoEspecialidadesCiclo1(1,27)="Especialidad - F"
	ListaEstadoEspecialidadesCiclo1(2,27)=Cbool(EstadoEspecialidadFCiclo1)
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	Dim  ListaEstadoEspecialidadesCiclo2(2,27)
	
	
	EstadoGeneralDigestivaCiclo2=Request.Form("GeneralDigestivaCiclo2")
		
	ListaEstadoEspecialidadesCiclo2(0,0)="GDIG2"
	ListaEstadoEspecialidadesCiclo2(1,0)="General-Digestiva"
	ListaEstadoEspecialidadesCiclo2(2,0)=Cbool(EstadoGeneralDigestivaCiclo2)
	
	
	
	EstadoCardiacaCiclo2=Request.Form("CardiacaCiclo2")
		
	ListaEstadoEspecialidadesCiclo2(0,1)="CARD2"
	ListaEstadoEspecialidadesCiclo2(1,1)="Cardiaca"
	ListaEstadoEspecialidadesCiclo2(2,1)=Cbool(EstadoCardiacaCiclo2)
	

		
	EstadoToracicaCiclo2=Request.Form("ToracicaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,2)="TORA2"
	ListaEstadoEspecialidadesCiclo2(1,2)="Torácica"
	ListaEstadoEspecialidadesCiclo2(2,2)=Cbool(EstadoToracicaCiclo2)
	
	
	
	EstadoMaxilofacialCiclo2=Request.Form("MaxilofacialCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,3)="MAXI2"
	ListaEstadoEspecialidadesCiclo2(1,3)="Maxilofacial"
	ListaEstadoEspecialidadesCiclo2(2,3)=Cbool(EstadoMaxilofacialCiclo2)
	
	
	
	EstadoPlasticaCiclo2=Request.Form("PlasticaCiclo2")

	ListaEstadoEspecialidadesCiclo2(0,4)="PLAS2"
	ListaEstadoEspecialidadesCiclo2(1,4)="Plástica"
	ListaEstadoEspecialidadesCiclo2(2,4)=Cbool(EstadoPlasticaCiclo2)



	EstadoVascularCiclo2=Request.Form("VascularCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,5)="VASC2"
	ListaEstadoEspecialidadesCiclo2(1,5)="Vascular"
	ListaEstadoEspecialidadesCiclo2(2,5)=Cbool(EstadoVascularCiclo2)
	
	
	
	EstadoNeurocirugiaCiclo2=Request.Form("NeurocirugiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,6)="NEUR2"
	ListaEstadoEspecialidadesCiclo2(1,6)="Neurocirugía"
	ListaEstadoEspecialidadesCiclo2(2,6)=Cbool(EstadoNeurocirugiaCiclo2)
	
	
		
	EstadoTraumatologiaCiclo2=Request.Form("TraumatologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,7)="TRAU2"
	ListaEstadoEspecialidadesCiclo2(1,7)="Traumatología"
	ListaEstadoEspecialidadesCiclo2(2,7)=Cbool(EstadoTraumatologiaCiclo2)
	
	
		
	EstadoUrologiaCiclo2=Request.Form("UrologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,8)="UROL2"
	ListaEstadoEspecialidadesCiclo2(1,8)="Urología"
	ListaEstadoEspecialidadesCiclo2(2,8)=Cbool(EstadoUrologiaCiclo2)
	
	
		
	EstadoPediatriaCiclo2=Request.Form("PediatricaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,9)="PEDI2"
	ListaEstadoEspecialidadesCiclo2(1,9)="Pediátrica"
	ListaEstadoEspecialidadesCiclo2(2,9)=Cbool(EstadoPediatriaCiclo2)
	
	
	
	
	
	
	EstadoHematologiaCiclo2=Request.Form("HematologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,10)="HEMA2"
	ListaEstadoEspecialidadesCiclo2(1,10)="Hematología"
	ListaEstadoEspecialidadesCiclo2(2,10)=Cbool(EstadoHematologiaCiclo2)
	
	
	
	EstadoNeumologiaCiclo2=Request.Form("NeumologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,11)="NEUM2"
	ListaEstadoEspecialidadesCiclo2(1,11)="Neumología"
	ListaEstadoEspecialidadesCiclo2(2,11)=Cbool(EstadoNeumologiaCiclo2)
	
	
	
	EstadoCardiologiaCiclo2=Request.Form("CardiologiaCiclo2")

	ListaEstadoEspecialidadesCiclo2(0,12)="CGIA2"
	ListaEstadoEspecialidadesCiclo2(1,12)="Cardiología"
	ListaEstadoEspecialidadesCiclo2(2,12)=Cbool(EstadoCardiologiaCiclo2)
	
	
		
	EstadoDigestivoCiclo2=Request.Form("DigestivoCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,13)="DIGE2"
	ListaEstadoEspecialidadesCiclo2(1,13)="Digestivo"
	ListaEstadoEspecialidadesCiclo2(2,13)=Cbool(EstadoDigestivoCiclo2)
	
	
		
	EstadoNefrologiaCiclo2=Request.Form("NefrologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,14)="NEFR2"
	ListaEstadoEspecialidadesCiclo2(1,14)="Nefrología"
	ListaEstadoEspecialidadesCiclo2(2,14)=Cbool(EstadoNefrologiaCiclo2)
	
	
	
	EstadoMedicinaInternaCiclo2=Request.Form("MedicinaInternaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,15)="MINT2"
	ListaEstadoEspecialidadesCiclo2(1,15)="Medicina Interna"
	ListaEstadoEspecialidadesCiclo2(2,15)=Cbool(EstadoMedicinaInternaCiclo2)
	
	
		
	EstadoEndocrinologiaCiclo2=Request.Form("EndocrinologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,16)="ENDO2"
	ListaEstadoEspecialidadesCiclo2(1,16)="Endocrinología"
	ListaEstadoEspecialidadesCiclo2(2,16)=Cbool(EstadoEndocrinologiaCiclo2)
	
	
		
	EstadoReumatologiaCiclo2=Request.Form("ReumatologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,17)="REUM2"
	ListaEstadoEspecialidadesCiclo2(1,17)="Reumatología"
	ListaEstadoEspecialidadesCiclo2(2,17)=Cbool(EstadoReumatologiaCiclo2)

	
	
		
	EstadoOncologiaCiclo2=Request.Form("OncologiaCiclo2")

	ListaEstadoEspecialidadesCiclo2(0,18)="ONCO2"
	ListaEstadoEspecialidadesCiclo2(1,18)="Oncología"
	ListaEstadoEspecialidadesCiclo2(2,18)=Cbool(EstadoOncologiaCiclo2)
	
	
		
	EstadoNeurologiaCiclo2=Request.Form("NeurologiaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,19)="NGIA2"
	ListaEstadoEspecialidadesCiclo2(1,19)="Neurología"
	ListaEstadoEspecialidadesCiclo2(2,19)=Cbool(EstadoNeurologiaCiclo2)
	
	
		
	EstadoInfecciososCiclo2=Request.Form("InfecciososCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,20)="INFE2"
	ListaEstadoEspecialidadesCiclo2(1,20)="Infecciosos"
	ListaEstadoEspecialidadesCiclo2(2,20)=Cbool(EstadoInfecciososCiclo2)
	
	
		
	EstadoMedicinaIntensivaCiclo2=Request.Form("MedicinaIntensivaCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,21)="MSIV2"
	ListaEstadoEspecialidadesCiclo2(1,21)="Medicina Intensiva"
	ListaEstadoEspecialidadesCiclo2(2,21)=Cbool(EstadoMedicinaIntensivaCiclo2)
	
	
	
	
	
	
	
	EstadoEspecialidadACiclo2=Request.Form("EspecialidadACiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,22)="ESPA2"
	ListaEstadoEspecialidadesCiclo2(1,22)="Especialidad - A"
	ListaEstadoEspecialidadesCiclo2(2,22)=Cbool(EstadoEspecialidadACiclo2)
	
	
	EstadoEspecialidadBCiclo2=Request.Form("EspecialidadBCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,23)="ESPB2"
	ListaEstadoEspecialidadesCiclo2(1,23)="Especialidad - B"
	ListaEstadoEspecialidadesCiclo2(2,23)=Cbool(EstadoEspecialidadBCiclo2)
	
	
	EstadoEspecialidadCCiclo2=Request.Form("EspecialidadCCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,24)="ESPC2"
	ListaEstadoEspecialidadesCiclo2(1,24)="Especialidad - C"
	ListaEstadoEspecialidadesCiclo2(2,24)=Cbool(EstadoEspecialidadCCiclo2)
	




	EstadoEspecialidadDCiclo2=Request.Form("EspecialidadDCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,25)="ESPD2"
	ListaEstadoEspecialidadesCiclo2(1,25)="Especialidad - D"
	ListaEstadoEspecialidadesCiclo2(2,25)=Cbool(EstadoEspecialidadDCiclo2)
	
	
	EstadoEspecialidadECiclo2=Request.Form("EspecialidadECiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,26)="ESPE2"
	ListaEstadoEspecialidadesCiclo2(1,26)="Especialidad - E"
	ListaEstadoEspecialidadesCiclo2(2,26)=Cbool(EstadoEspecialidadECiclo2)
	
	
	EstadoEspecialidadFCiclo2=Request.Form("EspecialidadFCiclo2")
	
	ListaEstadoEspecialidadesCiclo2(0,27)="ESPF2"
	ListaEstadoEspecialidadesCiclo2(1,27)="Especialidad - F"
	ListaEstadoEspecialidadesCiclo2(2,27)=Cbool(EstadoEspecialidadFCiclo2)
	

	
	
	Dim NombreEspecialidadPrimeraSeleccionCiclo1 
	Dim ReferenciaEspecialidadPrimeraSeleccionCiclo1
	
	Dim ReferenciaEspecialidadSegundaSeleccionCiclo1
	Dim NombreEspecialidadSegundaSeleccionCiclo1
	
	Dim EncontradaPrimeraSeleccion
	EncontradaPrimeraSeleccion=False
		
	Dim i
	
	For i=0 to 27
		
		if ListaEstadoEspecialidadesCiclo1(2,i)=True And EncontradaPrimeraSeleccion=False then
							
							EncontradaPrimeraSeleccion=True
							
							ReferenciaEspecialidadPrimeraSeleccionCiclo1=ListaEstadoEspecialidadesCiclo1(0,i)
							NombreEspecialidadPrimeraSeleccionCiclo1=ListaEstadoEspecialidadesCiclo1(1,i)
			
		Elseif ListaEstadoEspecialidadesCiclo1(2,i)=True And EncontradaPrimeraSeleccion=True then
		
		
							ReferenciaEspecialidadSegundaSeleccionCiclo1=ListaEstadoEspecialidadesCiclo1(0,i)
							NombreEspecialidadSegundaSeleccionCiclo1=ListaEstadoEspecialidadesCiclo1(1,i)
		
		End if
	Next
	
	
	
	
	
	
	Dim NombreEspecialidadUnicaSeleccionCiclo2 
	Dim ReferenciaEspecialidadUnicaSeleccionCiclo2
	
	Dim EncontradaUnicaSeleccion
	EncontradaUnicaSeleccion=False
		
	Dim e
	
	For e=0 to 27
		
		if ListaEstadoEspecialidadesCiclo2(2,e)=True And EncontradaUnicaSeleccion=False then
							
							EncontradaUnicaSeleccion=True
							
							ReferenciaEspecialidadUnicaSeleccionCiclo2=ListaEstadoEspecialidadesCiclo2(0,e)
							NombreEspecialidadUnicaSeleccionCiclo2=ListaEstadoEspecialidadesCiclo2(1,e)
							
		End if
	Next
	
	
Dim ObjConn
Dim ObjRs

Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_Alumnos", ObjConn,3,3

ObjRs.AddNew

ObjRs ("NUMERO_ALUMNOA") = CStr(NumDocAsignacionPracticas)

ObjRs ("DNI_ALUMNOA") = CStr(CadenaDNIAlumnoa)
ObjRs ("NOMBRE_ALUMNOA") = CStr(CadenaNombreAlumnoa)
ObjRs ("APELLIDOS_ALUMNOA") = CStr(CadenaApellidosAlumnoa)
ObjRs ("E_MAIL_ALUMNOA") = CStr(CadenaEmailAlumnoa)

ObjRs ("DENOMINACION_SERVICIO_SELECCIONADO_1") = "CIRUGÍA"			
ObjRs ("1_ESPECIALIDAD_SELECCIONADA_CICLO_1") = CStr(NombreEspecialidadPrimeraSeleccionCiclo1)
ObjRs ("REF_1_ESPECIALIDAD_SELECCIONADA_CICLO_1") = CStr(ReferenciaEspecialidadPrimeraSeleccionCiclo1)

ObjRs ("DENOMINACION_SERVICIO_SELECCIONADO_2") = "CIRUGÍA"
ObjRs ("2_ESPECIALIDAD_SELECCIONADA_CICLO_1") = CStr(NombreEspecialidadSegundaSeleccionCiclo1)
ObjRs ("REF_2_ESPECIALIDAD_SELECCIONADA_CICLO_1") = CStr(ReferenciaEspecialidadSegundaSeleccionCiclo1)

ObjRs ("DENOMINACION_SERVICIO_SELECCIONADO_3") = "MEDICINA INTERNA"
ObjRs ("3_ESPECIALIDAD_SELECCIONADA_CICLO_2") = CStr(NombreEspecialidadUnicaSeleccionCiclo2)
ObjRs ("REF_3_ESPECIALIDAD_SELECCIONADA_CICLO_2") = CStr(ReferenciaEspecialidadUnicaSeleccionCiclo2)

ObjRs ("FECHA_SELECCION_ESPECIALIDADES") = CStr(CadenaFechaExtendida)
ObjRs ("HORA_SELECCION_ESPECIALIDADES") = Time()												 'CStr(CadenaHora) 

ObjRs ("ALUMNOA_HA_ASIGNADO_SUS_PRACTICAS") = True

ObjRs.Update

ObjRs.Close
Set ObjRs = Nothing

ObjConn.Close
Set ObjConn = Nothing




' #################################################################################################
Dim ObjConn1
Set ObjConn1 = Server.CreateObject("ADODB.Connection")
ObjConn1.Open "dsndbgpc"

Dim SQL_BUSQUEDA_SELECCION_1_CICLO_1
SQL_BUSQUEDA_SELECCION_1_CICLO_1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='" & ReferenciaEspecialidadPrimeraSeleccionCiclo1 &"'"

Dim ObjRs1
Set ObjRs1=Server.CreateObject("ADODB.Recordset")

ObjRs1.CursorType = 3
ObjRs1.LockType = 3

ObjRS1.Open SQL_BUSQUEDA_SELECCION_1_CICLO_1, ObjConn1

Dim ValorActualPlazasDisponibles_seleccion_1_ciclo1
ValorActualPlazasDisponibles_seleccion_1_ciclo1 = ObjRs1.Fields("PLAZAS_DISPONIBLES_C1")

Dim NuevoValorPlazasDisponibles_seleccion_1_ciclo1
NuevoValorPlazasDisponibles_seleccion_1_ciclo1=Cint(ValorActualPlazasDisponibles_seleccion_1_ciclo1)-1

ObjRs1 ("PLAZAS_DISPONIBLES_C1") = CStr(NuevoValorPlazasDisponibles_seleccion_1_ciclo1)

ObjRs1.Update

ObjRs1.Close
Set ObjRs1 = Nothing
	
ObjConn1.Close
Set ObjConn1 = Nothing


' #################################################################################################
Dim ObjConn2
Set ObjConn2 = Server.CreateObject("ADODB.Connection") 
ObjConn2.Open "dsndbgpc"

Dim SQL_BUSQUEDA_SELECCION_2_CICLO_1
SQL_BUSQUEDA_SELECCION_2_CICLO_1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='" & ReferenciaEspecialidadSegundaSeleccionCiclo1 &"'"

Dim ObjRs2
Set ObjRs2 = Server.CreateObject("ADODB.Recordset")

ObjRs2.CursorType = 3
ObjRs2.LockType = 3

ObjRS2.Open SQL_BUSQUEDA_SELECCION_2_CICLO_1, ObjConn2

Dim ValorActualPlazasDisponibles_seleccion_2_ciclo1
ValorActualPlazasDisponibles_seleccion_2_ciclo1 = ObjRs2.Fields("PLAZAS_DISPONIBLES_C1")

Dim NuevoValorPlazasDisponibles_seleccion_2_ciclo1
NuevoValorPlazasDisponibles_seleccion_2_ciclo1=Cint(ValorActualPlazasDisponibles_seleccion_2_ciclo1)-1

ObjRs2 ("PLAZAS_DISPONIBLES_C1") = CStr(NuevoValorPlazasDisponibles_seleccion_2_ciclo1)

ObjRs2.Update

ObjRs2.Close
Set ObjRs2 = Nothing
	
ObjConn2.Close
Set ObjConn2 = Nothing



' #################################################################################################
Dim ObjConn3
Set ObjConn3 = Server.CreateObject("ADODB.Connection") 
ObjConn3.Open "dsndbgpc"

Dim SQL_BUSQUEDA_SELECCION_3_CICLO_2
SQL_BUSQUEDA_SELECCION_3_CICLO_2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='" & ReferenciaEspecialidadUnicaSeleccionCiclo2 &"'"

Dim ObjRs3
Set ObjRs3 = Server.CreateObject("ADODB.Recordset")

ObjRs3.CursorType = 3
ObjRs3.LockType = 3

ObjRS3.Open SQL_BUSQUEDA_SELECCION_3_CICLO_2, ObjConn3

Dim ValorActualPlazasDisponibles_seleccion_3_ciclo2
ValorActualPlazasDisponibles_seleccion_3_ciclo2 = ObjRs3.Fields("PLAZAS_DISPONIBLES_C2")

Dim NuevoValorPlazasDisponibles_seleccion_3_ciclo2
NuevoValorPlazasDisponibles_seleccion_3_ciclo2=Cint(ValorActualPlazasDisponibles_seleccion_3_ciclo2)-1

ObjRs3 ("PLAZAS_DISPONIBLES_C2") = CStr(NuevoValorPlazasDisponibles_seleccion_3_ciclo2)

ObjRs3.Update

ObjRs3.Close
Set ObjRs3 = Nothing
	
ObjConn3.Close
Set ObjConn3 = Nothing
%>

<form name="grabardatosalumnoa" method="post" action="grabardatosalumnoa.asp" autocomplete="off">

<%
Response.Redirect("enlacegrabatopdf.asp?RefNumDocAsignacionPracticas=" & NumDocAsignacionPracticas)
%>

</form>

</body>
</html>