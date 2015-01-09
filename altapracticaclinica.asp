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


<style type="text/css">

.textocabecera1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 25px;
	background-color: #8E8E46;
	border: 2px solid #505027;
	color: #FFFFFF;
	font-weight: bold;
}

.textomenuinicio1 {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	*background-color:#CCCCCC;
	*border: 2px solid #505027;
	color: #000000;
	font-weight:normal;
}

.textotitulocabecera {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color: #CCCC99;
	border: 1px solid #000000;
	font-weight: normal;
}

.textotitulocampo1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #CCCC99;
	border: 1px solid #000000;
	font-weight: bold;
}

.textotitulocampo1espfree {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #99CC66;
	border: 1px solid #000000;
	font-weight: bold;
}

. textotituloactivarbotongrabar{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	*background-color: #CCCC99;
	*border: 1px solid #000000;
	font-weight: bold;
}

.textotitulocampo2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 13px;
	background-color: #FFFFFF; 
	border: none;
}

.textotitulocampo3 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 15px;
	background-color: #505027;
 	color: #FFFFFF;
  	font-weight:bold; 
 	border: 2px solid #505027;
	
 }

.textotitulosubcabeceras1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	background-color: #8E8E46;
 	color: #FFFFFF;
  	font-weight:bold; 
 	border: 2px solid #505027;
 }

.campotexto1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 19px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform: uppercase;
}

.campotextofechahora

{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:inherit;
}

.campotextoestado1{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #FF0000;
	font-weight:bold;
	text-transform: uppercase;
}


.campotextoestado2{

	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #009900;
	font-weight:bold;
	text-transform: uppercase;
}


.campoemail{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
	background-color:#CCCCCC;
	border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform: lowercase;
}

.stilocombobox {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	background-color:#666666;
	border: 1px solid #000000;
	color: #FFFFFF;
	font-weight: normal;
}


.cabeceraciclos
{
		font-family:Arial, Helvetica, sans-serif;
		font-size: 14px;
		font-weight:bold;
		color:#006699;
}


.textocirugia 
{
	color: #000000;
    font-family:Arial, Helvetica, sans-serif;
	font-size: 14px;
	background-color: #CCCCCC;
	border: 1px solid #000000;
	*font-weight: normal;
	font-weight:bold;
}


#bloquedatos
{
		font-family:Arial, Helvetica, sans-serif;
		font-size: 14px;
		font-weight:bold;
		
		color:#006699;
		border: 1px #solid #ff0000;
		display:block;
		
		*display: none;
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

				 
/* Base for label styling */
[type="checkbox"]:not(:checked),
[type="checkbox"]:checked {
  position: absolute;
  left: -9999px;
}
[type="checkbox"]:not(:checked) + label,
[type="checkbox"]:checked + label {
  position: relative;
  padding-left: 25px;
  cursor: pointer;
}

/* checkbox aspect */
[type="checkbox"]:not(:checked) + label:before,
[type="checkbox"]:checked + label:before {
  content: '';
  position: absolute;
  left:0; top: 2px;
  width: 17px; height: 17px;
  border: 1px solid #aaa;
  background: #f8f8f8;
  border-radius: 3px;
  box-shadow: inset 0 1px 3px rgba(0,0,0,.3)
}
/* checked mark aspect */
[type="checkbox"]:not(:checked) + label:after,
[type="checkbox"]:checked + label:after {
  content: '✔';
  position: absolute;
  top: 0; left: 4px;
  font-size: 14px;
  
  
  color:#FF0000;

   /*
  color: #09ad7e;
  color: #505027;
  color:#8E8E46;
 */
 
  transition: all .2s;
  font-weight:bold ;
}


/* checked mark aspect changes */
[type="checkbox"]:not(:checked) + label:after {
  opacity: 0;
  transform: scale(0);
}
[type="checkbox"]:checked + label:after {
  opacity: 1;
  transform: scale(1);
}
/* disabled checkbox */
[type="checkbox"]:disabled:not(:checked) + label:before,
[type="checkbox"]:disabled:checked + label:before {
  box-shadow: none;
  border-color: #bbb;
  background-color: #ddd;
}
[type="checkbox"]:disabled:checked + label:after {
  color: #999;
}
[type="checkbox"]:disabled + label {
  color: #aaa;
}
/* accessibility */
[type="checkbox"]:checked:focus + label:before,
[type="checkbox"]:not(:checked):focus + label:before {
  border: 1px dotted blue;
}

/* hover style just for information */
label:hover:before {
  border: 1px solid #4778d9!important;
}


#fecha
 {
	
/* width:80px;
	font-family: Tahoma, Verdana, Arial, sans-serif;
	font-size: .8em;
	color: #996600;
	background : #FFE3D7;
	text-align: center; 
*/
	
	font-family: Arial, Helvetica, sans-serif;
	font-size: 17px;
	background-color:#CCCCCC;
	* border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:inherit;
	text-align: center; 
	
	}

#hora 
{
/*	width:100px;
	font-family: Tahoma, Verdana, Arial, sans-serif;
	font-size: .9em;
	color: #996600;
	background : #FFE3D7;
	text-align: center;
*/	

	font-family: Arial, Helvetica, sans-serif;
	font-size: 17px;
	background-color:#CCCCCC;
	* border: 1px solid #000000;
	color: #000000;
	font-weight:bold;
	text-transform:lowercase;

}











 
#area_invalidar
{
 float:  left;
 border-style: solid;
 border-width: 0px;
 width: 1300px;
 height: 2050px;
}
 
#area_invalidar p
{
 padding: 2px 10px;
}
 

.clase_area_invalidar
{
 *position: absolute;
 *margin:auto;

position: absolute;
top:0;
left:0;
right:0;
bottom:0;

margin: auto;

*background: #83C24A;

height: 100px;
width: 200px;
box-shadow: 0 0 4px rgba(0,0,0,.3);
}


*html {background: #DBDBDB}



 

.clase_panel_carga
{
 position: absolute;
 width: 100%;
 height: 50%;
 
 /* Opacidad para Internet Explorer */
 filter:alpha(opacity=70);

 /* Opacidad en CSS3 estándar */
 opacity: 0.7;
 background-color: #333333;
 background-image:url('gif/loader.gif');
 background-repeat:no-repeat;
 background-position:center;
 z-index: 10;
 visibility: hidden;
}











</style>



<script type="text/javascript">

var contadorc1=0;
var contadorc1cirugia=0;
var contadorc1medint=0;
		
var contadorc2=0;
var contadorc2cirugia=0;
var contadorc2medint=0;
	
		
function validar(check,ciclo,tiposervicio,servicio) 
		
{
	
		var check;
		var ciclo;
		var tiposervicio;
		var servicio;
		
		var nopasar= false;
	
		
		if (check.checked==true && ciclo==1 && tiposervicio==1) 
					
			{
				contadorc1++;
				contadorc1cirugia++;
	
				
				if (contadorc1cirugia==2 &&  contadorc2cirugia==1)
					
					
					{
						alert("NO PUEDE SELECCIONAR TRES ESPECIALIDADES DENTROS DE 'SERVICIOS QUIRÚRGICOS'...");
		
						contadorc1--;
						contadorc1cirugia--;
						check.checked=false;
		
						nopasar=true;
					}
				

				if (contadorc1>2 && tiposervicio==1)
					
					{
						alert("SOLO SE PERMITE SELECCIONAR DOS ESPECIALIDADES DEL PRIMER CUATRIMESTRE ENTRE LOS SERVICIOS 'QUIRÚRGICOS' Y/O 'MEDICOS'...");
																					
						contadorc1--;
						contadorc1cirugia--;
		
						if (check.checked==true)
						
							{
								check.checked=false;
							}
						
						 	nopasar=true;
						
					}

				
				if (nopasar==false)
					
					{
						if (servicio=='gd1')
							{
								document.getElementById("idgeneraldigestivaciclo2").disabled=true;
							}
										
						if (servicio=='car1')
							{
								document.getElementById("idcardca2").disabled=true;
							}
										
						if (servicio=='tor1')
							{
								document.getElementById("idtoracica2").disabled=true;
							}
										
						if (servicio=='max1')
							{
								document.getElementById("idmaxilofacial2").disabled=true;
							}
										
						if (servicio=='pla1')
							{
								document.getElementById("idplastica2").disabled=true;
							}
										
						if (servicio=='vas1')
							{
								document.getElementById("idvascular2").disabled=true;
							}
										
						if (servicio=='neu1')
							{
								document.getElementById("idneurocirugia2").disabled=true;
							}
										
						if (servicio=='tra1')
							{
								document.getElementById("idtraumatologia2").disabled=true;
							}
				
						if (servicio=='uro1')
							{
								document.getElementById("idurologia2").disabled=true;
							}
										
						if (servicio=='ped1')
							{
								document.getElementById("idpediatrica2").disabled=true;
							}
					   	
						if (servicio=='espa1')
							{
								document.getElementById("idespecialidada2").disabled=true;
							}
					 	
						if (servicio=='espb1')
							{
								document.getElementById("idespecialidadb2").disabled=true;
							}
					 	
						if (servicio=='espc1')
							{
								document.getElementById("idespecialidadc2").disabled=true;
							}
					
					} 
			} 
								
												
		else 
												
												
		if (check.checked==false && ciclo==1 && tiposervicio==1)
												
			{
				contadorc1--;
				contadorc1cirugia--;

				if (servicio=='gd1')
					{
						document.getElementById("idgeneraldigestivaciclo2").disabled=false;
					}
				
				if (servicio=='car1')
					{
						document.getElementById("idcardca2").disabled=false;
					}
				
				if (servicio=='tor1')
					{
						document.getElementById("idtoracica2").disabled=false;
					}
				
				if (servicio=='max1')
					{
						document.getElementById("idmaxilofacial2").disabled=false;
					}
				
				if (servicio=='pla1')
					{
						document.getElementById("idplastica2").disabled=false;
					}
				
				if (servicio=='vas1')
					{
						document.getElementById("idvascular2").disabled=false;
					}
				
				if (servicio=='neu1')
					{
						document.getElementById("idneurocirugia2").disabled=false;
					}
				
				if (servicio=='tra1')
					{
						document.getElementById("idtraumatologia2").disabled=false;
					}
				
				if (servicio=='uro1')
					{
						document.getElementById("idurologia2").disabled=false;
					}
				
				if (servicio=='ped1')
					{
						document.getElementById("idpediatrica2").disabled=false;
					}
				
				if (servicio=='espa1')
					{
						document.getElementById("idespecialidada2").disabled=false;
					}
				
				if (servicio=='espb1')
					{
						document.getElementById("idespecialidadb2").disabled=false;
					}
				
				if (servicio=='espc1')
					{
						document.getElementById("idespecialidadc2").disabled=false;
					}
	
			}
												
												
		else
												
												
		if (check.checked==true && ciclo==1 && tiposervicio==2)
												
			{

				contadorc1++;
				contadorc1medint++;


					if (contadorc1medint==2 && contadorc2medint>0)
	
						{
							alert("NO PUEDE SELECCIONAR TRES ESPECIALIDADES DENTRO DE 'SERVICIOS MÉDICOS'...");
			
							contadorc1--;
							contadorc1medint--;
							check.checked=false;
			
							nopasar=true;
						}


				if (contadorc1>2 && tiposervicio==2)
					{
						alert("SOLO SE PERMITE SELECCIONAR DOS ESPECIALIDADES DEL PRIMER CUATRIMESTRE ENTRE LOS SERVICIOS 'QUIRÚRGICOS' Y/O 'MEDICOS'...");
						
						contadorc1--;
						contadorc1medint--;
					
						if (check.checked==true)

							{
								check.checked=false;
							}
							
							 nopasar=true;
					}
													
													
				if (nopasar==false)

					{	
													
						if (servicio=='hem1')
							{
								document.getElementById("idhematologia2").disabled=true;
							}
						
						if (servicio=='neum1')
							{
								document.getElementById("idneumologia2").disabled=true;
							}
						
						if (servicio=='card1')
							{
								document.getElementById("idcardiologia2").disabled=true;
							}
						
						if (servicio=='dige1')
							{
								document.getElementById("iddigestivo2").disabled=true;
							}
						
						if (servicio=='nef1')
							{
								document.getElementById("idnefrologia2").disabled=true;
							}
						
						if (servicio=='mint1')
							{
								document.getElementById("idmedicinainterna2").disabled=true;
							}
						
						if (servicio=='endo1')
							{
								document.getElementById("idendocrinologia2").disabled=true;
							}
						
						if (servicio=='reu1')
							{
								document.getElementById("idreumatologia2").disabled=true;
							}
						
						if (servicio=='onco1')
							{
								document.getElementById("idoncologia2").disabled=true;
							}
						
						if (servicio=='neuro1')
							{
								document.getElementById("idneurologia2").disabled=true;
							}
						
						if (servicio=='infe1')
							{
								document.getElementById("idinfecciosos2").disabled=true;
							}
						
						if (servicio=='minten1')
							{
								document.getElementById("idmedicinaintensiva2").disabled=true;
							}
						
						if (servicio=='espd1')
							{
								document.getElementById("idespecialidadd2").disabled=true;
							}
						
						if (servicio=='espe1')
							{
								document.getElementById("idespecialidade2").disabled=true;
							}
						
						if (servicio=='espf1')
							{
								document.getElementById("idespecialidadf2").disabled=true;
							}
					
					}
			} 
		
			else 

			if (check.checked==false && ciclo==1 && tiposervicio==2)
												
				{
												
					contadorc1--;
					contadorc1medint--;


					if (servicio=='hem1')
						{
							document.getElementById("idhematologia2").disabled=false;
						}
					
					if (servicio=='neum1')
						{
							document.getElementById("idneumologia2").disabled=false;
						}
					
					if (servicio=='card1')
						{
							document.getElementById("idcardiologia2").disabled=false;
						}
					
					if (servicio=='dige1')
						{
							document.getElementById("iddigestivo2").disabled=false;
						}
					
					if (servicio=='nef1')
						{
							document.getElementById("idnefrologia2").disabled=false;
						}
					
					if (servicio=='mint1')
						{
							document.getElementById("idmedicinainterna2").disabled=false;
						}
					
					if (servicio=='endo1')
						{
							document.getElementById("idendocrinologia2").disabled=false;
						}
					
					if (servicio=='reu1')
						{
							document.getElementById("idreumatologia2").disabled=false;
						}
					
					if (servicio=='onco1')
						{
							document.getElementById("idoncologia2").disabled=false;
						}
					
					if (servicio=='neuro1')
						{
							document.getElementById("idneurologia2").disabled=false;
						}
					
					if (servicio=='infe1')
						{
							document.getElementById("idinfecciosos2").disabled=false;
						}
					
					if (servicio=='minten1')
						{
							document.getElementById("idmedicinaintensiva2").disabled=false;
						}
				
					if (servicio=='espd1')
						{
							document.getElementById("idespecialidadd2").disabled=false;
						}
					
					if (servicio=='espe1')
						{
							document.getElementById("idespecialidade2").disabled=false;
						}
					
					if (servicio=='espf1')
						{
							document.getElementById("idespecialidadf2").disabled=false;
						}
				
				}
												
		else
												
		if (check.checked==true && ciclo==2 && tiposervicio==1)
															
			{

				contadorc2++;
				contadorc2cirugia++;
	
				
				if (contadorc1cirugia==2 &&  contadorc2cirugia==1)
					
					
					{
						alert("NO PUEDE SELECCIONAR TRES ESPECIALIDADES DENTROS DE 'SERVICIOS QUIRÚRGICOS'...");
		
						contadorc2--;
						contadorc2cirugia--;
						check.checked=false;
		
						nopasar=true;
					}


				if (contadorc2>1 && tiposervicio==1)
					
					{
						alert("YA SELECCIONO UNA ESPECIALIDAD DEL SEGUNDO CUATRIMESTRE (SEGUNDO CICLO)...");
						
						contadorc2--;
						contadorc2cirugia--;
					
						if (check.checked==true)
							{
								check.checked=false;
							}
							 	nopasar=true;
									
					}
				
				
				if (nopasar==false)

					{

						if (servicio=='gd2')
							{
								document.getElementById("idgeneraldigestivaciclo1").disabled=true;
							}
						
						if (servicio=='car2')
							{
								document.getElementById("idcardca1").disabled=true;
							}
						
						if (servicio=='tor2')
							{
								document.getElementById("idtoracica1").disabled=true;
							}
						
						if (servicio=='max2')
							{
								document.getElementById("idmaxilofacial1").disabled=true;
							}
						
						if (servicio=='pla2')
							{
								document.getElementById("idplastica1").disabled=true;
							}
						
						if (servicio=='vas2')
							{
								document.getElementById("idvascular1").disabled=true;
							}
						
						if (servicio=='neu2')
							{
								document.getElementById("idneurocirugia1").disabled=true;
							}
						
						if (servicio=='tra2')
							{
								document.getElementById("idtraumatologia1").disabled=true;
							}
						
						if (servicio=='uro2')
							{
								document.getElementById("idurologia1").disabled=true;
							}
						
						if (servicio=='ped2')
							{
								document.getElementById("idpediatrica1").disabled=true;
							}
						
						if (servicio=='espa2')
							{
								document.getElementById("idespecialidada1").disabled=true;
							}									
						
						if (servicio=='espb2')
							{
								document.getElementById("idespecialidadb1").disabled=true;
							}									
						
						if (servicio=='espc2')
							{
								document.getElementById("idespecialidadc1").disabled=true;
							}									
					} 
													
			}
												
			else 
												
																	
			if (check.checked==false && ciclo==2 && tiposervicio==1)
												
				{
												
					contadorc2--;
					contadorc2cirugia--;


					if (servicio=='gd2')
						{
							document.getElementById("idgeneraldigestivaciclo1").disabled=false;
						}
					
					if (servicio=='car2')
						{
							document.getElementById("idcardca1").disabled=false;
						}
					
					if (servicio=='tor2')
						{
							document.getElementById("idtoracica1").disabled=false;
						}
					
					if (servicio=='max2')
						{
							document.getElementById("idmaxilofacial1").disabled=false;
						}
					
					if (servicio=='pla2')
						{
							document.getElementById("idplastica1").disabled=false;
						}
					
					if (servicio=='vas2')
						{
							document.getElementById("idvascular1").disabled=false;
						}
					
					if (servicio=='neu2')
						{
							document.getElementById("idneurocirugia1").disabled=false;
						}
					
					if (servicio=='tra2')
						{
							document.getElementById("idtraumatologia1").disabled=false;
						}
					
					if (servicio=='uro2')
						{
							document.getElementById("idurologia1").disabled=false;
						}
					
					if (servicio=='ped2')
						{
							document.getElementById("idpediatrica1").disabled=false;
						}
					if (servicio=='espa2')
						{
							document.getElementById("idespecialidada1").disabled=false;
						}									

					if (servicio=='espb2')
						{
							document.getElementById("idespecialidadb1").disabled=false;
						}									

					if (servicio=='espc2')
						{
							document.getElementById("idespecialidadc1").disabled=false;
						}									

				}
												
			else

									
			if (check.checked==true && ciclo==2 && tiposervicio==2)

				{
					contadorc2++;
					contadorc2medint++;

					if (contadorc1medint==2 && contadorc2medint>0)
	
						{
							alert("NO PUEDE SELECCIONAR TRES ESPECIALIDADES DENTRO DE 'SERVICIOS MÉDICOS'...");
			
							contadorc2--;
							contadorc2medint--;
							check.checked=false;
			
							nopasar=true;
						}
	
	
					if (contadorc2>1 && tiposervicio==2)
				   		{
							alert("YA SELECCIONO UNA ESPECIALIDAD DEL SEGUNDO CUATRIMESTRE (SEGUNDO CICLO)...");
							
							contadorc2--;
							contadorc2medint--;
						
							if (check.checked==true)
										{
													check.checked=false;
										}
										
										 nopasar=true;
						}
													
													
					if (nopasar==false)
													
						{	
																
							if (servicio=='hem2')
								{
									document.getElementById("idhematologia1").disabled=true;
								}
							
							if (servicio=='neum2')
								{
									document.getElementById("idneumologia1").disabled=true;
								}
					
							if (servicio=='card2')
								{
									document.getElementById("idcardiologia1").disabled=true;
								}
							
							if (servicio=='dige2')
								{
									document.getElementById("iddigestivo1").disabled=true;
								}
							
							if (servicio=='nef2')
								{
									document.getElementById("idnefrologia1").disabled=true;
								}
							
							if (servicio=='mint2')
								{
									document.getElementById("idmedicinainterna1").disabled=true;
								}
							
							if (servicio=='endo2')
								{
									document.getElementById("idendocrinologia1").disabled=true;
								}
							
							if (servicio=='reu2')
								{
									document.getElementById("idreumatologia1").disabled=true;
								}
							
							if (servicio=='onco2')
								{
									document.getElementById("idoncologia1").disabled=true;
								}
							
							if (servicio=='neuro2')
								{
									document.getElementById("idneurologia1").disabled=true;
								}
							
							if (servicio=='infe2')
								{
									document.getElementById("idinfecciosos1").disabled=true;
								}
							
							if (servicio=='minten2')
								{
									document.getElementById("idmedicinaintensiva1").disabled=true;
								}

							if (servicio=='espd2')
								{
									document.getElementById("idespecialidadd1").disabled=true;
								}
							
							if (servicio=='espe2')
								{
									document.getElementById("idespecialidade1").disabled=true;
								}
							
							if (servicio=='espf2')
								{
									document.getElementById("idespecialidadf1").disabled=true;
								}
						
						
						
						
						}	
												   
				}
												
			else
												
			if (check.checked==false && ciclo==2 && tiposervicio==2)
												
				{
												
					contadorc2--;
					contadorc2medint--;

					if (servicio=='hem2')
						{
							document.getElementById("idhematologia1").disabled=false;
						}
					
					if (servicio=='neum2')
						{
							document.getElementById("idneumologia1").disabled=false;
						}
					
					if (servicio=='card2')
						{
							document.getElementById("idcardiologia1").disabled=false;
						}
					
					if (servicio=='dige2')
						{
							document.getElementById("iddigestivo1").disabled=false;
						}
					
					if (servicio=='nef2')
						{
							document.getElementById("idnefrologia1").disabled=false;
						}
					
					if (servicio=='mint2')
						{
							document.getElementById("idmedicinainterna1").disabled=false;
						}
					
					if (servicio=='endo2')
						{
							document.getElementById("idendocrinologia1").disabled=false;
						}
					
					if (servicio=='reu2')
						{
							document.getElementById("idreumatologia1").disabled=false;
						}
					
					if (servicio=='onco2')
						{
							document.getElementById("idoncologia1").disabled=false;
						}
					
					if (servicio=='neuro2')
						{
							document.getElementById("idneurologia1").disabled=false;
						}
					
					if (servicio=='infe2')
						{
							document.getElementById("idinfecciosos1").disabled=false;
						}
					
					if (servicio=='minten2')
						{
							document.getElementById("idmedicinaintensiva1").disabled=false;
						}
												
					if (servicio=='espd2')
						{
							document.getElementById("idespecialidadd1").disabled=false;
						}
					
					if (servicio=='espe2')
						{
							document.getElementById("idespecialidade1").disabled=false;
						}
					
					if (servicio=='espf2')
						{
							document.getElementById("idespecialidadf1").disabled=false;
						}
				
				}
				
				
		if (contadorc1==2 && contadorc2==1)
			{
				document.getElementById("idimagefinalizado").style="visibility: visible";
				document.getElementById("mensajeaceptacionformulario").style.display = "block";
				document.getElementById("idaceptaformulario").focus();
			}
		else
			{
				document.getElementById("mensajeaceptacionformulario").style.display = "none";
				document.getElementById("idimagefinalizado").style="visibility: hidden";
			}
	
}



function fecha()
{
	fecha = new Date()
	mes = fecha.getMonth()
	diaMes = fecha.getDate()
	diaSemana = fecha.getDay()
	anio = fecha.getFullYear()
	dias = new Array('Domingo','Lunes','Martes','Miercoles','Jueves','Viernes','Sábado')
	meses = new Array('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre')
	document.write('<span id="fecha">')
	document.write (dias[diaSemana] + ", " + diaMes + " de " + meses[mes] + " de " + anio)
	document.write ('</span>')
	
}


function hora()
{
	var fecha = new Date()
	var hora = fecha.getHours()
	var minuto = fecha.getMinutes()
	var segundo = fecha.getSeconds()
	if (hora < 10) {hora = "0" + hora}
	if (minuto < 10) {minuto = "0" + minuto}
	if (segundo < 10) {segundo = "0" + segundo}
	var hora_completa = hora + ":" + minuto + ":" + segundo
	document.getElementById('hora').firstChild.nodeValue = hora_completa
	tiempo = setTimeout('hora()',1000)
}


function inicio()
{
	document.write('<span id="hora">')
	document.write ('000000</span>')
	hora()
}



function habilitabotonactualizarficha()
{
	if(document.formularioaltapracticaclinica.AceptaFormularioOk.checked == true)            
		{
			document.getElementById("idenviardatos").style="visibility:visible";
		} 
	else 
		{
			document.getElementById("idenviardatos").style="visibility:hidden";
		}
}					
				




function valida_dni()

{
			valor = document.formularioaltapracticaclinica.DNIAlumnoa.value;
			valor=valor.toUpperCase();
			
			var letras = ['T', 'R', 'W', 'A', 'G', 'M', 'Y', 'F', 'P', 'D', 'X', 'B', 'N', 'J', 'Z', 'S', 'Q', 'V', 'H', 'L', 'C', 'K', 'E', 'T'];
			
			var huboerror1 = false;
			
			
			if( !(/^\d{8}[A-Z]$/.test(valor)) ) 
			
						{
									if (document.formularioaltapracticaclinica.DNIAlumnoa.value.length<9)
									{
									huboerror1 = true;
									alert("DEBE DE ESCRIBIR SU 'DNI' COMPLETO. (8 CARACTERES MAS LETRA).");
									}

									document.getElementById("identradaerronea1").style='visibility:visible';
									document.getElementById("identradaerronea1").innerHTML="<td width='32' id='identradaerronea1'><img src='pngs/error_32_x_32.png'/><td>";
									
									document.formularioaltapracticaclinica.DNIAlumnoa.focus();
						}
			 
			 
			if(valor.charAt(8) != letras[(valor.substring(0, 8))%23]) 
			
						{
									if (document.formularioaltapracticaclinica.DNIAlumnoa.value.length>8)
									{
									huboerror1 = true;
									alert("LA LETRA DE SU DNI NO CORRESPONDE CON LOS DIGITOS ESCRITOS...");
									}
									
									document.getElementById("identradaerronea1").style='visibility:visible';
									document.getElementById("identradaerronea1").innerHTML="<td width='32' id='identradaerronea1'><img src='pngs/error_32_x_32.png'/><td>";
									
									document.formularioaltapracticaclinica.DNIAlumnoa.focus();
						}

			if (huboerror1==false)
						
						{									
									document.getElementById("identradaerronea1").style='visibility:hidden';
									
									document.getElementById("identradacorrecta1").style='visibility:visible';
									document.getElementById("identradacorrecta1").innerHTML="<td width='32' id='identradacorrecta1'><img src='pngs/acepta_32_x_32.png'/><td>";
									
									document.formularioaltapracticaclinica.botonvalidardni.style='visibility:hidden';
									document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:visible';
									
									document.formularioaltapracticaclinica.NombreAlumnoa.disabled=false;
									document.formularioaltapracticaclinica.NombreAlumnoa.focus();
						}

}


function reentrar_dni()

{
			document.getElementById("identradacorrecta1").style='visibility:hidden';
			
			document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:hidden';
						
			document.formularioaltapracticaclinica.botonvalidardni.style='visibility:visible';
			
			document.formularioaltapracticaclinica.DNIAlumnoa.focus();
}





function valida_nombre()

{
		
				if (document.formularioaltapracticaclinica.NombreAlumnoa.value.length==0)
									{
									   alert("DEBE DE ESCRIBIR SU 'NOMBRE'...");
									   
									   document.getElementById("identradaerronea2").style='visibility:visible';
									   document.getElementById("identradaerronea2").innerHTML="<td width='32' id='identradaerronea2'><img src='pngs/error_32_x_32.png'/><td>";
								 
									   document.formularioaltapracticaclinica.NombreAlumnoa.focus();
									   return 0;
									}
			
				else
				
									{
									
									   document.getElementById("identradaerronea2").style='visibility:hidden';
																
									   document.getElementById("identradacorrecta2").style='visibility:visible';
									   document.getElementById("identradacorrecta2").innerHTML="<td width='32' id='identradacorrecta2'><img src='pngs/acepta_32_x_32.png'/><td>";
																
									   document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:hidden';
									   document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:visible';
																
									  
									   document.formularioaltapracticaclinica.ApellidosAlumnoa.disabled=false;
									   document.formularioaltapracticaclinica.ApellidosAlumnoa.focus();
									}							   
									  


} 






function reentrar_nombre()

{

			document.getElementById("identradacorrecta2").style='visibility:hidden';
			
			document.formularioaltapracticaclinica.botonvalidardni.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:hidden';
						
			document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:visible';
			
			document.formularioaltapracticaclinica.NombreAlumnoa.focus();
}





function valida_apellidos()
{

				if (document.formularioaltapracticaclinica.ApellidosAlumnoa.value.length==0)
									{
									   alert("DEBE DE ESCRIBIR SUS 'APELLIDOS'...");
									   
									   document.getElementById("identradaerronea3").style='visibility:visible';
									   document.getElementById("identradaerronea3").innerHTML="<td width='32' id='identradaerronea3'><img src='pngs/error_32_x_32.png'/><td>";
								 
									   document.formularioaltapracticaclinica.ApellidosAlumnoa.focus();
									   return 0;
									}
			
				else
				
									{
									
									   document.getElementById("identradaerronea3").style='visibility:hidden';
																
									   document.getElementById("identradacorrecta3").style='visibility:visible';
									   document.getElementById("identradacorrecta3").innerHTML="<td width='32' id='identradacorrecta3'><img src='pngs/acepta_32_x_32.png'/><td>";
																
									   document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:hidden';
									   document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:visible';
																
									   document.formularioaltapracticaclinica.EmailAlumnoa.disabled=false;
									   document.formularioaltapracticaclinica.EmailAlumnoa.focus();
									
									}

} 



function reentrar_apellidos()

{
			document.getElementById("identradacorrecta3").style='visibility:hidden';
			
			document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:visible';
			
			document.formularioaltapracticaclinica.botonvalidardni.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:hidden';
			
			document.formularioaltapracticaclinica.ApellidosAlumnoa.focus();
}






function valida_mail(EmailAlumnoa) 

{

     						vmail = document.getElementById("idemailalumnoa").value;
		
							var filter=/^[A-Za-z][A-Za-z0-9_.]*@[A-Za-z0-9_]+.[A-Za-z0-9_.]+[A-za-z]$/;

	
							if (vmail.length == 0 ) 
							
										{
											   alert("DEBE DE ESCRIBIR UNA DIRECCIÓN DE 'E-MAIL'...");
												
											   document.getElementById("identradaerronea4").style='visibility:visible';
											   document.getElementById("identradaerronea4").innerHTML="<td width='32' id='identradaerronea4'><img src='pngs/error_32_x_32.png'/><td>";
										 
											   document.formularioaltapracticaclinica.EmailAlumnoa.focus();
					
												return true;
												
										}
		
							if (filter.test(vmail))
							
										{
													document.getElementById("identradaerronea4").style='visibility:hidden';
																													
													document.getElementById("identradacorrecta4").style='visibility:visible';
													document.getElementById("identradacorrecta4").innerHTML="<td width='32' id='identradacorrecta4'><img src='pngs/acepta_32_x_32.png'/><td>";
																			
													document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:hidden';
																			
													document.formularioaltapracticaclinica.GeneralDigestivaCiclo1.focus();
													
													return true;
										}
								
		
							else
		
										{
													alert("LA DIRECCIÓN DE CORREO '"+vmail+"' ESTA MAL ESCRITA...");
												
													document.getElementById("identradaerronea4").style='visibility:visible';
													document.getElementById("identradaerronea4").innerHTML="<td width='32' id='identradaerronea4'><img src='pngs/error_32_x_32.png'/><td>";
															 
												    document.formularioaltapracticaclinica.EmailAlumnoa.focus();
												    return false;
									   }
				
								
} 





function reentrar_email()

{
			document.getElementById("identradacorrecta4").style='visibility:hidden';
			
			document.formularioaltapracticaclinica.botonvalidarcampoemail.style='visibility:visible';
			
			document.formularioaltapracticaclinica.botonvalidardni.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarnombre.style='visibility:hidden';
			document.formularioaltapracticaclinica.botonvalidarapellidos.style='visibility:hidden';
			
			document.formularioaltapracticaclinica.EmailAlumnoa.focus();
}




















window.onload = function()


{
	
	
	procesar('panel_carga')
	
	
	document.formularioaltapracticaclinica.FechaAsignacionPracticas.value = (dias[diaSemana] + ", " + diaMes + " de " + meses[mes] + " de " + anio);

	var fecha = new Date()
	var hora = fecha.getHours()
	var minuto = fecha.getMinutes()
	var segundo = fecha.getSeconds()
	if (hora < 10) {hora = "0" + hora}
	if (minuto < 10) {minuto = "0" + minuto}
	if (segundo < 10) {segundo = "0" + segundo}
	var hora_completa = hora + ":" + minuto + ":" + segundo
	document.getElementById('hora').firstChild.nodeValue = hora_completa
	tiempo = setTimeout('hora()',1000)

	document.formularioaltapracticaclinica.HoraAsignacionPracticas.value = hora_completa;
	
	document.getElementById("iddnialumnoa").focus();
	
	
	
	
	
	
}







function stopRKey(evt) 
					
									{
												var evt = (evt) ? evt : ((event) ? event : null);
												var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
												
												if ((evt.keyCode == 13) && (node.type=="text")) 
												
															{
																		return false;
															}
									}
					
									document.onkeypress = stopRKey;
									
	
	
	
	
	
	
	
	
	
function procesar(idPanelCarga)

{
			 //Obtener referencia a DIV de panel de carga
			 var panelCarga = document.getElementById(idPanelCarga);
			 
			 //Mostrar DIV de panel de carga
			 panelCarga.style.visibility = "visible";
			 
			 //Ocultar DIV de panel de carga, después de n segundos
			 setTimeout(function() { panelCarga.style.visibility = "hidden"; }, 1500);
}
	
	
	
	


						
									
	
	

</script>



<title>GESTIÓN DE PRÁCTICAS CLÍNICAS</title>

</head>

<body leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0"  bgcolor="#FFFFFF" >



<%
	
	Dim ObjConn	
	Dim ObjRs    	
	
	Dim ClaveAlumnoPasaraDB
	
	Dim FechaGrabacionAsignacion
	Dim HoraGrabacionAsignacion


if Request.Form="" then 
%>



<div id="area_invalidar" class="clase_area_invalidar">

<div id="panel_carga" class="clase_panel_carga"></div>



<form name="formularioaltapracticaclinica" method="post" action="grabardatosalumnoa.asp" autocomplete="off">


<%
Dim SqlContarRegistros
Dim TotalRegistrosExistentesTabla

SqlContarRegistros = "SELECT * FROM tbl_Alumnos"

Set ObjConn = Server.CreateObject("ADODB.Connection")
Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjConn.Open "dsndbgpc"

ObjRs.CursorType = 1

ObjRs.Open SqlContarRegistros, ObjConn

TotalRegistrosExistentesTabla=ObjRs.RecordCount


ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>



<table width="1287" align="center" border="0">
<tr>

<%
		If TotalRegistrosExistentesTabla=0 Then
            
			%>
			<td width="72"  align="center"><a href="lstalumnoascse.asp" target="_blank"><img src="pngs/infopantalla_64_x_64_no.png"/></a></td>
	  		<td width="104" class="textomenuinicio1" align="left"><a href="lstalumnoascse.asp" target="_blank">Listado Alumnos/as por Ciclos, Servicios y Especialidades por Pantalla e Impresora</a></td>
<%
			
		Else
		
			%>
            <td width="72"  align="center"><a href="lstalumnoascse.asp" target="_blank"><img src="pngs/infopantalla_64_x_64_si.png"/></a></td>
	  		<td width="104" class="textomenuinicio1" align="left"><a href="lstalumnoascse.asp" target="_blank">Listado Alumnos/as por Ciclos, Servicios y Especialidades por Pantalla  e Impresora</a></td>
      <%
		 
		End If
%>       
         


<%
	Dim DirName
	DirName="/gpc/pdfs"
	
	Dim Fso
	Dim Folder
	Dim FileCount
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Folder = Fso.GetFolder(server.mappath(DirName))
	
	FileCount = Folder.Files.Count
	
	If FileCount=0 Then

            %>
			<td width="72"  align="center"><a href="verfolderfilespdf.asp" target="_blank"><img src="pngs/folderpdfs_64_x_64_no.png"/></a></td>
	  		<td width="78" class="textomenuinicio1" align="left"><a href="verfolderfilespdf.asp" target="_blank">Acceso a Carpeta de ficheros PDF emitidos por los Alumnos/as</a></td>
<%

	Else

			%>
            <td width="72"  align="center"><a href="verfolderfilespdf.asp" target="_blank"><img src="pngs/folderpdfs_64_x_64_si.png"/></a></td>
	  		<td width="78" class="textomenuinicio1" align="left"><a href="verfolderfilespdf.asp" target="_blank">Acceso a Carpeta de ficheros PDF emitidos por los Alumnos/as</a></td>
<%

	End If
%>

               
            <td width="72" align="center"><a href="admonplazasesp.asp" target="_blank"><img src="pngs/administradores_64_x_64.png" /></a></td>
            
<!--
            <td width="78" align="center"><a href="enconstruccion.asp" target="_blank"><img src="pngs/administradores_64_x_64.png" /></a></td>
            -->
	  		
            <td width="102" class="textomenuinicio1" align="left"><a href="admonplazasesp.asp" target="_blank">Área de Administrador. (Gestión de Número de Plazas en Especialidades en la Base de Datos)</a></td>
<!--
            <td width="320" class="textomenuinicio1" align="left"><a href="enconstruccion.asp" target="_blank">Área de Administrador. (Gestión de Número de Plazas en Especialidades en la Base de Datos)</a></td>
            -->
    
    
  
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
	
	For Each x in FoldersExistentes2.SubFolders
				
				'NombreSubcarpeta = x.Name
				FolderCount2 = FolderCount2+1
				'Response.Write NombreSubcarpeta & "<br>"
	Next
	
	If FolderCount2=0 Then

            %>
            <td width="77" align="center"><a href="selechistorico.asp" target="_blank"><img src="pngs/cajavacia_1_95_x_98.png" /></a></td>
            <td width="126" class="textomenuinicio1" align="left"><a href="selechistorico.asp" target="_blank">Acceso a Histórico de Promociones de Alumnos/as de Años Anteriores</a></td>
<%
	Else
%>             
             
      		<td width="105" align="center"><a href="selechistorico.asp" target="_blank"><img src="pngs/pdftohistorico_1_95_x_68.png" /></a></td>
      <td width="95" class="textomenuinicio1" align="left"><a href="selechistorico.asp" target="_blank">Acceso a Histórico de Promociones de Alumnos/as de Años Anteriores</a></td>
    
<%
	End If
%>  
    
    
    
    
    </tr>
</table>



  
<table width="800" border="0" align="center">
<tr>
    <td width="800" align="center"><img src="jpg/logougrfmed_300_x_185.jpg" width="300" height="185" /></td>
</tr>
</table>


<table width="800" align="center" border="0">
<tr>
<td width="850" scope="col" class="textocabecera1"><div align="center">PATOLOGÍAS MÉDICO-QUIRÚRGICAS I Y II</div></td>
</tr>
</table>



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

TotalRegistrosTabla=ObjRs.RecordCount+1

CadenaNumRegistroActual=String(3-Len(CStr(TotalRegistrosTabla)),"0")+CStr(TotalRegistrosTabla)


ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>

<br/>

<table width="1100" align="center" border="0">

    <tr>
	  <td width="299" height="12" align="left" class="textotitulosubcabeceras1">Nº ASIGNACIÓN DE PRÁCTICAS:</td>
                        <td width="66" class="campotexto1" align="center"><%Response.Write(CadenaNumRegistroActual)%></td>
                        
                        <td width="274" class="textotitulosubcabeceras1">FECHA Y HORA ASIGNACIÓN:</td>
                        
                        <td width="344" class="campotextofechahora"><div align="center"><script>fecha()</script></div></td>
                        <td width="74" class="campotextofechahora"><div align="center"><script>inicio()</script></div></td>
            </tr>
</table>


<table width="1100" align="center" border="0">
    <tr>
	  <td width="182" ><input type="hidden" name="ClaveAlumnoPasaraDB" value="<%=CadenaNumRegistroActual%>"/></td>
                        <td width="542" ><input type="hidden" name="FechaAsignacionPracticas" id="idfecha" size=75  value=""/></td>
                        <td width="362" ><input type="hidden" name="HoraAsignacionPracticas" id="idhora" value=""/></td>
            <tr>
</table>


<table width="1100" align="center" border="0">
	<tr>
	  <td width="248" height="12" align="left" class="textotitulosubcabeceras1">DATOS DEL ALUMNO/A.</td>
            </tr>
</table>


<table width="1100" align="center" border="0">

	<tr>
                		<td width="342" height="34" align="right" class="textotitulocampo1">DNI:</td>
	  					<td width="10"></td>
	  					<td width="543" height="34" align="left"><input type="text" class="campotexto1" name="DNIAlumnoa" id="iddnialumnoa" size="9" maxlength="9" value="" onfocus="reentrar_dni()"/></td>
                                    
	  					<td width="32" id="identradacorrecta1"></td>
      					<td width="32" id="identradaerronea1"></td>
                                   
                        <td width="115" align="right"><input type="button" class="" name="botonvalidardni" id="idbotonvalidardni" value="Validar y Aceptar" onclick="valida_dni()"/></td>
	</tr>


    <tr>
	                    <td width="342" height="12" align="right" class="textotitulocampo1">NOMBRE:</td>
                        <td width="10"></td>
                        <td width="543" height="12" align="left"><input type="text" class="campotexto1" name="NombreAlumnoa" size="20" maxlength="20" value=""  disabled="disabled" onfocus="reentrar_nombre()" /></td>
                        
	  					<td width="32" id="identradacorrecta2"></td>
      					<td width="32" id="identradaerronea2"></td>
                        
                        <td width="115" align="right"><input type="button" class="" name="botonvalidarnombre" id="idbotonvalidarnombre" value="Validar y Aceptar" style="visibility: hidden" onclick="valida_nombre()"/></td>
    </tr>


    <tr>
      <td width="342" height="12" align="right" class="textotitulocampo1">APELLIDOS:</td>
      <td width="10"></td>
      <td width="543" height="12" align="left"><input type="text" class="campotexto1" name="ApellidosAlumnoa" size="25" maxlength="25" value=""  disabled="disabled" onfocus="reentrar_apellidos()"/></td>
                        
	  					<td width="32" id="identradacorrecta3"></td>
      					<td width="32" id="identradaerronea3"></td>
                        
                        
      <td width="115" align="right"><input type="button" class="" name="botonvalidarapellidos" id="idbotonvalidarapellidos" value="Validar y Aceptar" style="visibility: hidden" onclick="valida_apellidos()"/></td>
    </tr>

   
   
    <tr>
      <td width="342" height="34" align="right" class="textotitulocampo1">DIRECCIÓN DE CORREO E-MAIL:</td>
      <td width="10"></td>

      <td width="543" height="34" align="left"><input type="text" value="" class="campoemail" name="EmailAlumnoa" id="idemailalumnoa"size="50" maxlength="50" disabled="disabled" onfocus="reentrar_email()" /></td>
 

 	  				<td width="32" id="identradacorrecta4"></td>
  					<td width="32" id="identradaerronea4"></td>
                        
      <td width="115" align="right"><input type="button" class="" name="botonvalidarcampoemail" id="idbbotonvalidarcampoemail" value="Validar y Aceptar" style="visibility: hidden" onclick="valida_mail(EmailAlumnoa)" /></td>
    </tr>
</table>



<br />



<table width="1100" align="center" border="0">
<tr>
<td width="1016" height="12" align="left" class="textotitulosubcabeceras1">SELECCIÓN DE ESPECIALIDADES EN LOS SERVICIOS QUIRÚRGICOS</td>
</tr>
</table>



<table width="1100" align="center" border="1" cellspacing="0">

<tr>
<td width="378" height="12" align="center" class="textotitulocampo3">1º CICLO [Primer Cuatrimestre]</td>
<td width="412" height="12" align="center" class="textotitulocampo3">2º CICLO [Segundo Cuatrimestre]</td>
</tr>



<tr>
<td width="378" height="12" align="center" class="textotitulocampo3"></td>
<td width="412" height="12" align="center" class="textotitulocampo3"></td>
</tr>


</table>



<br />



<table width="1100" align="center" border="0" cellspacing="0">

<tr>

<td width="171" height="12" align="center" class="textotitulocampo3"></td>
<td width="21" height="12" align="left" class="textotitulocampo3"></td>
<td width="79" height="12" align="left" class="textotitulocampo3"></td>
<td width="107" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="144" height="12" align="center" class="textotitulocampo3">ESTADO DEL</th>
<td width="40" height="12" align="center" class="textotitulocampo3"></td>
<td width="125" height="12" align="left" class="textotitulocampo3"></td>
<td width="134" height="12" align="left" class="textotitulocampo3"></td>
<td width="107" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
<td width="152" height="12" align="center" class="textotitulocampo3">ESTADO DEL </th></tr>

<tr>
<td width="171" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="21" height="12" align="left" class="textotitulocampo3"></td>
<td width="79" height="12" align="left" class="textotitulocampo3"></td>
<td width="107" height="12" align="center" class="textotitulocampo3">DISPONIBLES</td>
<td width="144" height="12" align="center" class="textotitulocampo3">SERVICIO</th>
<td width="40" height="12" align="center" class="textotitulocampo3"></td>
<td width="125" height="12" align="left" class="textotitulocampo3">ESPECIALIDAD</td>
<td width="134" height="12" align="left" class="textotitulocampo3"></td>
<td width="107" height="12" align="center" class="textotitulocampo3">DISPONIBLES</td>
<td width="152" height="12" align="center" class="textotitulocampo3">SERVICIO</td>

</tr>

</table>


<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
'ObjConn.Open "dsndbgpc"
ObjConn.Open="driver={Microsoft Access Driver (*.mdb)}; dbq=c:\inetpub\wwwroot\gpc\data\dbgpc.mdb"


Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn, 3, 3

Dim SQLGDIG1
SQLGDIG1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='GDIG1'"

Set ObjRs=ObjConn.Execute(SQLGDIG1)
%>

<table width="1100" align="center" border="0">

<tr>
<td width="169" height="12" align="right" class="textotitulocampo1">General-Digestiva:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>

      <td width="20" height="12" align="center"><input type='checkbox' name="GeneralDigestivaCiclo1" value="1"  id="idgeneraldigestivaciclo1" onclick='validar(formularioaltapracticaclinica.GeneralDigestivaCiclo1,1,1,"gd1")' /><label for="idgeneraldigestivaciclo1"></label></td>
<%
Else
%>
	  <td width="20" height="12" align="center"><input type='checkbox' name="idgeneraldigestivaciclo" value="1"  id="idgeneraldigestivaciclo1" style="visibility: hidden" /></td>
	  <%
End If
%>


<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>


<%
Dim EstadoServicioGDIG1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then 

		EstadoServicioGDIG1="CERRADO"

		%> 
		<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioGDIG1)%></td>
		<%

Else

		EstadoServicioGDIG1="DISPONIBLE"

		%>
		<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioGDIG1)%></td>
	  <%

End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 


Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn, 3, 3

Dim SQLGDIG2
SQLGDIG2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='GDIG2'"

Set ObjRs=ObjConn.Execute(SQLGDIG2)
%>


<td width="24" height="12" align="left"></td>

<td width="169" height="12" align="right" class="textotitulocampo1">General-Digestiva:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0  then 
%>

			<td width="28" height="12" align="center"><input type='checkbox' name="GeneralDigestivaCiclo2" value="1" id="idgeneraldigestivaciclo2" onclick='validar(formularioaltapracticaclinica.GeneralDigestivaCiclo2,2,1,"gd2")'><label for="idgeneraldigestivaciclo2"></label></td>
<%
Else
%>
	 		 <td width="20" height="12" align="center"><input type='checkbox' name="GeneralDigestivaCiclo2" value="1" id="idgeneraldigestivaciclo2" style="visibility:hidden" /></td>

<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioGDIG2

if ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

%>


<%
EstadoServicioGDIG2="CERRADO"
%> 

<td width="117" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioGDIG2)%></td>
<%
else
EstadoServicioGDIG2="DISPONIBLE"
%>
<td width="104" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioGDIG2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLCARD1
SQLCARD1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='CARD1'"

Set ObjRs=ObjConn.Execute(SQLCARD1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Cardiaca:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="CardiacaCiclo1" value="1" id="idcardca1" onclick='validar(formularioaltapracticaclinica.CardiacaCiclo1,1,1,"car1")'><label for="idcardca1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="CardiacaCiclo1" value="1" id="idcardca1" style="visibility:hidden" /></td>
<%
End If
%>


<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>


<%
Dim EstadoServicioCARD1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioCARD1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioCARD1)%></td>
			<%
Else
			EstadoServicioCARD1="DISPONIBLE"
			%>
           	<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioCARD1)%></td>
<%
End if
%>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>








<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLCARD2
SQLCARD2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='CARD2'"

Set ObjRs=ObjConn.Execute(SQLCARD2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Cardiaca:</td>


<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

			<td width="28" height="12" align="center"><input type='checkbox' name="CardiacaCiclo2" value="1" id="idcardca2" onclick='validar(formularioaltapracticaclinica.CardiacaCiclo2,2,1,"car2")' /><label for="idcardca2"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="CardiacaCiclo2" value="1" id="idcardca2" style="visibility:hidden" /></td>

<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioCARD2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioCARD2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioCARD2)%></td>
<%
Else
			EstadoServicioCARD2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioCARD2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>






















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLTORA1
SQLTORA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='TORA1'"

Set ObjRs=ObjConn.Execute(SQLTORA1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Torácica:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="ToracicaCiclo1" value="1" id="idtoracica1" onclick='validar(formularioaltapracticaclinica.ToracicaCiclo1,1,1,"tor1")' /><label for="idtoracica1"></label></td>

<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="ToracicaCiclo1" value="1" id="idtoracica1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioTORA1

if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

EstadoServicioTORA1="CERRADO"
%> 
<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioTORA1)%></td>
<%
else
EstadoServicioTORA1="DISPONIBLE"
%>
<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioTORA1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLTORA2
SQLTORA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='TORA2'"

Set ObjRs=ObjConn.Execute(SQLTORA2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Torácica:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

			<td width="28" height="12" align="center"><input type='checkbox' name="ToracicaCiclo2" value="1" id="idtoracica2" onclick='validar(formularioaltapracticaclinica.ToracicaCiclo2,2,1,"tor2")' /><label for="idtoracica2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="ToracicaCiclo2" value="1" id="idtoracica2" style="visibility:hidden"/></td>

<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioTORA2

if ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

EstadoServicioTORA2="CERRADO"
%> 
<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioTORA2)%></td>
<%
else
EstadoServicioTORA2="DISPONIBLE"
%>
<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioTORA2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMAXI1
SQLMAXI1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MAXI1'"

Set ObjRs=ObjConn.Execute(SQLMAXI1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Maxilofacial:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="MaxilofacialCiclo1" value="1" id="idmaxilofacial1" onclick='validar(formularioaltapracticaclinica.MaxilofacialCiclo1,1,1,"max1")' /><label for="idmaxilofacial1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="MaxilofacialCiclo1" value="1" id="idmaxilofacial1" style="visibility:hidden"/></td>
<%
End If
%>


<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>


<%
Dim EstadoServicioMAXI1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioMAXI1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMAXI1)%></td>
<%
Else
			EstadoServicioMAXI1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMAXI1)%></td>
<%
End if
%>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMAXI2
SQLMAXI2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MAXI2'"

Set ObjRs=ObjConn.Execute(SQLMAXI2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Maxilofacial:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="MaxilofacialCiclo2" value="1" id="idmaxilofacial2" onclick='validar(formularioaltapracticaclinica.MaxilofacialCiclo2,2,1,"max2")' /><label for="idmaxilofacial2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="MaxilofacialCiclo2" value="1" id="idmaxilofacial2" style="visibility:hidden"/></td>

<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioMAXI2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then


            EstadoServicioMAXI2="CERRADO"
            %> 
            <td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMAXI2)%></td>
<%
Else
			EstadoServicioMAXI2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMAXI2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLPLAS1
SQLPLAS1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='PLAS1'"

Set ObjRs=ObjConn.Execute(SQLPLAS1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Plástica:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="PlasticaCiclo1" value="1" id="idplastica1" onclick='validar(formularioaltapracticaclinica.PlasticaCiclo1,1,1,"pla1")' /><label for="idplastica1"></label></td>
<%
Else
%>

			<td width="20" height="12" align="center"><input type='checkbox' name="PlasticaCiclo1" value="1" id="idplastica1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioPLAS1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioPLAS1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioPLAS1)%></td>
<%
Else
			EstadoServicioPLAS1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioPLAS1)%></td>

<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 

ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")

ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLPLAS2
SQLPLAS2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='PLAS2'"

Set ObjRs=ObjConn.Execute(SQLPLAS2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Plástica:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

			<td width="28" height="12" align="center"><input type='checkbox' name="PlasticaCiclo2" value="1" id="idplastica2" onclick='validar(formularioaltapracticaclinica.PlasticaCiclo2,2,1,"pla2")' /><label for="idplastica2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="PlasticaCiclo2" value="1" id="idplastica2" style="visibility:hidden"/></td>
<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioPLAS2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioPLAS2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioPLAS2)%></td>
<%
Else
			EstadoServicioPLAS2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioPLAS2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLVASC1
SQLVASC1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='VASC1'"

Set ObjRs=ObjConn.Execute(SQLVASC1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Vascular:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="VascularCiclo1" value="1" id="idvascular1" onclick='validar(formularioaltapracticaclinica.VascularCiclo1,1,1,"vas1")' /><label for="idvascular1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="VascularCiclo1" value="1" id="idvascular1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioVASC1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioVASC1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioVASC1)%></td>
<%
Else
			EstadoServicioVASC1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioVASC1)%></td>
<%
End if
%>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLVASC2
SQLVASC2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='VASC2'"

Set ObjRs=ObjConn.Execute(SQLVASC2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Vascular:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

			<td width="28" height="12" align="center"><input type='checkbox' name="VascularCiclo2" value="1" id="idvascular2" onclick='validar(formularioaltapracticaclinica.VascularCiclo2,2,1,"vas2")' /><label for="idvascular2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="VascularCiclo2" value="1" id="idvascular2" style="visibility:hidden"/></td>

<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioVASC2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioVASC2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioVASC2)%></td>
<%
Else
			EstadoServicioVASC2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioVASC2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>
















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEUR1
SQLNEUR1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEUR1'"

Set ObjRs=ObjConn.Execute(SQLNEUR1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Neurocirugía:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>

			<td width="20" height="12" align="center"><input type='checkbox' name="NeurocirugiaCiclo1" value="1" id="idneurocirugia1" onclick='validar(formularioaltapracticaclinica.NeurocirugiaCiclo1,1,1,"neu1")' /><label for="idneurocirugia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="NeurocirugiaCiclo1" value="1" id="idneurocirugia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioNEUR1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioNEUR1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEUR1)%></td>
<%
Else
			EstadoServicioNEUR1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEUR1)%></td>
<%
End if
%>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEUR2
SQLNEUR2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEUR2'"

Set ObjRs=ObjConn.Execute(SQLNEUR2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neurocirugía:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="NeurocirugiaCiclo2" value="1" id="idneurocirugia2" onclick='validar(formularioaltapracticaclinica.NeurocirugiaCiclo2,2,1,"neu2")' /><label for="idneurocirugia2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="NeurocirugiaCiclo2" value="1" id="idneurocirugia2" style="visibility:hidden"/></td>
<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioNEUR2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioNEUR2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEUR2)%></td>
<%
Else
			EstadoServicioNEUR2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEUR2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>


















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLTRAU1
SQLTRAU1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='TRAU1'"

Set ObjRs=ObjConn.Execute(SQLTRAU1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Traumatología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="TraumatologiaCiclo1" value="1" id="idtraumatologia1" onclick='validar(formularioaltapracticaclinica.TraumatologiaCiclo1,1,1,"tra1")' /><label for="idtraumatologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="TraumatologiaCiclo1" value="1" id="idtraumatologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioTRAU1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioTRAU1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioTRAU1)%></td>
<%
Else
			EstadoServicioTRAU1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioTRAU1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLTRAU2
SQLTRAU2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='TRAU2'"

Set ObjRs=ObjConn.Execute(SQLTRAU2)
%>


<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Traumatología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
            <td width="28" height="12" align="center"><input type='checkbox' name="TraumatologiaCiclo2" value="1" id="idtraumatologia2" onclick='validar(formularioaltapracticaclinica.TraumatologiaCiclo2,2,1,"tra2")' /><label for="idtraumatologia2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="TraumatologiaCiclo2" value="1" id="idtraumatologia2" style="visibility:hidden"/></td>
<%
End If
%>


<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>


<%
Dim EstadoServicioTRAU2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioTRAU2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioTRAU2)%></td>
<%
Else
			EstadoServicioTRAU2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioTRAU2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>



























<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLUROL1
SQLUROL1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='UROL1'"

Set ObjRs=ObjConn.Execute(SQLUROL1)
%>


<tr>
<td height="12" class="textotitulocampo1" align="right">Urología:</td>


<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="UrologiaCiclo1" value="1" id="idurologia1" onclick='validar(formularioaltapracticaclinica.UrologiaCiclo1,1,1,"uro1")' /><label for="idurologia1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="UrologiaCiclo1" value="1" id="idurologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioUROL1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

		EstadoServicioUROL1="CERRADO"
		%> 
		<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioUROL1)%></td>
<%
Else
		EstadoServicioUROL1="DISPONIBLE"
		%>
		<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioUROL1)%></td>
<%
End if
%>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLUROL2
SQLUROL2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='UROL2'"

Set ObjRs=ObjConn.Execute(SQLUROL2)
%>

<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Urología:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="UrologiaCiclo2" value="1" id="idurologia2" onclick='validar(formularioaltapracticaclinica.UrologiaCiclo2,2,1,"uro2")' /><label for="idurologia2"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="UrologiaCiclo2" value="1" id="idurologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioUROL2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

		EstadoServicioUROL2="CERRADO"
		%> 
		<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioUROL2)%></td>
<%
Else
		EstadoServicioUROL2="DISPONIBLE"
		%>
		<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioUROL2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLPEDI1
SQLPEDI1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='PEDI1'"

Set ObjRs=ObjConn.Execute(SQLPEDI1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Pediátrica:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="PediatricaCiclo1" value="1" id="idpediatrica1" onclick='validar(formularioaltapracticaclinica.PediatricaCiclo1,1,1,"ped1")' /><label for="idpediatrica1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="PediatricaCiclo1" value="1" id="idpediatrica1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioPEDI1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioPEDI1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioPEDI1)%></td>
<%
Else
			EstadoServicioPEDI1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioPEDI1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLPEDI2
SQLPEDI2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='PEDI2'"

Set ObjRs=ObjConn.Execute(SQLPEDI2)
%>

<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Pediátrica:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="PediatricaCiclo2" value="1" id="idpediatrica2" onclick='validar(formularioaltapracticaclinica.PediatricaCiclo2,2,1,"ped2")' /><label for="idpediatrica2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="PediatricaCiclo2" value="1" id="idpediatrica2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioPEDI2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioPEDI2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioPEDI2)%></td>
<%
Else
			EstadoServicioPEDI2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioPEDI2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>




<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->

<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPA1
SQLESPA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPA1'"

Set ObjRs=ObjConn.Execute(SQLESPA1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - A:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadACiclo1" value="1" id="idespecialidada1" onclick='validar(formularioaltapracticaclinica.EspecialidadACiclo1,1,1,"espa1")' /><label for="idespecialidada1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadACiclo1" value="1" id="idespecialidada1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPA1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPA1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPA1)%></td>
<%
Else
			EstadoServicioESPA1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPA1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPA2
SQLESPA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPA2'"

Set ObjRs=ObjConn.Execute(SQLESPA2)
%>

<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - A:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="EspecialidadACiclo2" value="1" id="idespecialidada2" onclick='validar(formularioaltapracticaclinica.EspecialidadACiclo2,2,1,"espa2")' /><label for="idespecialidada2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadACiclo2" value="1" id="idespecialidada2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPA2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPA2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPA2)%></td>
<%
Else
			EstadoServicioESPA2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPA2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPB1
SQLESPB1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPB1'"

Set ObjRs=ObjConn.Execute(SQLESPB1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - B:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadBCiclo1" value="1" id="idespecialidadb1" onclick='validar(formularioaltapracticaclinica.EspecialidadBCiclo1,1,1,"espb1")' /><label for="idespecialidadb1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadBCiclo1" value="1" id="idespecialidadb1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPB1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPB1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPB1)%></td>
<%
Else
			EstadoServicioESPB1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPB1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPB2
SQLESPB2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPB2'"

Set ObjRs=ObjConn.Execute(SQLESPB2)
%>

<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - B:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="EspecialidadBCiclo2" value="1" id="idespecialidadb2" onclick='validar(formularioaltapracticaclinica.EspecialidadBCiclo2,2,1,"espb2")' /><label for="idespecialidadb2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadBCiclo2" value="1" id="idespecialidadb2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPB2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPB2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPB2)%></td>
<%
Else
			EstadoServicioESPB2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPB2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>


















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPC1
SQLESPC1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPC1'"

Set ObjRs=ObjConn.Execute(SQLESPC1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - C:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadCCiclo1" value="1" id="idespecialidadc1" onclick='validar(formularioaltapracticaclinica.EspecialidadCCiclo1,1,1,"espc1")' /><label for="idespecialidadc1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadCCiclo1" value="1" id="idespecialidadc1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="75" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPC1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPC1="CERRADO"
			%> 
			<td width="105" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPC1)%></td>
<%
Else
			EstadoServicioESPC1="DISPONIBLE"
			%>
			<td width="96" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPC1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPC2
SQLESPC2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPC2'"

Set ObjRs=ObjConn.Execute(SQLESPC2)
%>

<td width="24" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - C:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="28" height="12" align="center"><input type='checkbox' name="EspecialidadCCiclo2" value="1" id="idespecialidadc2" onclick='validar(formularioaltapracticaclinica.EspecialidadCCiclo2,2,1,"espc2")' /><label for="idespecialidadc2"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadCCiclo2" value="1" id="idespecialidadc2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="99" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPC2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPC2="CERRADO"
			%> 
			<td width="117" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPC2)%></td>
<%
Else
			EstadoServicioESPC2="DISPONIBLE"
			%>
			<td width="104" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPC2)%></td>
<%
End if
%>
</tr>
</table>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->


<br />



<table width="1100" align="center" border="0">
        <tr>
                <td width="935" height="12" align="left" class="textotitulosubcabeceras1">SELECCIÓN DE ESPECIALIDADES EN LOS SERVICIOS MÉDICOS.</td>
        </tr>
</table>


<table width="1100" align="center" border="0"  cellspacing="0">

        <tr>
                <td width="378" height="12" align="center" class="textotitulocampo3">1º CICLO [Primer Cuatrimestre]</td>
                <td width="412" height="12" align="center" class="textotitulocampo3">2º CICLO [Segundo Cuatrimestre]</td>
        </tr>
        
        <tr>
                <td width="378" height="12" align="center" class="textotitulocampo3"></td>
                <td width="412" height="12" align="center" class="textotitulocampo3"></td>
        </tr>

</table>


<br />



<table width="1100" align="center" border="0" cellspacing="0">

        <tr>
                <td width="178" height="12" align="center" class="textotitulocampo3"></td>
          <td width="20" height="12" align="left" class="textotitulocampo3"></td>
          <td width="68" height="12" align="left" class="textotitulocampo3"></td>
          <td width="106" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
      <td width="129" height="12" align="center" class="textotitulocampo3">ESTADO DEL</th>
          <td width="59" height="12" align="center" class="textotitulocampo3"></td>
          <td width="27" height="12" align="left" class="textotitulocampo3"></td>
          <td width="195" height="12" align="left" class="textotitulocampo3"></td>
          <td width="140" height="12" align="center" class="textotitulocampo3">PLAZAS</td>
    <td width="158" height="12" align="center" class="textotitulocampo3">ESTADO DEL </th>    </tr>
        </td>
        
        <tr>
                <td width="178" height="12" align="center" class="textotitulocampo3">ESPECIALIDAD</td>
          <td width="20" height="12" align="left" class="textotitulocampo3"></td>
          <td width="68" height="12" align="left" class="textotitulocampo3"></td>
          <td width="106" height="12" align="center" class="textotitulocampo3">DISPONIBLES</td>
      <td width="129" height="12" align="center" class="textotitulocampo3">SERVICIO</th>
          <td width="59" height="12" align="center" class="textotitulocampo3"></td>
          <td width="27" height="12" align="left" class="textotitulocampo3"></td>
          <td width="195" height="12" align="left" class="textotitulocampo3">ESPECIALIDAD</td>
          <td width="140" height="12" align="center" class="textotitulocampo3">DISPONIBLES</td>
          <td width="158" height="12" align="center" class="textotitulocampo3">SERVICIO</td>
    </tr>

</table>



<table width="1100" align="center" border="0">


<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLHEMA1
SQLHEMA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='HEMA1'"

Set ObjRs=ObjConn.Execute(SQLHEMA1)
%>

<tr>
<td width="169" height="12" align="right" class="textotitulocampo1">Hematología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>

	  <td width="24" height="12" align="center"><input type='checkbox' name="HematologiaCiclo1"  value="1" id="idhematologia1" onclick='validar(formularioaltapracticaclinica.HematologiaCiclo1,1,2,"hem1")' /><label for="idhematologia1"></label></td>
<%
Else
%>

	  <td width="20" height="12" align="center"><input type='checkbox' name="HematologiaCiclo1"  value="1" id="idhematologia1" style="visibility:hidden"/></td>

<%
End If
%>


<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioHEMA1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioHEMA1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioHEMA1)%></td>
<%
Else
			EstadoServicioHEMA1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioHEMA1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLHEMA2
SQLHEMA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='HEMA2'"

Set ObjRs=ObjConn.Execute(SQLHEMA2)
%>

<td width="21" height="12" align="left"></td>
<td width="169" height="12" align="right" class="textotitulocampo1">Hematología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="HematologiaCiclo2" value="1" id="idhematologia2" onclick='validar(formularioaltapracticaclinica.HematologiaCiclo2,2,2,"hem2")' /><label for="idhematologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="HematologiaCiclo2" value="1" id="idhematologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioHEMA2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioHEMA2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioHEMA2)%></td>
<%
Else
			EstadoServicioHEMA2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioHEMA2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEUM1
SQLNEUM1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEUM1'"

Set ObjRs=ObjConn.Execute(SQLNEUM1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Neumología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>

	  <td width="24" height="12" align="center"><input type='checkbox' name="NeumologiaCiclo1" value="1"  id="idneumologia1" onclick='validar(formularioaltapracticaclinica.NeumologiaCiclo1,1,2,"neum1")' /><label for="idneumologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="NeumologiaCiclo1" value="1"  id="idneumologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioNEUM1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioNEUM1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEUM1)%></td>
<%
Else
			EstadoServicioNEUM1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEUM1)%></td>
<%
End if
%>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEUM2
SQLNEUM2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEUM2'"

Set ObjRs=ObjConn.Execute(SQLNEUM2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neumología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

	  <td width="26" height="12" align="center"><input type='checkbox' name="NeumologiaCiclo2" value="1" id="idneumologia2" onclick='validar(formularioaltapracticaclinica.NeumologiaCiclo2,2,2,"neum2")' /><label for="idneumologia2"></label></td>
<%
else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="NeumologiaCiclo2" value="1" id="idneumologia2" style="visibility:hidden"/></td> 

<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioNEUM2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioNEUM2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEUM2)%></td>
<%
Else
			EstadoServicioNEUM2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEUM2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLCGIA1
SQLCGIA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='CGIA1'"

Set ObjRs=ObjConn.Execute(SQLCGIA1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Cardiología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>

	  <td width="24" height="12" align="center"><input type='checkbox' name="CardiologiaCiclo1" value="1" id="idcardiologia1" onclick='validar(formularioaltapracticaclinica.CardiologiaCiclo1,1,2,"card1")' /><label for="idcardiologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="CardiologiaCiclo1" value="1" id="idcardiologia1" style="visibility:hidden"/></td>
<%
End If
%>


<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>


<%
Dim EstadoServicioCGIA1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioCGIA1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioCGIA1)%></td>
<%
Else
			EstadoServicioCGIA1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioCGIA1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLCGIA2
SQLCGIA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='CGIA2'"

Set ObjRs=ObjConn.Execute(SQLCGIA2)
%>


<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Cardiología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

	  <td width="26" height="12" align="center"><input type='checkbox' name="CardiologiaCiclo2" value="1" id="idcardiologia2" onclick='validar(formularioaltapracticaclinica.CardiologiaCiclo2,2,2,"card2")' /><label for="idcardiologia2"></label></td>
<%
Else
%>

	  <td width="22" height="12" align="center"><input type='checkbox' name="CardiologiaCiclo2" value="1" id="idcardiologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioCGIA2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioCGIA2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioCGIA2)%></td>
<%
Else
			EstadoServicioCGIA2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioCGIA2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>























<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLDIGE1
SQLDIGE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='DIGE1'"

Set ObjRs=ObjConn.Execute(SQLDIGE1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Digestivo:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
				
	  <td width="24" height="12" align="center"><input type='checkbox' name="DigestivoCiclo1" value="1" id="iddigestivo1" onclick='validar(formularioaltapracticaclinica.DigestivoCiclo1,1,2,"dige1")' /><label for="iddigestivo1"></label></td>
<%
Else
%>

			<td width="20" height="12" align="center"><input type='checkbox' name="DigestivoCiclo1" value="1" id="iddigestivo1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioDIGE1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioDIGE1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioDIGE1)%></td>
<%
else
			EstadoServicioDIGE1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioDIGE1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLDIGE2
SQLDIGE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='DIGE2'"

Set ObjRs=ObjConn.Execute(SQLDIGE2)
%>


<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Digestivo:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="DigestivoCiclo2" value="1" id="iddigestivo2" onclick='validar(formularioaltapracticaclinica.DigestivoCiclo2,2,2,"dige2")' /><label for="iddigestivo2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="DigestivoCiclo2" value="1" id="iddigestivo2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioDIGE2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioDIGE2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioDIGE2)%></td>
<%
Else
			EstadoServicioDIGE2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioDIGE2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNEFR1
SQLNEFR1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NEFR1'"

Set ObjRs=ObjConn.Execute(SQLNEFR1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Nefrología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="NefrologiaCiclo1" value="1" id="idnefrologia1" onclick='validar(formularioaltapracticaclinica.NefrologiaCiclo1,1,2,"nef1")' /><label for="idnefrologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="NefrologiaCiclo1" value="1" id="idnefrologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioNEFR1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioNEFR1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEFR1)%></td>
<%
Else
			EstadoServicioNEFR1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEFR1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNEFR2
SQLNEFR2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NEFR2'"

Set ObjRs=ObjConn.Execute(SQLNEFR2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Nefrología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="NefrologiaCiclo2" value="1" id="idnefrologia2" onclick='validar(formularioaltapracticaclinica.NefrologiaCiclo2,2,2,"nef2")' /><label for="idnefrologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="NefrologiaCiclo2" value="1" id="idnefrologia2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioNEFR2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioNEFR2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNEFR2)%></td>
<%
Else
			EstadoServicioNEFR2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNEFR2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>













<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMINT1
SQLMINT1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MINT1'"

Set ObjRs=ObjConn.Execute(SQLMINT1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Medicina Interna:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="MedicinaInternaCiclo1" value="1" id="idmedicinainterna1" onclick='validar(formularioaltapracticaclinica.MedicinaInternaCiclo1,1,2,"mint1")' /><label for="idmedicinainterna1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="MedicinaInternaCiclo1" value="1" id="idmedicinainterna1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioMINT1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

		EstadoServicioMINT1="CERRADO"
		%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMINT1)%></td>
<%
Else
		EstadoServicioMINT1="DISPONIBLE"
		%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMINT1)%></td>
<%
End if
%>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMINT2
SQLMINT2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MINT2'"

Set ObjRs=ObjConn.Execute(SQLMINT2)
%>


<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Medicina Interna:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="MedicinaInternaCiclo2" value="1" id="idmedicinainterna2" onclick='validar(formularioaltapracticaclinica.MedicinaInternaCiclo2,2,2,"mint2")' /><label for="idmedicinainterna2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="MedicinaInternaCiclo2" value="1" id="idmedicinainterna2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioMINT2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioMINT2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMINT2)%></td>
<%
Else
			EstadoServicioMINT2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMINT2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLENDO1
SQLENDO1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ENDO1'"

Set ObjRs=ObjConn.Execute(SQLENDO1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Endocrinología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="EndocrinologiaCiclo1" value="1" id="idendocrinologia1" onclick='validar(formularioaltapracticaclinica.EndocrinologiaCiclo1,1,2,"endo1")' /><label for="idendocrinologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EndocrinologiaCiclo1" value="1" id="idendocrinologia1" style="visibility:hidden"/></td>
<%
End If
%>


<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioENDO1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioENDO1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioENDO1)%></td>
<%
Else
			EstadoServicioENDO1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioENDO1)%></td>
<%
End if
%>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLENDO2
SQLENDO2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ENDO2'"

Set ObjRs=ObjConn.Execute(SQLENDO2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Endocrinología:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="EndocrinologiaCiclo2" value="1" id="idendocrinologia2" onclick='validar(formularioaltapracticaclinica.EndocrinologiaCiclo2,2,2,"endo2")' /><label for="idendocrinologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="EndocrinologiaCiclo2" value="1" id="idendocrinologia2" style="visibility:hidden"/></td>
<%
End If
%>


<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioENDO2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioENDO2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioENDO2)%></td>
<%
Else
			EstadoServicioENDO2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioENDO2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLREUM1
SQLREUM1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='REUM1'"

Set ObjRs=ObjConn.Execute(SQLREUM1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Reumatología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="ReumatologiaCiclo1" value="1" id="idreumatologia1" onclick='validar(formularioaltapracticaclinica.ReumatologiaCiclo1,1,2,"reu1")' /><label for="idreumatologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="ReumatologiaCiclo1" value="1" id="idreumatologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioREUM1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioREUM1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioREUM1)%></td>
<%
Else
			EstadoServicioREUM1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioREUM1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>










<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLREUM2
SQLREUM2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='REUM2'"

Set ObjRs=ObjConn.Execute(SQLREUM2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Reumatología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="ReumatologiaCiclo2" value="1" id="idreumatologia2" onclick='validar(formularioaltapracticaclinica.ReumatologiaCiclo2,2,2,"reu2")' /><label for="idreumatologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="ReumatologiaCiclo2" value="1" id="idreumatologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioREUM2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioREUM2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioREUM2)%></td>
<%
Else
			EstadoServicioREUM2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioREUM2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLONCO1
SQLONCO1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ONCO1'"

Set ObjRs=ObjConn.Execute(SQLONCO1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Oncología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="OncologiaCiclo1" value="1" id="idoncologia1" onclick='validar(formularioaltapracticaclinica.OncologiaCiclo1,1,2,"onco1")' /><label for="idoncologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="OncologiaCiclo1" value="1" id="idoncologia1" style="visibility:hidden"/></td>			
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioONCO1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioONCO1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioONCO1)%></td>
<%
Else
			EstadoServicioONCO1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioONCO1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLONCO2
SQLONCO2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ONCO2'"

Set ObjRs=ObjConn.Execute(SQLONCO2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Oncología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="OncologiaCiclo2" value="1" id="idoncologia2" onclick='validar(formularioaltapracticaclinica.OncologiaCiclo2,2,2,"onco2")' /><label for="idoncologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="OncologiaCiclo2" value="1" id="idoncologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioONCO2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioONCO2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioONCO2)%></td>
<%
Else
			EstadoServicioONCO2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioONCO2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLNGIA1
SQLNGIA1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='NGIA1'"

Set ObjRs=ObjConn.Execute(SQLNGIA1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Neurología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="NeurologiaCiclo1" value="1" id="idneurologia1" onclick='validar(formularioaltapracticaclinica.NeurologiaCiclo1,1,2,"neuro1")' /><label for="idneurologia1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="NeurologiaCiclo1" value="1" id="idneurologia1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioNGIA1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioNGIA1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNGIA1)%></td>
<%
Else
			EstadoServicioNGIA1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNGIA1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLNGIA2
SQLNGIA2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='NGIA2'"

Set ObjRs=ObjConn.Execute(SQLNGIA2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Neurología:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="NeurologiaCiclo2" value="1" id="idneurologia2" onclick='validar(formularioaltapracticaclinica.NeurologiaCiclo2,2,2,"neuro2")' /><label for="idneurologia2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="NeurologiaCiclo2" value="1" id="idneurologia2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioNGIA2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioNGIA2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioNGIA2)%></td>
<%
Else
			EstadoServicioNGIA2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioNGIA2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLINFE1
SQLINFE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='INFE1'"

Set ObjRs=ObjConn.Execute(SQLINFE1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Infecciosos:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="InfecciososCiclo1" value="1" id="idinfecciosos1" onclick='validar(formularioaltapracticaclinica.InfecciososCiclo1,1,2,"infe1")' /><label for="idinfecciosos1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="InfecciososCiclo1" value="1" id="idinfecciosos1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioINFE1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioINFE1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioINFE1)%></td>
<%
Else
			EstadoServicioINFE1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioINFE1)%></td>
<%
End if
%>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLINFE2
SQLINFE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='INFE2'"

Set ObjRs=ObjConn.Execute(SQLINFE2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Infecciosos:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
	  <td width="26" height="12" align="center"><input type='checkbox' name="InfecciososCiclo2" value="1" id="idinfecciosos2" onclick='validar(formularioaltapracticaclinica.InfecciososCiclo2,2,2,"infe2")' /><label for="idinfecciosos2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="InfecciososCiclo2" value="1" id="idinfecciosos2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioINFE2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioINFE2="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioINFE2)%></td>
<%
Else
			EstadoServicioINFE2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioINFE2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>












<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLMSIV1
SQLMSIV1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='MSIV1'"

Set ObjRs=ObjConn.Execute(SQLMSIV1)
%>

<tr>
<td height="12" class="textotitulocampo1" align="right">Medicina Intensiva:</td>


<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
	  <td width="24" height="12" align="center"><input type='checkbox' name="MedicinaIntensivaCiclo1" value="1" id="idmedicinaintensiva1" onclick='validar(formularioaltapracticaclinica.MedicinaIntensivaCiclo1,1,2,"minten1")' /><label for="idmedicinaintensiva1"></label></td>
<%
Else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="MedicinaIntensivaCiclo1" value="1" id="idmedicinaintensiva1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioMSIV1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioMSIV1="CERRADO"
			%> 
	  <td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMSIV1)%></td>
<%
else
			EstadoServicioMSIV1="DISPONIBLE"
			%>
	  <td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMSIV1)%></td>
<%
End if
%>




<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>









<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLMSIV2
SQLMSIV2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='MSIV2'"

Set ObjRs=ObjConn.Execute(SQLMSIV2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1" align="right">Medicina Intensiva:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>

	  <td width="26" height="12" align="center"><input type='checkbox' name="MedicinaIntensivaCiclo2" value="1" id="idmedicinaintensiva2" onclick='validar(formularioaltapracticaclinica.MedicinaIntensivaCiclo2,2,2,"minten2")' /><label for="idmedicinaintensiva2"></label></td>
<%
Else
%>
	  <td width="22" height="12" align="center"><input type='checkbox' name="MedicinaIntensivaCiclo2" value="1" id="idmedicinaintensiva2" style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioMSIV2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioMSIV2="CERRADO"
			%> 
	  		<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioMSIV2)%></td>
<%
Else
			EstadoServicioMSIV2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioMSIV2)%></td>
<%
End if
%>
</tr>


<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>





<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->

<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPD1
SQLESPD1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPD1'"

Set ObjRs=ObjConn.Execute(SQLESPD1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - D:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="24" height="12" align="center"><input type='checkbox' name="EspecialidadDCiclo1" value="1" id="idespecialidadd1" onclick='validar(formularioaltapracticaclinica.EspecialidadDCiclo1,1,2,"espd1")' /><label for="idespecialidadd1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadDCiclo1" value="1" id="idespecialidadd1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPD1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPD1="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPD1)%></td>
<%
Else
			EstadoServicioESPD1="DISPONIBLE"
			%>
			<td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPD1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPD2
SQLESPD2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPD2'"

Set ObjRs=ObjConn.Execute(SQLESPD2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - D:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="26" height="12" align="center"><input type='checkbox' name="EspecialidadDCiclo2" value="1" id="idespecialidadd2" onclick='validar(formularioaltapracticaclinica.EspecialidadDCiclo2,2,2,"espd2")' /><label for="idespecialidadd2"></label></td>
<%
Else
%>
			<td width="22" height="12" align="center"><input type='checkbox' name="EspecialidadDCiclo2" value="1" id="idespecialidadd2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPD2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPD2="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPD2)%></td>
<%
Else
			EstadoServicioESPD2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPD2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>














<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPE1
SQLESPE1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPE1'"

Set ObjRs=ObjConn.Execute(SQLESPE1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - E:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="24" height="12" align="center"><input type='checkbox' name="EspecialidadECiclo1" value="1" id="idespecialidade1" onclick='validar(formularioaltapracticaclinica.EspecialidadECiclo1,1,2,"espe1")' /><label for="idespecialidade1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadECiclo1" value="1" id="idespecialidade1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPE1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPE1="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPE1)%></td>
<%
Else
			EstadoServicioESPE1="DISPONIBLE"
			%>
			<td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPE1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPE2
SQLESPE2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPE2'"

Set ObjRs=ObjConn.Execute(SQLESPE2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - E:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="26" height="12" align="center"><input type='checkbox' name="EspecialidadECiclo2" value="1" id="idespecialidade2" onclick='validar(formularioaltapracticaclinica.EspecialidadECiclo2,2,2,"espe2")' /><label for="idespecialidade2"></label></td>
<%
Else
%>
			<td width="22" height="12" align="center"><input type='checkbox' name="EspecialidadECiclo2" value="1" id="idespecialidade2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPE2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPE2="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPE2)%></td>
<%
Else
			EstadoServicioESPE2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPE2)%></td>
<%
End if
%>
</tr>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>


















<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_1", ObjConn,3,3

Dim SQLESPF1
SQLESPF1="SELECT * FROM tbl_ServiciosCiclo_1 WHERE CLAVE_SERVICIO='ESPF1'"

Set ObjRs=ObjConn.Execute(SQLESPF1)
%>

<tr>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - F:</td>

<%
if ObjRs.Fields("PLAZAS_DISPONIBLES_C1")>0 then 
%>
			<td width="24" height="12" align="center"><input type='checkbox' name="EspecialidadFCiclo1" value="1" id="idespecialidadf1" onclick='validar(formularioaltapracticaclinica.EspecialidadFCiclo1,1,2,"espf1")' /><label for="idespecialidadf1"></label></td>
<%
else
%>
			<td width="20" height="12" align="center"><input type='checkbox' name="EspecialidadFCiclo1" value="1" id="idespecialidadf1" style="visibility:hidden"/></td>
<%
End If
%>

<td width="70" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C1"))%></td>

<%
Dim EstadoServicioESPF1

If ObjRs.Fields("PLAZAS_DISPONIBLES_C1")=0 then

			EstadoServicioESPF1="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPF1)%></td>
<%
Else
			EstadoServicioESPF1="DISPONIBLE"
			%>
			<td width="89" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPF1)%></td>
<%
End if
%>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<%
Set ObjConn = Server.CreateObject("ADODB.Connection") 
ObjConn.Open "dsndbgpc" 

Set ObjRs = Server.CreateObject("ADODB.Recordset")
ObjRs.Open "tbl_ServiciosCiclo_2", ObjConn,3,3

Dim SQLESPF2
SQLESPF2="SELECT * FROM tbl_ServiciosCiclo_2 WHERE CLAVE_SERVICIO='ESPF2'"

Set ObjRs=ObjConn.Execute(SQLESPF2)
%>

<td width="21" height="12" align="left"></td>
<td height="12" class="textotitulocampo1espfree" align="right">Especialidad - F:</td>

<%
If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")>0 then 
%>
			<td width="26" height="12" align="center"><input type='checkbox' name="EspecialidadFCiclo2" value="1" id="idespecialidadf2" onclick='validar(formularioaltapracticaclinica.EspecialidadFCiclo2,2,2,"espf2")' /><label for="idespecialidadf2"></label></td>
<%
Else
%>
			<td width="22" height="12" align="center"><input type='checkbox' name="EspecialidadFCiclo2" value="1" id="idespecialidadf2"  style="visibility:hidden"/></td>
<%
End If
%>

<td width="92" height="12" align="center" class="campotexto1"><%Response.Write(ObjRs.Fields("PLAZAS_DISPONIBLES_C2"))%></td>

<%
Dim EstadoServicioESPF2

If ObjRs.Fields("PLAZAS_DISPONIBLES_C2")=0 then

			EstadoServicioESPF2="CERRADO"
			%> 
			<td width="116" height="12" align="center" class="campotextoestado1"><%Response.Write(EstadoServicioESPF2)%></td>
<%
Else
			EstadoServicioESPF2="DISPONIBLE"
			%>
			<td width="112" height="12" align="center" class="campotextoestado2"><%Response.Write(EstadoServicioESPF2)%></td>
<%
End if
%>
</tr>
</table>



<%
ObjRs.Close
Set ObjRs = Nothing
			
ObjConn.Close
Set ObjConn = Nothing 
%>











<!--
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
########################################################################################################################
-->


















<br /><br /><br />


<div id="mensajeaceptacionformulario" style="display: none">
    <table align="center" width="800" border="0">
        <tr>
            <td align="center"><input type="checkbox" value="0" id="idaceptaformulario" name="AceptaFormularioOk" onclick="habilitabotonactualizarficha()"/><label for="idaceptaformulario"></label></td>
            <td align="left">Si esta conforme con todos los datos introducidos y son correctos en el formulario que acaba de rellenar, marca esta casilla y pulse el siguiente botón  <b>'GRABAR ESTE FORMULARIO DEL ALUMNO/A EN LA BASE DE DATOS E IMPRIMIR DOCUMENTO PDF' </b>y generar su informe de asignación de prácticas clínicas PDF para imprimir.</td>
      </tr>
    </table>
</div>   

<br /><br />

<table width="1100" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr>
           <td width="184"><img src="jpg/seleccionfinalizada_170_x_170.jpg" id="idimagefinalizado" style="visibility:hidden" /></td>
      <td width="916" align="left" valign="top"><input type="Submit" name="enviardatos" id="idenviardatos" value="GRABAR ESTE FORMULARIO DEL ALUMNO/A EN LA BASE DE DATOS E IMPRIMIR DOCUMENTO PDF"  style="visibility: hidden" /></td>
    </tr>
</table>


<br/>

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



</form>


 </div>


<%
End if
%>


</body>
</html>