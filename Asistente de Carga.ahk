; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1

VersionDelPrograma := "0.5 Alpha"
UltimoUsado := "Ninguno"

;Lee la configuracion del archivo
IniRead, ColumnaNumeroODT, Configuracion.ini, Columnas De Extraccion, ColumnaNumeroODT
IniRead, ColumnaNumElemento, Configuracion.ini, Columnas De Extraccion, ColumnaNumElemento
IniRead, ColumnaPInversion, Configuracion.ini, Columnas De Extraccion, ColumnaPInversion
IniRead, ColumnaFecha, Configuracion.ini, Columnas De Extraccion, ColumnaFecha
IniRead, ColumnaNomCalle1, Configuracion.ini, Columnas De Extraccion, ColumnaNomCalle1
IniRead, ColumnaNumCalleOY, Configuracion.ini, Columnas De Extraccion, ColumnaNumCalleOY
IniRead, ColumnaNomCalle2, Configuracion.ini, Columnas De Extraccion, ColumnaNomCalle2
IniRead, ColumnaLongitud, Configuracion.ini, Columnas De Extraccion, ColumnaLongitud
IniRead, ColumnaMaterialPosterior, Configuracion.ini, Columnas De Extraccion, ColumnaMaterialPosterior
IniRead, ColumnaDiametroPosterior, Configuracion.ini, Columnas De Extraccion, ColumnaDiametroPosterior
IniRead, ColumnaSubcuenca, Configuracion.ini, Columnas De Extraccion, ColumnaSubcuenca
IniRead, ColumnaMetrosRastreados, Configuracion.ini, Columnas De Extraccion, ColumnaMetrosRastreados


;Menu Principal "Inicio"
	Inicio:
		Gui, Inicio:Default
		Gui,  Add, Button, w500 h50 gCargaMarcoyTapa, Carga de Renovacion Marco y Tapa
		Gui,  Add, Button, w500 h50 gCargaCañeriasAgua, Carga de Renovacion Cañerias de Agua
		Gui,  Add, Button, w500 h50 gCargaCañeriasCloaca, Carga de Renovacion Cañerias de Cloaca
		Gui,  Add, Button, w500 h50 gCargaArtefactoValvula, Carga de Renovacion Artefacto Valvula
		Gui,  Add, Button, w500 h50 gCargaArtefactoHidrante, Carga de Renovacion Artefacto Hidrante
		Gui,  Add, Button, W500 h50, Configuracion
		Gui,  Show, , Menu Principal - Programa de Automatización de carga en Gred %VersionDelPrograma%
	return
;Menu Principal "Inicio"

	;Inicio SubRutina Carga Marco y Tapa
		CargaMarcoyTapa:
			if (NOT((UltimoUsado = "CargaMarcoyTapa") OR (UltimoUsado = "Ninguno"))){
			MsgBox,,,"Recuerde configurar las Columnas"
			}
			UltimoUsado := "CargaMarcoyTapa"
			Gui, Inicio: Hide
			Gui, MarcoyTapa:Default
			CustomColor = FFFFFF  ; Can be any RGB color (it will be made transparent below).
			Gui, +LastFound +AlwaysOnTop +Caption -ToolWindow +Border +SysMenu +MinimizeBox +Resize -MaximizeBox ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
			Gui, Color, %CustomColor%
			Gui, Font, s10  ; Set a large font size (32-point). 
			Gui, Add, Text, w300 h180 vMyText cBlack,  ; XX & YY serve to auto-size the window.
			Gui, Add, Edit, w100 
			Gui, Add, UpDown, vMyUpDown gExtraccion Range1-500, 1
			Gui, Font, s6  ; Set a large font size (32-point). 
			Gui, Add, Text,cBlack, F1 Para Ver la Ayuda
			;Make all pixels of this color transparent and make the text itself translucent (150)
			;WinSet, TransColor, %CustomColor% 250
			;SetTimer, UpdateOSD, 500
			Gosub, UpdateOSD  ; Make the first update immediate rather than waiting for the timer.
			Gui, Show, x10 y10 NoActivate AutoSize ; NoActivate avoids deactivating the currently active window.
		return
		
		UpdateOSD:
			GuiControl, , MyText,Calle:`n%CalleParaPegar%`nNumeroElemento:`n%NumElemento%`nFecha:`n%Fecha%`nProyInversion:`n%PInversion%`nNumeroODT:`n%NumeroODT%`nLongitud:`n%Longitud%`nDiametro:`n%Diametro%`nMaterial:`n%Material%
		return

		MarcoyTapaGuiClose:
			Gui, Inicio: Show
			Gui, Destroy
		Return
	;Fin SubRutina Carga Marco y Tapa
	
	;Inicio SubRutina Carga de Renovacion Cañerias de Agua
		CargaCañeriasAgua:
			if (NOT((UltimoUsado = "CargaCañeriasAgua") OR (UltimoUsado = "Ninguno"))){
			MsgBox,,,"Recuerde configurar las Columnas"
			}
			UltimoUsado := "CargaCañeriasAgua"
			Gui, Inicio: Hide
			Gui, CargaCañeriasAgua:Default
			CustomColor = FFFFFF  ; Can be any RGB color (it will be made transparent below).
			Gui, +LastFound +AlwaysOnTop +Caption -ToolWindow +Border +SysMenu +MinimizeBox +Resize -MaximizeBox ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
			Gui, Color, %CustomColor%
			Gui, Font, s10  ; Set a large font size (32-point). 
			Gui, Add, Text, w300 h180 vMyText cBlack,  ; XX & YY serve to auto-size the window.
			Gui, Add, Edit, w100 
			Gui, Add, UpDown, vMyUpDown gExtraccion Range1-500, 1
			Gui, Font, s6  ; Set a large font size (32-point). 
			Gui, Add, Text,cBlack, F1 Para Ver la Ayuda
			;Make all pixels of this color transparent and make the text itself translucent (150)
			;WinSet, TransColor, %CustomColor% 250
			;SetTimer, UpdateOSD, 500
			Gosub, UpdateOSDAgua  ; Make the first update immediate rather than waiting for the timer.
			Gui, Show, x10 y10 NoActivate AutoSize ; NoActivate avoids deactivating the currently active window.
		return
		
		UpdateOSDAgua:
			GuiControl, , MyText,Calle:`n%CalleParaPegar%`nNumeroElemento:`n%NumElemento%`nFecha:`n%Fecha%`nProyInversion:`n%PInversion%`nNumeroODT:`n%NumeroODT%`nLongitud:`n%Longitud%`nDiametro:`n%Diametro%`nMaterial:`n%Material%
			Gui, Submit, NoHide
		return

		CargaCañeriasAguaGuiClose:
			Gui, Inicio: Show
			Gui, Destroy
		Return
	;Fin SubRutina Carga de Renovacion Cañerias de Agua
	
	;Menu de Configuracion
		InicioButtonConfiguracion:
			Gui, Configuracion:Default
			Gui,  Add, Text, 		 						w270 	h20 , Columna Numero ODT
			Gui,  Add, Edit,	vColumnaNumeroODT 	 		w60 	h20 , %ColumnaNumeroODT%
			Gui,  Add, Text, 								w270 	h20 , Columna Numero de Elemento
			Gui,  Add, Edit,	vColumnaNumElemento			w60 	h20 , %ColumnaNumElemento%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Proyecto de Inversion
			Gui,  Add, Edit,	vColumnaPInversion			w60 	h20 , %ColumnaPInversion%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Fecha
			Gui,  Add, Edit,	vColumnaFecha 	 			w60 	h20 , %ColumnaFecha%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Nombre de Calle 1
			Gui,  Add, Edit,	vColumnaNomCalle1			w60 	h20 , %ColumnaNomCalle1%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Nombre de Entrecalle o Numero
			Gui,  Add, Edit,	vColumnaNumCalleOY			w60 	h20 , %ColumnaNumCalleOY%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Nombre de Calle 2
			Gui,  Add, Edit,	vColumnaNomCalle2	 		w60 	h20 , %ColumnaNomCalle2%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Longitud Informada
			Gui,  Add, Edit,	vColumnaLongitud	 		w60 	h20 , %ColumnaLongitud%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Material Posterior
			Gui,  Add, Edit,	vColumnaMaterialPosterior	w60 	h20 , %ColumnaMaterialPosterior%
			Gui,  Add, Text, 		 						w270 	h20 , Columna Diametro Posterior
			Gui,  Add, Edit,	vColumnaDiametroPosterior	w60 	h20 , %ColumnaDiametroPosterior%
			Gui,  Add, Button, w100 h50 Default, Guardar
			Gui,  Add, Button, x+50 w100 h50, Cancelar
			Gui,  Show, ,Configuracion de Opciones
		Return
		
		ConfiguracionButtonGuardar:
			Gui, Submit
			IniWrite, %ColumnaNumeroODT%, Configuracion.ini, Columnas De Extraccion, ColumnaNumeroODT
			IniWrite, %ColumnaNumElemento%, Configuracion.ini, Columnas De Extraccion, ColumnaNumElemento
			IniWrite, %ColumnaPInversion%, Configuracion.ini, Columnas De Extraccion, ColumnaPInversion
			IniWrite, %ColumnaFecha%, Configuracion.ini, Columnas De Extraccion, ColumnaFecha
			IniWrite, %ColumnaNomCalle1%, Configuracion.ini, Columnas De Extraccion, ColumnaNomCalle1
			IniWrite, %ColumnaNumCalleOY%, Configuracion.ini, Columnas De Extraccion, ColumnaNumCalleOY
			IniWrite, %ColumnaNomCalle2%, Configuracion.ini, Columnas De Extraccion, ColumnaNomCalle2
			IniWrite, %ColumnaLongitud%, Configuracion.ini, Columnas De Extraccion, ColumnaLongitud
			IniWrite, %ColumnaMaterialPosterior%, Configuracion.ini, Columnas De Extraccion, ColumnaMaterialPosterior
			IniWrite, %ColumnaDiametroPosterior%, Configuracion.ini, Columnas De Extraccion, ColumnaDiametroPosterior
			Gui, Destroy
		return
		
		ConfiguracionGuiClose:
			Gui, Destroy
		Return

		ConfiguracionButtonCancelar:
			Gui, Destroy
		Return
	;Fin Menu de Configuracion


CargaCañeriasCloaca:
return

CargaArtefactoValvula:
return

CargaArtefactoHidrante:
return





F1::
TextoDeAyuda =
(
F1 => Ayuda del Programa.

F2 => Leer Linea desde Excel/Actualiza Interface.
F3 => Escribir nombre de Calle.
F4 => Escribir altura.
F5 => Escribir entre-calle.
F6 => Seleccionar la herramienta de selección de boca calle.
F7 => Autocompletado de RedLine

F12 => Salir del Programa
)

MsgBox, 0, Ayuda para el Usuario, %TextoDeAyuda%
return

F2::
Extraccion:
PlanillaActiva := ComObjActive("Excel.Application") ;creates a handle to your currently active excel sheet
FilaSeleccionada:= MyUpDown ;Es el control para subir y bajar
NumeroODT := RegExReplace(PlanillaActiva.Range(ColumnaNumeroODT . FilaSeleccionada).Value, ".000000", "") ;Borra los .000000
NumElemento := PlanillaActiva.Range(ColumnaNumElemento . FilaSeleccionada).Value
PInversion := RegExReplace(PlanillaActiva.Range(ColumnaPInversion . FilaSeleccionada).Value, ".000000", "") ;Borra los .000000
Fecha := PlanillaActiva.Range(ColumnaFecha . FilaSeleccionada).Value
NomCalle1 := PlanillaActiva.Range(ColumnaNomCalle1 . FilaSeleccionada).Value
NumCalleOY := RegExReplace(PlanillaActiva.Range(ColumnaNumCalleOY . FilaSeleccionada).Value, ".000000", "")  ;Borra los .000000
NomCalle2 := PlanillaActiva.Range(ColumnaNomCalle2 . FilaSeleccionada).Value
Diametro :=RegExReplace(PlanillaActiva.Range(ColumnaDiametroPosterior . FilaSeleccionada).Value, ".000000", "")  ;Borra los .000000
Longitud :=RegExReplace(PlanillaActiva.Range(ColumnaLongitud . FilaSeleccionada).Value, ".000000", "")  ;Borra los .000000
Material :=PlanillaActiva.Range(ColumnaMaterialPosterior . FilaSeleccionada).Value

If NumCalleOY = Y
{
    NumCalle := ""
    CalleParaPegar := NomCalle1 " y " NomCalle2
}
Else
{
    NumCalle := NumCalleOY
    CalleParaPegar := NomCalle1 " " NumCalle
}
Gosub, UpdateOSD
Return

F3::
NomCalle1:
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 50
SendEvent, %NomCalle1%
SetKeyDelay, %CurrentKeyDelay%
Return

F4::
Altura:
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 50
SendEvent, %NumCalle%
SetKeyDelay, %CurrentKeyDelay%
Return

F5::
EntreCalle:
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 50
SendEvent, %NomCalle2%
SetKeyDelay, %CurrentKeyDelay%
Return

F6::
SelecBocaReg:
Send, {Alt Down}{r}{Alt Up}
Sleep, 50
Send, {b}
Sleep, 200
Return

F7::
Completado:
Sleep, 333
Click, 272, 165 Left, 1
Sleep, 200
Click, 272, 194 Left, 1
Sleep, 200
Click, 273, 190 Left, 1
Sleep, 200
Click, 273, 258 Left, 1
Sleep, 200
Click, 153, 216 Left, 1
Sleep, 200
Send, Renov. M y T
Sleep, 200
Click, 170, 241 Left, 1
Sleep, 200
Send, %NumeroODT%
Sleep, 100
Click, 267, 292 Left, 1
Sleep, 200
Send, {c}{Enter}
Click, 255, 343 Left, 1
Sleep, 500
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 10
SendEvent, 
(LTrim
Calle: %CalleParaPegar%
P. Inv.: %PInversion%
Fecha: %Fecha%
)
SetKeyDelay, %CurrentKeyDelay%
Click, 273, 561 Left, 1
Sleep, 100
Click, 76, 294 Left, 1
Sleep, 100
Click, 1100, 540 Left, 1
Sleep, 10
Return

F12::ExitApp

#a::
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
SendEvent, %NumeroODT%
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 200
Send, {Tab}
SendRaw, RENOV. CL
Return

#s::
Send, Calle: %CalleParaPegar% – Fecha: %Fecha% – Proyecto de Inversión: %PInversion% – Longitud Renovación: %Longitud% mts – Material: %Material% – Diametro: %Diametro%
Return


#q::
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
SendEvent, %NumeroODT%
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
SendRaw, RENOV. AG
Return

#w::
Send, Calle: %CalleParaPegar% – Fecha: %Fecha% – Proyecto de Inversión: %PInversion% – Longitud Renovación: %Longitud% mts – Material: %Material% – Diametro: %Diametro%
Return

F8::
NumArtefacto:
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 50
SendEvent, %NumElemento%
SetKeyDelay, %CurrentKeyDelay%
Return

F9::
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
SendEvent, %NumeroODT%
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Down}
Sleep, 400
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
Sleep, 200
Send, {Tab}
SendRaw, RENOV. ARTEFACTO
Return

F10::
Send, Calle: %CalleParaPegar% – Fecha: %Fecha% – Proyecto de Inversión: %PInversion% – Diametro: %Diametro%
Return

#r::
WinActivate, GisAysa - Internet Explorer ahk_class IEFrame
Sleep, 50
Send, {Tab}{Tab}{Tab}{i}{n}{t}{e}{r}
Sleep, 1500
Send,{Tab}{r}{a}{s}{t}{r}{e}{o}{Space}{s}{i}{s}
Sleep, 1000
Send,{Tab}{Tab}{c}{l}{o}{a}
Sleep, 500
Send,{Tab}{Tab}{RShift Down}{r}{RShift Up}{a}{s}{t}{r}{e}{o}{Space}
Return

#t::
WinActivate, GisAysa - Internet Explorer ahk_class IEFrame
Sleep, 50
Send, {RShift Down}{f}{RShift Up}{e}{c}{h}{a}{:}{0}{2}{/}{/}{2}{0}{2}{2}{Space}{Space}{RShift Down}{m}{RShift Up}{e}{t}{r}{o}{s}{:}
Return