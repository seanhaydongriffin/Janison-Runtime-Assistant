#AutoIt3Wrapper_Run_Au3Check=n
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>


if $CmdLine[0] = 0 Then

	Local $myedit, $msg

	$gui = GUICreate("libreoffice_launch_csv", 640, 480)

	$help = 	'This tool is designed to launch CSV files as text into LibreOffice Calc.' & @CRLF & @CRLF & _
				'Complete the following steps to setup the tool:' & @CRLF & @CRLF & _
				'1. Copy the two sub-routines below into the "Standard.Module1" of LibreOffice Basic.' & @CRLF & @CRLF & _
				'2. In LibreOffice, Tools -> Customize -> Events, set "Save In" to "LibreOffice", "Document is going to be closed" to macro "close_csv".' & @CRLF & @CRLF & _
				'3. In LibreOffice, Tools -> Customize -> Keyboard, set "Ctrl+S" to macro "save_as_csv".' & @CRLF & @CRLF & _
				'4. In LibreOffice, Tools -> Customize -> Toolbars, deselect "Save" and add the macro "save_as_csv" with the "Save" icon.' & @CRLF & @CRLF & _
				'5. In Windows Explorer, right-click on a CSV file and select "Open With -> Choose Program ...".  Browse to select this tool, then check "Always use the selected program to open this kind of file" and click "OK".' & @CRLF & @CRLF & _
				'' & @CRLF & _
				'sub close_csv' & @CRLF & _
				'	Dim iBox as Integer' & @CRLF & _
				'	iBox = MB_YESNO' & @CRLF & _
				'	If MsgBox ("Do you want to save?", iBox) = IDYES Then' & @CRLF & _
				'	  save_as_csv' & @CRLF & _
				'	End IF' & @CRLF & _
				'	ThisComponent.setModified(False)' & @CRLF & _
				'end sub' & @CRLF & _
				'' & @CRLF & _
				'sub save_as_csv' & @CRLF & _
				'	dim document   as object' & @CRLF & _
				'	dim dispatcher as object' & @CRLF & _
				'	document   = ThisComponent.CurrentController.Frame' & @CRLF & _
				'	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")' & @CRLF & _
				'	dim args1(2) as new com.sun.star.beans.PropertyValue' & @CRLF & _
				'	args1(0).Name = "URL"' & @CRLF & _
				'	args1(0).Value = thiscomponent.getURL()' & @CRLF & _
				'	args1(1).Name = "FilterName"' & @CRLF & _
				'	args1(1).Value = "Text - txt - csv (StarCalc)"' & @CRLF & _
				'	args1(2).Name = "FilterOptions"' & @CRLF & _
				'	args1(2).Value = "44,34,ANSI,1,,0,true,true,true"' & @CRLF & _
				'	dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())' & @CRLF & _
				'end sub'

	$myedit = GUICtrlCreateEdit($help, 10, 10, 620, 420, $ES_AUTOVSCROLL + $WS_VSCROLL)
	$close = GUICtrlCreateButton("Close",5,440,80)
	GUICtrlSetState($close, $GUI_FOCUS)

	GUISetState(@SW_SHOW)

	; Loop until the user exits.
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $close
				ExitLoop

		EndSwitch
	WEnd
	GUIDelete()
	Exit
EndIf

local $fn = $CmdLine[1], $OpenPar[2], $oDesktop, $f_num, $col_str = ""

For $col_num = 1 to 1000

	if $col_num > 1 Then

		$col_str = $col_str & "/"
	EndIf

	$col_str = $col_str & $col_num & "/2"

Next

$filteroptions = "44,34,0,1," & $col_str

$OpenPar[0] = setProp("Filtername", "Text - txt - csv (StarCalc)")
$OpenPar[1] = setProp("FilterOptions", $filteroptions)

$oSM = Objcreate("com.sun.star.ServiceManager")
$oDesktop = $oSM.createInstance("com.sun.star.frame.Desktop")   ; Create a desktop object:
$cURL = Convert2URL($fn)
$oCurCom = $oDesktop.loadComponentFromURL( $cURL, "_blank", 0, $OpenPar)


Func setProp($cName, $uValue)

	$oSM = Objcreate("com.sun.star.ServiceManager")
	$oPropertyValue = $oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	$oPropertyValue.Name = $cName
	$oPropertyValue.Value = $uValue
	$setOOoProp = $oPropertyValue
	Return $setOOoProp
EndFunc

Func Convert2URL($fname)

    $fname = StringReplace($fname, ":", "|") 	; двухиточие Ц на |
    $fname = StringReplace($fname, " ", "%20")  ; пробел Ц на %20
    $fname = "file:///" & StringReplace($fname, "\", "/")
    Return $fname
EndFunc ;=== Convert2URL


