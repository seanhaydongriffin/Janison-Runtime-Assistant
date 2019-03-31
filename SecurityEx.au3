#include-once

dim $h__Advapi32Dll = "advapi32.dll"
dim $h__Kernel32Dll = "kernel32.dll"

; #FUNCTION# ====================================================================================================================================
; Name...........: _SetPrivilege
; Description ...: Enables or disables special privileges as required by some DllCalls
; Syntax.........: _SetPrivilege($avPrivilege)
; Parameters ....: $avPrivilege - An array of privileges and respective attributes
;                                 $SE_PRIVILEGE_ENABLED - The function enables the privilege
;                                 $SE_PRIVILEGE_REMOVED - The privilege is removed from the list of privileges in the token
;                                 0 - The function disables the privilege
; Requirement(s).: None
; Return values .: Success - An array of modified privileges and their respective previous attribute state
;                  Failure - An empty array
;                            Sets @error
; Author ........: engine
; Modified.......: FredAI
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......;
; ===============================================================================================================================================

Func _SetPrivilege($avPrivilege)
	Local $iDim = UBound($avPrivilege, 0), $avPrevState[1][2]
	If Not ( $iDim <= 2 And UBound($avPrivilege, $iDim) = 2 ) Then Return SetError(1300, 0, $avPrevState)
	If $iDim = 1 Then
		Local $avTemp[1][2]
		$avTemp[0][0] = $avPrivilege[0]
		$avTemp[0][1] = $avPrivilege[1]
		$avPrivilege = $avTemp
		$avTemp = 0
	EndIf
	Local $k, $tagTP = "dword", $iTokens = UBound($avPrivilege, 1)
	Do
		$k += 1
		$tagTP &= ";dword;long;dword"
	Until $k = $iTokens
	Local $tCurrState, $tPrevState, $pPrevState, $tLUID, $ahGCP, $avOPT, $aiGLE
	$tCurrState = DLLStructCreate($tagTP)
	$tPrevState = DllStructCreate($tagTP)
	$pPrevState = DllStructGetPtr($tPrevState)
	$tLUID = DllStructCreate("dword;long")
	DLLStructSetData($tCurrState, 1, $iTokens)
	For $i = 0 To $iTokens - 1
		DllCall($h__Advapi32Dll, "int", "LookupPrivilegeValue", _
			"str", "", _
			"str", $avPrivilege[$i][0], _
			"ptr", DllStructGetPtr($tLUID) )
		DLLStructSetData( $tCurrState, 3 * $i + 2, DllStructGetData($tLUID, 1) )
		DLLStructSetData( $tCurrState, 3 * $i + 3, DllStructGetData($tLUID, 2) )
		DLLStructSetData( $tCurrState, 3 * $i + 4, $avPrivilege[$i][1] )
	Next
	$ahGCP = DllCall($h__Kernel32Dll, "hwnd", "GetCurrentProcess")
	$avOPT = DllCall($h__Advapi32Dll, "int", "OpenProcessToken", _
		"hwnd", $ahGCP[0], _
		"dword", BitOR(0x00000020, 0x00000008), _
		"hwnd*", 0 )
	DllCall( $h__Advapi32Dll, "int", "AdjustTokenPrivileges", _
		"hwnd", $avOPT[3], _
		"int", False, _
		"ptr", DllStructGetPtr($tCurrState), _
		"dword", DllStructGetSize($tCurrState), _
		"ptr", $pPrevState, _
		"dword*", 0 )
	$aiGLE = DllCall($h__Kernel32Dll, "dword", "GetLastError")
	DllCall($h__Kernel32Dll, "int", "CloseHandle", "hwnd", $avOPT[3])
	Local $iCount = DllStructGetData($tPrevState, 1)
	If $iCount > 0 Then
		Local $pLUID, $avLPN, $tName, $avPrevState[$iCount][2]
		For $i = 0 To $iCount - 1
			$pLUID = $pPrevState + 12 * $i + 4
			$avLPN = DllCall($h__Advapi32Dll, "int", "LookupPrivilegeName", _
				"str", "", _
				"ptr", $pLUID, _
				"ptr", 0, _
				"dword*", 0 )
			$tName = DllStructCreate("char[" & $avLPN[4] & "]")
			DllCall($h__Advapi32Dll, "int", "LookupPrivilegeName", _
				"str", "", _
				"ptr", $pLUID, _
				"ptr", DllStructGetPtr($tName), _
				"dword*", DllStructGetSize($tName) )
			$avPrevState[$i][0] = DllStructGetData($tName, 1)
			$avPrevState[$i][1] = DllStructGetData($tPrevState, 3 * $i + 4)
		Next
	EndIf
	Return SetError($aiGLE[0], 0, $avPrevState)
EndFunc ;==> _SetPrivilege
