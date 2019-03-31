#RequireAdmin
;#include <Array.au3>


while True

	Local $arr = ProcessList("chrome.exe")
	;_ArrayDisplay($arr)

	for $i = 1 to $arr[0][0]

		Local priority = _ProcessGetPriority($arr[$i][1])

		if priority > 0 Then

			ProcessSetPriority($arr[$i][1], 0)
		EndIf
	Next

	Sleep(10000)

WEnd
