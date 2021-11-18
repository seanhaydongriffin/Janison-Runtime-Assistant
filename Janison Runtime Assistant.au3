#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <BlockInputEx.au3>
#include <Array.au3>
#include <WinAPI.au3>
#include <WinAPIProc.au3>
#include "_SysTray.au3"
#include <GuiEdit.au3>
#include <GuiToolbar.au3>
#include "SecurityEx.au3"
#include "Services.au3"
#include <Process.au3>
#include <File.au3>
#include "Toast.au3"

Global $app_name = "Janison Runtime Assistant"

Global $find_text = ""
Global $find_assigned_to_and_field_text = "onl_c32_s0001,service id"
Local $hotkeyset = False
Local $hotkey2set = False
Local $text_set = False
Local $current_script_window_title = ""
Local $curr_mouse_pos = MouseGetPos()
Local $old_mouse_pos = $curr_mouse_pos
local $mouse_move_pixels = 1
_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)

; Update Windows Hosts file and Credentials Manager with standard credentials to test automation machines

_Toast_Show(0, $app_name, "Updating hosts ...", -30, False, True)

Local $hosts_2d_arr[0][2]
Local $hosts_str = FileRead("C:\Windows\System32\drivers\etc\hosts")

; coffsauto mappings

for $ip_last_number = 1 to 20

	_ArrayAdd($hosts_2d_arr, "10.111.80." & $ip_last_number & "|coffsauto" & $ip_last_number & "highsierra")
;	_ArrayAdd($hosts_2d_arr, "10.111.80." & $ip_last_number & "|coffsauto" & $ip_last_number & "highsierra.janison.com.au")
	_ArrayAdd($hosts_2d_arr, "10.111.81." & $ip_last_number & "|coffsauto" & $ip_last_number)
;	_ArrayAdd($hosts_2d_arr, "10.111.81." & $ip_last_number & "|coffsauto" & $ip_last_number & ".janison.com.au")
Next

; AWS VM mappings

_ArrayAdd($hosts_2d_arr, "172.31.27.218|awsauto1")
_ArrayAdd($hosts_2d_arr, "18.225.36.109|awsauto1public")
_ArrayAdd($hosts_2d_arr, "172.31.15.158|awsauto2")
_ArrayAdd($hosts_2d_arr, "3.137.213.190|awsauto2public")
_ArrayAdd($hosts_2d_arr, "172.31.16.70|awsauto100")
_ArrayAdd($hosts_2d_arr, "3.17.142.144|awsauto100public")
_ArrayAdd($hosts_2d_arr, "172.31.6.50|awsauto200")
_ArrayAdd($hosts_2d_arr, "3.137.186.162|awsauto200public")

; Azure VM mappings

_ArrayAdd($hosts_2d_arr, "10.1.0.5|azureauto1")
_ArrayAdd($hosts_2d_arr, "10.1.0.7|azureauto3")
_ArrayAdd($hosts_2d_arr, "10.1.0.4|azureauto4")
_ArrayAdd($hosts_2d_arr, "10.1.0.6|azureauto5")
_ArrayAdd($hosts_2d_arr, "10.1.0.10|azureauto8")
_ArrayAdd($hosts_2d_arr, "10.1.0.8|azureauto10")
_ArrayAdd($hosts_2d_arr, "10.1.0.9|azureauto11")
_ArrayAdd($hosts_2d_arr, "10.1.0.12|azureauto13")
_ArrayAdd($hosts_2d_arr, "10.1.0.13|azureauto14")
_ArrayAdd($hosts_2d_arr, "10.1.0.14|azureauto15")
_ArrayAdd($hosts_2d_arr, "10.1.0.16|azureauto17")
_ArrayAdd($hosts_2d_arr, "10.1.0.18|azureauto19")
_ArrayAdd($hosts_2d_arr, "10.1.0.19|azureauto20")
_ArrayAdd($hosts_2d_arr, "10.1.0.15|azureauto22")
_ArrayAdd($hosts_2d_arr, "10.1.0.17|azureauto23")
_ArrayAdd($hosts_2d_arr, "10.1.0.20|azureauto24")
_ArrayAdd($hosts_2d_arr, "10.1.0.22|azureauto26")
_ArrayAdd($hosts_2d_arr, "10.1.0.23|azureauto27")
_ArrayAdd($hosts_2d_arr, "10.1.0.24|azureauto28")
_ArrayAdd($hosts_2d_arr, "10.1.0.25|azureauto29")
_ArrayAdd($hosts_2d_arr, "10.1.0.27|azureauto31")
_ArrayAdd($hosts_2d_arr, "10.1.0.29|azureauto33")
_ArrayAdd($hosts_2d_arr, "10.1.0.32|azureauto36")
_ArrayAdd($hosts_2d_arr, "10.1.0.33|azureauto37")
_ArrayAdd($hosts_2d_arr, "10.1.0.34|azureauto38")
_ArrayAdd($hosts_2d_arr, "10.1.0.35|azureauto39")
_ArrayAdd($hosts_2d_arr, "10.1.0.37|azureauto41")
_ArrayAdd($hosts_2d_arr, "10.1.0.40|azureauto44")
_ArrayAdd($hosts_2d_arr, "10.1.0.41|azureauto45")
_ArrayAdd($hosts_2d_arr, "10.1.0.42|azureauto46")
_ArrayAdd($hosts_2d_arr, "10.1.0.43|azureauto47")
_ArrayAdd($hosts_2d_arr, "10.1.0.45|azureauto49")
_ArrayAdd($hosts_2d_arr, "10.1.0.21|azureauto100")
_ArrayAdd($hosts_2d_arr, "10.1.0.30|azureauto103")
_ArrayAdd($hosts_2d_arr, "10.1.0.51|azureauto104")
_ArrayAdd($hosts_2d_arr, "10.1.0.53|azureauto106")
_ArrayAdd($hosts_2d_arr, "10.1.0.54|azureauto107")
_ArrayAdd($hosts_2d_arr, "10.1.0.28|azureauto110")
_ArrayAdd($hosts_2d_arr, "10.1.0.31|azureauto111")
_ArrayAdd($hosts_2d_arr, "10.1.0.49|azureauto113")
_ArrayAdd($hosts_2d_arr, "10.1.0.36|azureauto201")
_ArrayAdd($hosts_2d_arr, "10.1.0.38|azureauto202")
_ArrayAdd($hosts_2d_arr, "10.1.0.46|azureauto205")
_ArrayAdd($hosts_2d_arr, "10.1.0.47|azureauto206")
_ArrayAdd($hosts_2d_arr, "10.1.0.48|azureauto207")
_ArrayAdd($hosts_2d_arr, "10.1.0.50|azureauto209")
_ArrayAdd($hosts_2d_arr, "10.1.0.26|azureauto210")
_ArrayAdd($hosts_2d_arr, "10.1.0.39|azureauto213")

; update / add these to the hosts file

for $i = 0 to (UBound($hosts_2d_arr) - 1)

	Local $host_found = StringRegExp($hosts_str, "(?m)^" & $hosts_2d_arr[$i][0] & "[ \t]", 0)

	if $host_found = 1 Then

		$hosts_str = StringRegExpReplace($hosts_str, "(?m)^" & $hosts_2d_arr[$i][0] & "[ \t].*$", $hosts_2d_arr[$i][0] & " " & $hosts_2d_arr[$i][1])
	Else

		$hosts_str = $hosts_str & @CRLF & $hosts_2d_arr[$i][0] & " " & $hosts_2d_arr[$i][1]
	EndIf
Next

FileDelete("C:\Windows\System32\drivers\etc\hosts")
FileWrite("C:\Windows\System32\drivers\etc\hosts", $hosts_str)

; update VM credentials in Windows

_Toast_Show(0, $app_name, "Updating logins ...", -30, False, True)

for $i = 0 to (UBound($hosts_2d_arr) - 1)

	RunWait(@ComSpec & " /c cmdkey /delete:" & $hosts_2d_arr[$i][0], "", @SW_HIDE)
	RunWait(@ComSpec & " /c cmdkey /add:" & $hosts_2d_arr[$i][0] & " /user:localhost\auto /pass:janison", "", @SW_HIDE)
Next

; Check if this is currently a RDP session

_Toast_Show(0, $app_name, "RDP session check ...", -30, False, True)

Local $rdp_session = False

Local $iPID = Run(@ScriptDir & "\query.exe session 2", "", @SW_HIDE, $STDOUT_CHILD)
ProcessWaitClose($iPID)
Local $sOutput = StdoutRead($iPID)

if StringLen($sOutput) > 0 Then

	$rdp_session = True
EndIf

; Make drive mapping visible to all applications

_Toast_Show(0, $app_name, "Fix drive mappings ...", -30, False, True)

RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLinkedConnections", "REG_DWORD", 1)

; Allow remote execution access to the computer

_Toast_Show(0, $app_name, "Allow remote execution ...", -30, False, True)

RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "LocalAccountTokenFilterPolicy", "REG_DWORD", 1)

; Disable Open File Security Warnings

_Toast_Show(0, $app_name, "Disable security warnings ...", -30, False, True)

RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments", "SaveZoneInformation", "REG_DWORD", 1)
RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Associations", "LowRiskFileTypes", "REG_SZ", ".avi;.bat;.com;.cmd;.exe;.htm;.html;.lnk;.mpg;.mpeg;.mov;.mp3;.msi;.m3u;.rar;.reg;.txt;.vbs;.wav;.zip;")
RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Attachments", "SaveZoneInformation")
RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Associations", "LowRiskFileTypes")

_Toast_Hide()

Opt("WinTitleMatchMode", 3)

dim $hand, $pos
dim $playback_displayed = False
dim $timer = 500000
dim $timer2 = 500000
const $loop_delay = 250 ;1000

while True

	; count the number of seconds
	$timer = $timer + $loop_delay
	$timer2 = $timer2 + $loop_delay

	; every 20 seconds check the following
	if $timer2 > (20 * 1000) Then

		$timer2 = 0



		$proc="chromedriver.exe"
		$oWMI=ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		$oProcessColl=$oWMI.ExecQuery("Select * from Win32_Process where Name= " & '"'& $Proc & '"')

		For $Process In $oProcessColl

			; search all chromedriver processes for any with no child chrome process.
			;  these can be killed

			$chromedriver_child_process_arr = _ChildProcess($Process.ProcessId)
;			_ArrayDisplay($chromedriver_child_process_arr)

			Local $child_chrome_process_found = False

			for $i = 0 to (UBound($chromedriver_child_process_arr) - 1)

				if StringCompare($chromedriver_child_process_arr[$i][1], "chrome.exe") = 0 Then

					$child_chrome_process_found = True
					ExitLoop
				EndIf
			Next

			; if no chrome found for the chromedriver then no reason to keep it running.

			if $child_chrome_process_found = False Then

				ProcessClose($Process.ProcessId)
			EndIf

		Next






;		$proc="chrome.exe"
;		$oWMI=ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
;		$oProcessColl=$oWMI.ExecQuery("Select * from Win32_Process where Name= " & '"'& $Proc & '"')

;		For $Process In $oProcessColl

			; search all chrome processes for one containing a PID of an automated script (-scriptpid)

;			if StringInStr($Process.Commandline, "--scriptpid-") > 0 Then

;				Local $arr = StringRegExp($Process.Commandline, ".*--scriptpid-(\d++)", 3)

;				if UBound($arr) >= 1 Then

					; if the PID of the automated script no longer for this scriptpid, then we should kill the chromedriver associated with this chrome process

;					if ProcessExists(Number($arr[0])) = False Then

;						ConsoleWrite("This chrome PID " & $Process.ProcessId & " is related to Script PID " & Number($arr[0]) & " which no longer exists" & @CRLF)
;						Local $chromedriver_pid = _WinAPI_GetParentProcess($Process.ProcessId)
;						ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $chromedriver_pid = ' & $chromedriver_pid & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;						ProcessClose($chromedriver_pid)
;					EndIf
;				EndIf
;			EndIf



	;		if StringInStr($Process.Commandline, "--type=gpu-process") > 0 Then

	;			ProcessClose($Process.ProcessId)
	;		EndIf

			; Sean G 15 Sep 2020 - below process closes are potentially crashing latest versions of Chrome and messing with windows desktop

	;		if StringInStr($Process.Commandline, "--extension-process") > 0 Then

	;			ProcessClose($Process.ProcessId)
	;		EndIf

	;		if StringInStr($Process.Commandline, "--type=crashpad-handler") > 0 Then

	;			ProcessClose($Process.ProcessId)
	;		EndIf

	;		if StringInStr($Process.Commandline, "--type=watcher") > 0 Then

	;			ProcessClose($Process.ProcessId)
	;		EndIf
		;Next

		; Bring the command prompts to the front, so they are on top of the Replay App

		WinSetOnTop("[REGEXPTITLE:.*Executable.*; CLASS:ConsoleWindowClass]", "", 1)

		; Force all Chrome processes to idle priority (thus freeing more CPU to the host)

		Local $arr = ProcessList("chrome.exe")

		for $i = 1 to $arr[0][0]

			Local $priority = _ProcessGetPriority($arr[$i][1])

			if $priority > 0 Then

				ProcessSetPriority($arr[$i][1], 0)
			EndIf
		Next

		; Reconnect the S drive (in case it drops)

		Local $s_drive_status = DriveStatus("S:")

		if StringCompare($s_drive_status, "INVALID") = 0 Then

			Local $explorer_pid = ShellExecute("explorer", "S:\", "", "", @SW_HIDE)
			Sleep(100)
			ProcessClose($explorer_pid)
		EndIf

	EndIf

	; every 3 minutes check the following
	if $timer > (3 * 60 * 1000) Then

		$timer = 0

		; Check and turn off UAC
		if RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "ConsentPromptBehaviorAdmin") <> 0 or @error <> 0 Then

			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "ConsentPromptBehaviorAdmin", "REG_DWORD", 0)
		EndIf

		if RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Policies\System", "PromptOnSecureDesktop") <> 0 or @error <> 0 Then

			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Policies\System", "PromptOnSecureDesktop", "REG_DWORD", 0)
		EndIf

		; Check and force Screen lock off...
		if RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaveActive") <> 0 or @error <> 0 Then

			RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaveActive", "REG_SZ", 0)
		EndIf

		if RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaverIsSecure") <> 0 or @error <> 0 Then

			RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaverIsSecure", "REG_SZ", 0)
		EndIf

		if RegRead("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaveTimeOut") <> 0 or @error <> 0 Then

			RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop", "ScreenSaveTimeOut", "REG_SZ", 0)
		EndIf

		; Check and force Java control panel settings
		if FileExists("C:\Windows\Sun\Java\Deployment\deployment.config") = False Then

			FileWrite("C:\Windows\Sun\Java\Deployment\deployment.config", "deployment.system.config=file\:C\:/WINDOWS/Sun/Java/Deployment/deployment.properties")
		EndIf

		if FileExists("C:\Windows\Sun\Java\Deployment\deployment.properties") = False Then

			FileWrite("C:\Windows\Sun\Java\Deployment\deployment.properties", "deployment.security.mixcode=HIDE_RUN" & @CRLF & "deployment.security.revocation.check=NO_CHECK" & @CRLF & "deployment.security.level=MEDIUM")
		EndIf

		; If Antimalware exists then disable it
		if ProcessExists("MsMpEng.exe") = True Then

			stop_disable_service("MsMpSvc")
		EndIf

		; if the mouse pointer has moved, and if not move it to stop the screensaver activating in Windows
		$curr_mouse_pos = MouseGetPos()

		if $curr_mouse_pos[0] = $old_mouse_pos[0] and $curr_mouse_pos[1] = $old_mouse_pos[1] Then

			$mouse_move_pixels = $mouse_move_pixels * -1
			MouseMove($curr_mouse_pos[0] + $mouse_move_pixels, $curr_mouse_pos[1], 1)
			$curr_mouse_pos = MouseGetPos()
		EndIf

		$old_mouse_pos = $curr_mouse_pos
	EndIf

	; slim chrome


	; If the Firefox Authentication Required control is displayed
	if WinExists("[TITLE:Authentication Required; CLASS:MozillaDialogClass]") Then

		; Simply confirm it
		$hand = WinGetHandle("[TITLE:Authentication Required; CLASS:MozillaDialogClass]")
		ControlSend($hand, "", "", "{ENTER}")
	EndIf

	; if the Java Security Warning "Do you want to Continue?" appears
	if WinExists("[TITLE:Security Warning; CLASS:SunAwtDialog]") Then

		WinClose("[TITLE:Security Warning; CLASS:SunAwtDialog]")
	EndIf

	if WinExists("Security Warning", "More information") Then

		$pos = WinGetPos("Security Warning", "More information")
		MouseClick("left", Number($pos[0]) + 25, Number($pos[1]) + Number($pos[3]) - 55, 1, 0)
		ControlClick("Security Warning", "More information", "[CLASS:Button; INSTANCE:1]")
	EndIf

	if WinExists("Security Information") Then

		$pos = WinGetPos("Security Information")
		MouseClick("left", Number($pos[0]) + 25, Number($pos[1]) + Number($pos[3]) - 85, 1, 0)
		SendKeepActive("Security Information")
		Send("{ENTER}")
	EndIf

	if WinExists("Java Update Needed") Then

		$pos = WinGetPos("Java Update Needed")
		MouseClick("left", Number($pos[0]) + 20, Number($pos[1]) + Number($pos[3]) - 20, 1, 0)
		ControlClick("Java Update Needed", "", "[CLASS:Button; INSTANCE:3]")
	EndIf

	if WinExists("System Center Endpoint Protection", "At risk") Then

		WinClose("System Center Endpoint Protection", "At risk")
		Sleep(1000)
	EndIf

	if WinExists("[TITLE:Confirm setting cookie; CLASS:MozillaDialogClass]") Then

		$hand = WinGetHandle("[TITLE:Confirm setting cookie; CLASS:MozillaDialogClass]")
		ControlSend($hand, "", "", "{ENTER}")
	EndIf

	; Close only Firefox popups titled "Log In" (from Siebel in Delivery Train)
	if WinExists("Log In - Mozilla Firefox") = true Then

		Local $hand = WinGetHandle("Log In - Mozilla Firefox")
		Local $pid = WinGetProcess($hand)

		Local $arr = WinList("[CLASS:MozillaWindowClass]")
		Local $num_other_windows_with_same_pid = 0

		for $i = 1 to $arr[0][0]

			if $arr[$i][1] <> $hand Then

				if StringLen($arr[$i][0]) > 0 Then

					Local $tmp_pid = WinGetProcess($arr[$i][1])

					if $tmp_pid = $pid Then

						$num_other_windows_with_same_pid = $num_other_windows_with_same_pid + 1
					EndIf
				EndIf
			EndIf
		Next

		if $num_other_windows_with_same_pid > 0 Then

			WinClose($hand)
		EndIf
	EndIf

	; Close only Firefox popups titled "Spark - Welcome" (from Siebel in Delivery Train)
	if WinExists("Spark - Welcome - Mozilla Firefox") = true Then

		Local $hand = WinGetHandle("Spark - Welcome - Mozilla Firefox")
		Local $pid = WinGetProcess($hand)

		Local $arr = WinList("[CLASS:MozillaWindowClass]")
		Local $num_other_windows_with_same_pid = 0

		for $i = 1 to $arr[0][0]

			if $arr[$i][1] <> $hand Then

				if StringLen($arr[$i][0]) > 0 Then

					Local $tmp_pid = WinGetProcess($arr[$i][1])

					if $tmp_pid = $pid Then

						$num_other_windows_with_same_pid = $num_other_windows_with_same_pid + 1
					EndIf
				EndIf
			EndIf
		Next

		if $num_other_windows_with_same_pid > 0 Then

			WinClose($hand)
		EndIf
	EndIf

	if WinExists("[REGEXPTITLE:Apache Tomcat.*Error report.*]") = true Then

		WinClose("[REGEXPTITLE:Apache Tomcat.*Error report.*]")
	EndIf

	; Stop VMTools
	if ProcessExistsForCurrentUser("vmtoolsd.exe") = True Then

		KillProcessForCurrentUser("vmtoolsd.exe")
		_CleanTrayIcons()
	EndIf

	; Stop FwcMgmt
	if ProcessExistsForCurrentUser("FwcMgmt.exe") = True Then

		KillProcessForCurrentUser("FwcMgmt.exe")
		_CleanTrayIcons()
	EndIf

	; Stop Receiver
	if ProcessExistsForCurrentUser("Receiver.exe") = True Then

		KillProcessForCurrentUser("Receiver.exe")
		_CleanTrayIcons()
	EndIf

	; Stop msseces
	if ProcessExistsForCurrentUser("msseces.exe") = True Then

		KillProcessForCurrentUser("msseces.exe")
		_CleanTrayIcons()
	EndIf

	; Stop picaTWIHost
	if ProcessExistsForCurrentUser("picaTWIHost.exe") = True Then

		KillProcessForCurrentUser("picaTWIHost.exe")
		_CleanTrayIcons()
	EndIf

	; GS-Calc Insert Rows
	if WinExists("Insert Rows") Then

		Local $insert_rows_hand = WinGetHandle("Insert Rows")
		Local $hand = ControlGetHandle($insert_rows_hand, "", 1)

		if StringCompare(ControlGetText($insert_rows_hand, "", 1), "OK") = 0 Then

			ControlClick($insert_rows_hand, "", 1)
		EndIf
	EndIf

	; GS-Calc Insert Columns
	if WinExists("Insert Columns") Then

		Local $insert_columns_hand = WinGetHandle("Insert Columns")
		Local $hand = ControlGetHandle($insert_columns_hand, "", 1)

		if StringCompare(ControlGetText($insert_columns_hand, "", 1), "OK") = 0 Then

			ControlClick($insert_columns_hand, "", 1)
		EndIf
	EndIf

	; GS-Calc Error window when CSV already open
	if WinExists("Open File") Then

		Local $open_file_hand = WinGetHandle("Open File")
		Local $hand = ControlGetHandle($open_file_hand, "", 2)

		if StringCompare(ControlGetText($open_file_hand, "", 2), "Cancel") = 0 Then

			Local $parent_hand = _WinAPI_GetParent($open_file_hand)
			WinClose($open_file_hand)
			WinWaitClose($open_file_hand)
			WinClose($parent_hand)
		EndIf
	EndIf

	; GS-Calc Text Import
	if WinExists("Open Text File") Then

		Local $open_text_file_hand = WinGetHandle("Open Text File")
		Local $parent_hand = _WinAPI_GetParent($open_text_file_hand)
		ControlCommand($open_text_file_hand, "", "[CLASS:ComboBox; INSTANCE:5]", "SelectString", "UTF-8")
		ControlClick($open_text_file_hand, "", "[CLASS:Button; INSTANCE:9]")
		WinWaitClose($open_text_file_hand)

		WinActivate($parent_hand)

		; Briefly block user input for macro run below
		BlockInput(1)

		; Run a GS-Calc script that sets column width to auto for the entire worksheet
		ControlSend($parent_hand, "", "[CLASS:TableView; INSTANCE:1]", "{CTRLDOWN}{SHIFTDOWN}{F1}{SHIFTUP}{CTRLUP}")

		; Re-enable user input
		BlockInput(0)
	EndIf

	; GS-Calc put single quote at front of edited cell
	if ControlCommand("[REGEXPTITLE:GS-Calc 16.1 .*]", "", "[CLASS:Scintilla; INSTANCE:2]", "IsVisible") = 1 and $text_set = False Then

		$hand = ControlGetHandle("[REGEXPTITLE:GS-Calc 16.1 .*]", "", "[CLASS:Scintilla; INSTANCE:2]")

		BlockInput(1)
		$curr_text = _GUICtrlEdit_GetText($hand)
		gs_calc_edit_set_text($hand, "'" & $curr_text)
		BlockInput(0)

		$text_set = True
	EndIf

	if ControlCommand("[REGEXPTITLE:GS-Calc 16.1 .*]", "", "[CLASS:Scintilla; INSTANCE:2]", "IsVisible") = 0 Then

		$text_set = False
	EndIf

	; GS-Calc add "Ctrl + F" for finding
	if WinActive("[REGEXPTITLE:GS-Calc 16.1 .*]") = True and $hotkeyset = False Then

		$hotkeyset = True
		HotKeySet("^f", gs_calc_find)
	EndIf

	if WinActive("[REGEXPTITLE:GS-Calc 16.1 .*]") = False and $hotkeyset = True Then

		$hotkeyset = False
		HotKeySet("^f")
	EndIf

	; GS-Calc add "Ctrl + G" for finding
	if WinActive("[REGEXPTITLE:GS-Calc 16.1 .*]") = True and $hotkey2set = False Then

		$hotkey2set = True
		HotKeySet("^g", gs_calc_find_assigned_to_and_field)
	EndIf

	if WinActive("[REGEXPTITLE:GS-Calc 16.1 .*]") = False and $hotkey2set = True Then

		$hotkey2set = False
		HotKeySet("^g")
	EndIf

	if WinExists("Run As") Then

		ControlSend("Run As", "", "[CLASS:SysListView32; INSTANCE:1]", "{HOME}")
		ControlClick("Run As", "", "[CLASS:Button; INSTANCE:1]")
		Sleep(1000)
	EndIf

	if WinExists("Debug As") Then

		ControlSend("Debug As", "", "[CLASS:SysListView32; INSTANCE:1]", "{HOME}")
		ControlClick("Debug As", "", "[CLASS:Button; INSTANCE:1]")
		Sleep(1000)
	EndIf

	if WinExists("[TITLE:Restoring Network Connections]") = true Then

		WinClose("[TITLE:Restoring Network Connections]")
	EndIf

	Sleep($loop_delay)

WEnd


Func ProcessExistsForCurrentUser($process_name)


   Local $aAdjust, $aList = 0

   ; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
   Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
   _WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

   ; Retrieve user names for all processes the system
   If Not (@error Or @extended) Then
	   $aList = ProcessList($process_name)
	   Local $aData
	   For $i = 1 To $aList[0][0]
		   $aData = _WinAPI_GetProcessUser($aList[$i][1])

		   If IsArray($aData) Then

			  if StringCompare($aData[0], @username) = 0 Then

				  Return True
			   EndIf
		   EndIf
	   Next
   EndIf

   ; Enable SeDebugPrivilege privilege by default
   _WinAPI_AdjustTokenPrivileges($hToken, $aAdjust, 0, $aAdjust)
   _WinAPI_CloseHandle($hToken)

;   _ArrayDisplay($aList, '_WinAPI_GetProcessUser')

   return False
EndFunc


Func KillProcessForCurrentUser($process_name)


   Local $aAdjust, $aList = 0

   ; Enable "SeDebugPrivilege" privilege for obtain full access rights to another processes
   Local $hToken = _WinAPI_OpenProcessToken(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
   _WinAPI_AdjustTokenPrivileges($hToken, $SE_DEBUG_NAME, $SE_PRIVILEGE_ENABLED, $aAdjust)

   ; Retrieve user names for all processes the system
   If Not (@error Or @extended) Then
	   $aList = ProcessList($process_name)
	   Local $aData
	   For $i = 1 To $aList[0][0]
		   $aData = _WinAPI_GetProcessUser($aList[$i][1])

		   If IsArray($aData) Then

			  if StringCompare($aData[0], @username) = 0 Then

				  ConsoleWrite("found " & $aList[$i][1])
				  ProcessClose($aList[$i][1])
				  Return True
			   EndIf
		   EndIf
	   Next
   EndIf

   ; Enable SeDebugPrivilege privilege by default
   _WinAPI_AdjustTokenPrivileges($hToken, $aAdjust, 0, $aAdjust)
   _WinAPI_CloseHandle($hToken)

;   _ArrayDisplay($aList, '_WinAPI_GetProcessUser')

   return False
EndFunc


func _CleanTrayIcons()

   $count = _SysTrayIconCount()
   For $i = $count - 1 To 0 Step -1
	   $handle = _SysTrayIconHandle($i)
	   $pid = WinGetProcess($handle)
	   If $pid = -1 Then _SysTrayIconRemove($i)
   Next

EndFunc


Func block_mouse()

	_BlockInputEx(2)
	SplashImageOn("", "c:\mouse-off.gif", 128, 128, @DesktopWidth - 128 - 20, @DesktopHeight - 128 - 60, 3)
EndFunc

Func unblock_all()

	_BlockInputEx(0)
	SplashImageOn("", "c:\mouse-on.gif", 128, 128, @DesktopWidth - 128 - 20, @DesktopHeight - 128 - 60, 3)
EndFunc

Func SINK_OnObjectReady($objLatestEvent, $objAsyncContext)
    FileDelete(@ScriptDir & "\vnc_pass_update.txt")
	FileWrite(@ScriptDir & "\vnc_pass_update.txt", @UserName & "|" & @MDAY & "/" & @MON & "/" & @YEAR & " " & @HOUR & ":" & @MIN)
EndFunc   ;==>SINK_OnObjectReady


Func gs_calc_find()

	$find_text = InputBox("GS-Calc Find", "Enter the text to find", $find_text)

	if StringLen($find_text) > 0 Then

		; Briefly block user input for macro run below
		BlockInput(1)

		local $visible = ControlCommand("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "IsVisible")

		if $visible = 0 Then

			ControlSend("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "{F3}")
		EndIf

		local $ret = ControlSetText("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "")
		ControlSend("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, $find_text & "{F3}")

		; Re-enable user input
		BlockInput(0)
	EndIf
EndFunc

Func gs_calc_find_assigned_to_and_field()

	$find_assigned_to_and_field_text = InputBox("GS-Calc Find Assigned to & Field", "Enter <Assigned to value>,<Field name>", $find_assigned_to_and_field_text)

	if StringLen($find_assigned_to_and_field_text) > 0 and StringInStr($find_assigned_to_and_field_text, ",") > 0 Then

		; Briefly block user input for macro run below
		BlockInput(1)

		; Set the current address to "A1"
		$current_address_hand = ControlGetHandle("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 380)
		gs_calc_edit_set_text($current_address_hand, "sheet1!A1")

		; Click the Go to this reference button
		$toolbar_hand = ControlGetHandle("[REGEXPTITLE:GS-Calc 16.1 .*]", "", "[CLASS:ToolbarWindow32; INSTANCE:5]")
		_GUICtrlToolbar_ClickIndex($toolbar_hand, 24)

		Local $find_text_part = StringSplit($find_assigned_to_and_field_text, ",", 3)

		; Find the field name and column reference
		gs_calc_find_text($find_text_part[1])
		$current_address = ControlGetText("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 380)
		$column_ref = StringReplace($current_address, "sheet1!", "")
		$column_ref = StringRegExpReplace($column_ref, "[0-9]", "")
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $column_ref = ' & $column_ref & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		; Find the Assigned to and row reference
		gs_calc_find_text($find_text_part[0])
		$current_address = ControlGetText("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 380)
		$current_address = StringRegExpReplace($current_address, "sheet1![A-Z]*", "sheet1!" & $column_ref)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $current_address = ' & $current_address & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		; Set the current address
		gs_calc_edit_set_text($current_address_hand, $current_address)

		; Click the Go to this reference button
		_GUICtrlToolbar_ClickIndex($toolbar_hand, 24)

		; Re-enable user input
		BlockInput(0)
	EndIf
EndFunc

Func gs_calc_edit_set_text($hand, $text)

	_GUICtrlEdit_SetText($hand, "")

	if StringLen($text) > 0 Then

		Local $arr = StringToASCIIArray($text)

		for $i = 0 to (UBound($arr) - 1)

			_GUICtrlEdit_AppendText($hand, Chr($arr[$i]))
		Next
	EndIf

EndFunc

func gs_calc_find_text($text)

	local $visible = ControlCommand("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "IsVisible")

	if $visible = 0 Then

		ControlSend("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "{F3}")
	EndIf

	local $ret = ControlSetText("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, "")
	ControlSend("[REGEXPTITLE:GS-Calc 16.1 .*]", "", 1006, $text & "{F3}")

EndFunc

Func stop_disable_service($service_name)

	ShellExecuteWait(@ScriptDir & "\SetACL.exe", "-on """ & $service_name & """ -ot srv -actn ace -ace ""n:Users;p:full""", "", "", @SW_HIDE)
	_Service_SetStartType($service_name, $SERVICE_DISABLED)
	_Service_Stop($service_name)
EndFunc


;===============================================================================
; Function Name:    _ChildProcess
; Description:      Returns an array containing child process info
; Parameter(s):     $iParentPID - Parent Process PID
;
; Requirement(s):   AutoIt 3.8+
; Return Value(s):  two dimensional array $aChildren[][] as follows
;                   $aChildren[0][0] = Number of children found
;
;                   $aChildren[x][0] = Child Process's PID (called Handle, but not a handle object)
;                   $aChildren[x][1] = Child Process's Name
;                   $aChildren[x][2] = Child Process's Command Line
;                   $aChildren[x][3] = Child Process's Executable Path
;                   $aChildren[x][4] = Child Process's Creation Date and Time
;                   $aChildren[x][5] = Child Process's Session Id
;                   $aChildren[x][6] = Child Process's Status
;                   $aChildren[x][7] = Child Process's Termination Date
;
; AutoIt Version:   3.2.8.1
; Author:           JerryD
;===============================================================================

Func _ChildProcess ( $iParentPID )
    Local Const $wbemFlagReturnImmediately = 0x10, $wbemFlagForwardOnly = 0x20
    Local Const $sQuery = 'SELECT * FROM Win32_Process Where ParentProcessId = ' & $iParentPID & ' AND Handle <> ' & $iParentPID
    Local $aChildren[1][8]
    $aChildren[0][0] = 0
    $aChildren[0][1] = 'Name'
    $aChildren[0][2] = 'Command Line'
    $aChildren[0][3] = 'Executable Path'
    $aChildren[0][4] = 'Creation Date and Time'
    $aChildren[0][5] = 'Session Id'
    $aChildren[0][6] = 'Status'
    $aChildren[0][7] = 'Termination Date'

    Local $objWMIService = ObjGet ( 'winmgmts:\\localhost\root\CIMV2' )
    If NOT IsObj ( $objWMIService ) Then
        SetError ( 1 )
        Return $aChildren
    EndIf
    Local $colItems = $objWMIService.ExecQuery ( $sQuery, 'WQL', $wbemFlagReturnImmediately + $wbemFlagForwardOnly )
    If IsObj($colItems) then
        For $objItem In $colItems
            $aChildren[0][0] += 1
            ReDim $aChildren[$aChildren[0][0]+1][8]
            $aChildren[$aChildren[0][0]][0] = $objItem.Handle
            $aChildren[$aChildren[0][0]][1] = $objItem.Name
            $aChildren[$aChildren[0][0]][2] = $objItem.CommandLine
            $aChildren[$aChildren[0][0]][3] = $objItem.ExecutablePath
            $aChildren[$aChildren[0][0]][4] = $objItem.CreationDate
            $aChildren[$aChildren[0][0]][5] = $objItem.SessionId
            $aChildren[$aChildren[0][0]][6] = $objItem.Status
            $aChildren[$aChildren[0][0]][7] = $objItem.TerminationDate
        Next
    Else
        SetError ( 2 )
        Return $aChildren
    EndIf
    SetError ( 0 )
    Return $aChildren
EndFunc
