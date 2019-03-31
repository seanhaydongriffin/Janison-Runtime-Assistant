;!include nsDialogs.nsh
;!include LogicLib.nsh


; example1.nsi
;
; This script is perhaps one of the simplest NSIs you can make. All of the
; optional settings are left to their default settings. The installer simply 
; prompts the user asking them where to install, and drops a copy of example1.nsi
; there. 

XPStyle on

;--------------------------------

; The name of the installer
Name "Janison Runtime Assistant"

; The file to write
OutFile "setup.exe"

; The default installation directory
InstallDir "C:\Janison Runtime Assistant"

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------


; Pages

Page directory
Page instfiles


;--------------------------------


; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File "Janison Runtime Assistant.exe"
  File "PointerStick_x64.exe"
  File "query.exe"
  File "SetACL.dll"
  File "SetACL.exe"

  CreateDirectory "$SMPROGRAMS\Janison Runtime Assistant"
  CreateShortCut "$SMPROGRAMS\Janison Runtime Assistant\Janison Runtime Assistant.lnk" "$INSTDIR\Janison Runtime Assistant.exe"

SectionEnd ; end the section
