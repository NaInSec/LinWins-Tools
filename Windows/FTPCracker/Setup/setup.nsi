; Script generated with the Venis Install Wizard

; Define your application name
!define APPNAME "FTP Password Cracker"
!define APPNAMEANDVERSION "FTP Password Cracker 1.00.0034"

; Main Install settings
Name "${APPNAMEANDVERSION}"
InstallDir "$PROGRAMFILES\FTP Password Cracker"
InstallDirRegKey HKLM "Software\${APPNAME}" ""
OutFile "Setup.exe"

; Modern interface settings
!include "MUI.nsh"

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\FTPCracker.exe"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Set languages (first is default language)
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_RESERVEFILE_LANGDLL

Section "FTP Password Cracker" Section1

	; Set Section properties
	SetOverwrite on

	; Set Section Files and Shortcuts
	SetOutPath "$INSTDIR\"
	File "FTPCracker.exe"
	File "passwords.txt"
	File "usernames.txt"
	File "ksyAlphaWin.ocx"	
	
	SetOverwrite off
	SetOutPath "$SYSDIR\"
	File "COMCAT.DLL"
	File "COMDLG32.OCX"
	File "MSCOMCTL.OCX"
	File "MSWINSCK.OCX"
	File "asycfilt.dll"
	File "msvbvm60.dll"
	File "oleaut32.dll"
	File "olepro32.dll"

	CreateDirectory "$SMPROGRAMS\FTP Password Cracker"
	CreateShortCut "$SMPROGRAMS\FTP Password Cracker\FTP Password Cracker.lnk" "$INSTDIR\FTPCracker.exe"
	CreateShortCut "$SMPROGRAMS\FTP Password Cracker\Uninstall.lnk" "$INSTDIR\uninstall.exe"

SectionEnd

Section -FinishSection

	WriteRegStr HKLM "Software\${APPNAME}" "" "$INSTDIR"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\uninstall.exe"
	WriteUninstaller "$INSTDIR\uninstall.exe"

SectionEnd

; Modern install component descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${Section1} ""
!insertmacro MUI_FUNCTION_DESCRIPTION_END

;Uninstall section
Section Uninstall

	;Remove from registry...
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
	DeleteRegKey HKLM "SOFTWARE\${APPNAME}"

	; Delete self
	Delete "$INSTDIR\uninstall.exe"

	; Delete Shortcuts
	Delete "$SMPROGRAMS\FTP Password Cracker\FTP Password Cracker.lnk"
	Delete "$SMPROGRAMS\FTP Password Cracker\Uninstall.lnk"

	; Clean up FTP Password Cracker
	Delete "$INSTDIR\FTPCracker.exe"
	Delete "$INSTDIR\passwords.txt"
	Delete "$INSTDIR\usernames.txt"
	delete "$INSTDIR\ksyAlphaWin.ocx"
	; Clean up remaining dirs
	RMDir "$SMPROGRAMS\FTP Password Cracker"
	RMDir "$INSTDIR\"

SectionEnd

BrandingText "FTP Password Cracker"

; eof