#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\..\OneDrive\Pictures\konecta.ico
#AutoIt3Wrapper_Compression=0
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <MsgBoxConstants.au3>
#include <FTPEx.au3>
#include <Array.au3>
#include <File.au3>
#include <Process.au3>

Opt("TrayIconHide", 1)


Global $fileGenesis = "purecloud-windows-1.15.576.exe"

$CmdRUNAS = "runas /user:multienlace.com.co\despliegue.forti"
$CmdDIR = " ""C:\Users\"& @UserName &"\Downloads\purecloud-windows-1.15.576.exe"" "

_RunDos("start taskkill /F /IM Chrome.exe")
Sleep(500)
SFTPConnect()
Sleep(500)
ExistFiles()
Sleep(500)
ExecuteAdmin()
Sleep(500)
PureCloud()
Sleep(500)
MsgBox($MB_ICONINFORMATION, "RPAGenesisCloud - Informacion", "El proceso ha finalizado")
Exit

Func SFTPConnect()

   #Region Credencials SFTP
    Local $sServer = '190.131.252.46'
    Local $sUsername = 'despliegue.forti'
    Local $sPass = 'zAKYQl6*'

    Local $Err, $sFTP_Message, $sFileName
   #EndRegion

    Local $hOpen = _FTP_Open('MyFTP Control')
    Local $hConn = _FTP_Connect($hOpen, $sServer, $sUsername, $sPass)

	   If @error Then
		   MsgBox(BitOR(262144,48), 'RPAKonecta - Warning (PureCloud)', 'No se logro establecer la conexion con el servidor SFTP')
		   Exit
	   Else
		   _FTP_GetLastResponseInfo($Err, $sFTP_Message)

		   Local $aFile = _FTP_ListToArray($hConn, 2)

		   For $i = 1 to ubound($aFile) - 1
			 If($aFile[$i] = $fileGenesis) Then
				$sFileName = $aFile[5]
			   ExitLoop
			 EndIf
		  Next

		   $sDirectory =  "C:\Users\"& @UserName &"\Downloads\"

		   ConsoleWrite('FileName = '  & $sFileName  & @CRLF)
		   ConsoleWrite('Directory = ' & $sDirectory & @CRLF)

		   _FTP_FileGet($hConn, $sFileName, $sDirectory & $sFileName)
		   Sleep(1500)
		   _FTP_Close($hConn)

		   ConsoleWrite("Exit" & @CRLF)

	   EndIf

    Local $iFtpc = _FTP_Close($hConn)
    Local $iFtpo = _FTP_Close($hOpen)

EndFunc

Func ExistFiles()

   Local $sDirectory =  "C:\Users\"& @UserName &"\Downloads\"& $fileGenesis

   Local $iFileExists = FileExists($sDirectory)

    If $iFileExists Then
    Else
        MsgBox($MB_SYSTEMMODAL, "", "El Archivo no existe.")
		Exit
    EndIf

EndFunc

Func ExecuteAdmin()

   Run(@ComSpec & "" , @WindowsDir, @SW_SHOW)
   Sleep(2500)
   $validWnd = WinWait("[CLASS:ConsoleWindowClass]", "", 10)

   While 1
	 If WinExists($validWnd) <> 0 Then
		 ExitLoop
	 EndIf
   WEnd

   #Region Credencials Execution EXE
	If WinActivate($validWnd) Then
	  Sleep(500)
	   Send($CmdRUNAS & $CmdDIR,0)
	   Sleep(500)
	   Send("{ENTER}")
	   Sleep(1500)
	   Send("JnPt5**32W",1)
	   Sleep(1500)
	   Send("{ENTER}")
	EndIf
   #EndRegion

   WinClose($validWnd)

EndFunc

Func PureCloud()

   Run("C:\Users\"& @UserName &"\Downloads\"& $fileGenesis, "")

   Local $hWnd = WinWait("[TITLE:Configuración de PureCloud]", "", 5)

   While 1
	  If WinActivate($hWnd, "") Then
		ExitLoop
	  EndIf
   WEnd

   ControlClick($hWnd, "", "[Class:Button;Instance:2]")

   While 1
	  Local $sText = WinGetText($hWnd)

	  If StringInStr($sText, "Configuración finalizada") Then
		ExitLoop
	  EndIf
   WEnd

   ControlClick($hWnd, "", "[Class:Button;Instance:8]")

   While 1
	  Local $hWndPure = WinWait("[TITLE:PureCloud]", "", 5)

	  If WinActivate($hWndPure, "") Then
		ExitLoop
	  EndIf
   WEnd


	Sleep(5000)
	ControlClick($hWndPure, "", "[Class:Chrome_RenderWidgetHostHWND;Instance:1]")
	Sleep(1500)
	Send("{TAB 2}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(1500)
	Send("{TAB 12}")
	Sleep(500)
	Send("{ENTER}")

	Sleep(5000)
	ControlClick($hWndPure, "", "[Class:Chrome_RenderWidgetHostHWND;Instance:1]")
	Sleep(1500)
	Send("{TAB}")
	Sleep(500)
	Send("{ENTER}")
	Send("Continente Americano (Oeste de los EE. UU.)",1)
	Sleep(1500)
	Send("{ENTER}")
	Sleep(1500)
	Send("{TAB}")
	Sleep(1500)
	Send("{ENTER}")

#CS 	Sleep(5000)
; 	ControlClick($hWndPure, "", "[Class:Chrome_RenderWidgetHostHWND;Instance:1]")
; 	Sleep(1500)
; 	Send("{TAB}")
; 	Sleep(500)
; 	Send("{ENTER}")
; 	Send("Continente Americano (Oeste de los EE. UU.)",1)
; 	Sleep(1500)
; 	Send("{ENTER}")
; 	Sleep(1500)
; 	Send("{TAB}")
; 	Sleep(1500)
; 	Send("{ENTER}")
 #CE


EndFunc



