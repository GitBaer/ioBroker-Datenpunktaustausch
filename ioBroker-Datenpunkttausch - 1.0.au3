#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_Outfile=iOB-Datenpunktaustausch.Exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
Skript zum Ersetzen von ioBroker-Datenpunkten

Version: 1.0 (18.08.2021)

GitBaer / IOBaer
#ce

#include <MsgBoxConstants.au3>
#include <array.au3>
#include <WinAPIFiles.au3>
#include <FileConstants.au3>
#include <GuiEdit.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

#Region ### START Koda GUI section ###
$iob = GUICreate("ioBroker Datenpunktaustausch", 460, 226, 0, 0)
$start = GUICtrlCreateButton("Zwischenablage einlesen und alle ersetzen", 8, 40, 233, 25)
$Status = GUICtrlCreateEdit("", 8, 72, 441, 145, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_VSCROLL))
$Label1 = GUICtrlCreateLabel("INI-Datei mit Ersetzungen:", 8, 11, 127, 17)
$Pfad = GUICtrlCreateInput("", 136, 8, 281, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_READONLY))
$PfadAuswahl = GUICtrlCreateButton("...", 422, 6, 25, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

WinSetOnTop($iob, "", 1)

$Ergebnispruefung = ""

If FileExists(@ScriptDir & "\ersetzungen.ini") Then
	$inipfad = @ScriptDir & "\ersetzungen.ini"
	GUICtrlSetData($Pfad, $inipfad)
	_Logfile("INI-Datei erfolgreich ausgewählt")
Else
	GUICtrlSetData($Pfad, "Bitte INI-Datei wählen -->")
	_Logfile("INI-Datei konnte nicht automatisch gefunden werden.")
	_Logfile("--> Bitte INI-Datei manuell wählen.")
	$inipfad = ""

EndIf

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $PfadAuswahl
			WinSetOnTop($iob, "", 0)
			$inipfad = FileOpenDialog("Datei mit Ersetzungen wählen ...", @ScriptDir & "\", "INI-Einstellungsdatei (*.ini)", BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST, "ersetzungen.ini"))
			If @error Then
				WinSetOnTop($iob, "", 1)
				GUICtrlSetData($Pfad, "Bitte INI-Datei wählen -->")
				$inipfad = ""
				_Logfile("Fehler beim Auswählen der INI-Einstellungsdatei")
			Else
				WinSetOnTop($iob, "", 1)
				GUICtrlSetData($Pfad, $inipfad)
				_Logfile("INI-Datei erfolgreich ausgewählt")
			EndIf

		Case $start
			_Ersetzung()
	EndSwitch
WEnd

Func _Ersetzung()
	_Logfile("Skript gestartet ...")
	$clipeingelesen = ClipGet()
	If @error = 0 Then
		_Logfile('Zwischenablage erfolgreich eingelesen ("' & StringLeft($clipeingelesen, 20) & ' (...)" ...')
	Else
		_Logfile("Fehler beim Einlesen der Zwischenablage.")
		_Logfile("--> Skript wird abgebrochen.")
		Return
	EndIf
	$vergleich = StringCompare($clipeingelesen, $Ergebnispruefung)
	If $vergleich = 0 Then
		_Logfile("Fehler: Text der Zwischenablage wurde bereits verarbeitet.")
		_Logfile("--> Skript wird abgebrochen.")
		Return
	EndIf
	$IniToArray = IniReadSection($inipfad, "ioBroker")
	If @error = 0 Then
		_Logfile("Ini-Datei mit Ersetzungen erfolgreich eingelesen ...")
	Else
		_Logfile("Fehler beim Einlesen der INI-Datei.")
		_Logfile("--> Skript wird abgebrochen.")
		Return
	EndIf
	For $i = 1 To UBound($IniToArray, 1) - 1
		$clipeingelesen = StringReplace($clipeingelesen, $IniToArray[$i][0], $IniToArray[$i][1])
		_Logfile('String "' & $IniToArray[$i][0] & '" ' & @extended & ' x durch String "' & $IniToArray[$i][1] & '" ersetzt')
	Next
	$Ergebnispruefung = $clipeingelesen
	ClipPut($clipeingelesen)
	If @error = 0 Then
		_Logfile("Ergebnis erfolgreich in Zwischenablage kopiert ...")
	Else
		_Logfile("Fehler beim Kopieren in die Zwischenablage.")
		_Logfile("--> Skript wird abgebrochen.")
		Return
	EndIf
	_Logfile("Skript erfolgreich beendet.")
EndFunc   ;==>_Ersetzung

GUICtrlRead($Status)
Func _Logfile($logtext)
	GUICtrlSetData($Status, GUICtrlRead($Status) & @HOUR & ":" & @MIN & ":" & @SEC & " - " & $logtext & @CRLF)
	GUICtrlSendMsg($Status, "0xB7", 0, 0)
EndFunc   ;==>_Logfile
