#NoTrayIcon
#include <Array.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <TrayConstants.au3>

; Init stuff
Opt("TrayMenuMode", 3)
Global $devices
Global $trayIds[0]
Global $trayExit
Global $trayRefresh

; Get the sound device libary
FileInstall("EndPointController.exe", @TempDir & "\EndPointController.exe")
FileInstall("Audio.EndPoint.Controller.dll", @TempDir & "\Audio.EndPoint.Controller.dll")
Global $DEVICE_CONTROLLER = @TempDir & "\EndPointController.exe";

; Create the tray
CreateTray()
UpdateTray()

While 1
   HandleTrayAction()
WEnd

; Functions

Func UpdateDevices()

   ; Get all devices
   $devices = StringSplit(_getDOSOutput($DEVICE_CONTROLLER & " -f %d:%ws:%d:%d"), @CRLF)
   $size = $devices[0]

   ; Remove the leading index
   For $i = 1 To $size
	  $vals = StringSplit($devices[$i],":",2);
	  $devices[$i] = $vals[1]

	  ; Get the active device
	  if $vals[3] == "1" Then
		 $devices[0] = $vals[1]
	  EndIf
   Next
EndFunc

Func SetDevice($deviceName)

   ; Get a recent device list
   $devs = StringSplit(_getDOSOutput($DEVICE_CONTROLLER & " -f %d:%ws"), @CRLF)

   ; Find the device
   For $i = 1 To $devs[0]
	  $vals = StringSplit($devs[$i],":",2);
	  if $vals[1] == $deviceName Then
		 ; Set the device and return
		 _getDOSOutput($DEVICE_CONTROLLER & " " & $vals[0])
		 Return
	  EndIf
   Next

   ; If no device has been found: show an error!
   MsgBox($MB_OK && $MB_ICONERROR, "Device list not up to date!", "The device could not be found! You may need to refresh the list.")

EndFunc

Func CreateTray()

   ; The exit & refresh entries
   $trayExit = TrayCreateItem("Exit")
   $trayRefresh = TrayCreateItem("Refresh devices")
   TrayCreateItem("")

   ; Show the tray
   TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

   ; Set the text!
   TraySetToolTip("Change the audio playback device")
EndFunc

Func HandleTrayAction()
   $msg = TrayGetMsg()
   Switch $msg
   Case $trayExit
	  Exit 0;
   Case $trayRefresh
	  UpdateTray()
	  Return
   EndSwitch
   if $msg > 0 Then
	  HandleClickDevice($msg)
   EndIf
EndFunc

Func HandleClickDevice($id)
   For $i = 0 To UBound($trayIds) - 1
	  If $trayIds[$i] == $id Then
		 SetDevice($devices[$i + 1])
		 UpdateTray()
	  EndIf
   Next
EndFunc

Func ClearTray()
   For $id In $trayIds
	  TrayItemDelete($id)
   Next
EndFunc

Func UpdateTray()

   ; Clear the tray!
   ClearTray()

   ; Update the devices
   UpdateDevices()

   ; The device entries
   $size = UBound($devices)
   ReDim $trayIds[$size - 1]
   For $i = 1 To $size - 1

	  ; Create the tray item
	  $id = TrayCreateItem($devices[$i], -1, -1, $TRAY_ITEM_RADIO)

	  ; Set the active item as checked
	  If $devices[$i] == $devices[0] Then
		 TrayItemSetState(-1, $TRAY_CHECKED)
	  EndIf

	  ; remember the tray item id
	  $trayIds[$i-1] = $id
   Next

EndFunc

Func _getDOSOutput($command)
    Local $text = '', $Pid = Run('"' & @ComSpec & '" /c ' & $command, '', @SW_HIDE, 2 + 4)
    While 1
            $text &= StdoutRead($Pid, False, False)
            If @error Then ExitLoop
            Sleep(10)
    WEnd
    Return StringStripWS($text, 7)
EndFunc
