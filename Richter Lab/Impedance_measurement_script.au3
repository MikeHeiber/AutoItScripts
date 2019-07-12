; Copyright (c) 2019 Michael C. Heiber
; This script file is part of the AutoItScripts repository, which is subject to the MIT License.
; For more information, see the LICENSE file that accompanies this repository.
; The AutoItScripts repository can be found on Github at https://github.com/MikeHeiber/AutoItScripts
;
; Measurement script by Michael C. Heiber
;
; Sample Info
Local $device_name = "CW4_2A"
; Define the LED current range to be tested
;Local $LED_array[] = [0, 14, 15, 16, 17, 20, 23, 27, 35, 46, 65, 96] ; mA
Local $LED_array[] = [0, 14, 46] ; mA
; Define impedance light and dark preset file names
Local $impedance_light_preset = "IMPEDANCE-LIGHT.PRE"
Local $impedance_dark_preset = "IMPEDANCE-DARK.PRE"
; Define the impedance bias scan measurement frequency
Local $measurement_frequency = 500 ; Hz
; Define how many measurements to average over
Local $N_averaging = 10
; Define where to save the data
Local $data_dir = "C:\Documents and Settings\NovoDRS\My Documents\MikeHeiber\Data\"

; Setup data folder
If Not FileExists($data_dir&$device_name&" Impedance Data") Then
   DirCreate($data_dir&$device_name&" Impedance Data")
EndIf

; Run the LED Controller software if it's not already running
If Not WinActivate("Led Driver Control Panel") Then
   ; Run the LED driver exe
   run("C:\Documents and Settings\NovoDRS\Desktop\Mightex LED Controller\Software\LEDDriver.exe")
   ; Wait for port selection window to load
   WinWaitActive("Port Selection")
   ; Choose the driver port to be USB
   ControlClick("Port Selection", "", "[CLASS:TRzGroupButton; INSTANCE:1]")
   ControlClick("Port Selection", "", "[CLASS:TRzPanel; INSTANCE:1]")
   ; Wait for LED control panel to load
   WinWaitActive("Led Driver Control Panel")
   ; Set the "Parameters Setting Selection" to "Normal Setting"
   ControlClick("Led Driver Control Panel", "", "[CLASS:TRzGroupCheck; INSTANCE:2]")
   ; Set the "Current Mode Selection" to "Normal"
   ControlClick("Led Driver Control Panel", "", "[CLASS:TRzGroupButton; INSTANCE:2]")
   ; Click the checkbox to "Set All Channels"
   ControlClick("Led Driver Control Panel", "", "[CLASS:TRzCheckBox; INSTANCE:1]")
   ; Click the "Set Current Mode" button
   ControlClick("Led Driver Control Panel", "", "[CLASS:TRzBitBtn; INSTANCE:2]")
EndIf
; Set the LED "Max Current" to 1000 mA
ControlSetText("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:3]", "1000")
ControlClick("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:3]")
Send("{ENTER}")
; Set the LED "Set Current" to 0 mA
ControlSetText("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]", "0")
ControlClick("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]")
Send("{ENTER}")
; Click the LED Set Parameters button
ControlClick("Led Driver Control Panel", "", "[CLASS:TRzBitBtn; INSTANCE:6]")

; Run WINDETA if it's not already running
If Not WinActivate("WinDETA") Then
   run("C:\Program Files\Novocontrol\WinDETA\WinDETA.exe")
   ; Close the info popup window
   WinWaitActive("Info")
   ControlClick("Info", "", "[CLASS:Button; INSTANCE:1]")
EndIf

; -- Phase 1 -- dark impedance frequency scan
; Activate the WINDETA window
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   ; Load the dark impedance presets
   Send("!fl") ; Opens load preset window
   WinWaitActive("Load Preset File")
   ControlCommand("Load Preset File", "", 1184, "SendCommandID", 41062)
   ControlClick("Load Preset File", "", "[CLASS:SysListView32; INSTANCE:1]")
   Send("MikeHeiber{ENTER}")
   ControlSetText("Load Preset File", "", "[CLASS:Edit; INSTANCE:1]", $impedance_dark_preset)
   ControlClick("Load Preset File", "", "[CLASS:Button; INSTANCE:2]")
   ; Wait for presets to load
   WinWait("Multi Graphics")
   ; Set the measurement averaging
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   Send("!ma")
   WinWaitActive("Averaging Options")
   Send($N_averaging)
   ControlClick("Averaging Options", "", "[CLASS:Button; INSTANCE:3]")
   ; Run the dark impedance measurement
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   Send("!mt")
   WinWaitActive("Warning")
   ControlClick("Warning", "", "[CLASS:Button; INSTANCE:1]")
   ; Detect when the measurement is finished
   Sleep(10000) ; wait 10 sec
   AutoItSetOption("WinTitleMatchMode", 3) ; Set option that window name must match exactly
   WinWait("Online")
   AutoItSetOption("WinTitleMatchMode", 1) ; Reset to partial match
   ; Save the dark impedance data
   Send("!fa")
   WinWaitActive("Save Result File as ASCII")
   ; Click OK on the save settings
   ControlClick("Save Result File as ASCII", "", "[CLASS:Button; INSTANCE:9]")
   WinWaitActive("Save Result File")
   ; Navigate to device data folder
   Sleep(1000)
   Send("MikeHeiber{ENTER}")
   Sleep(1000)
   Send("Data{ENTER}")
   Sleep(1000)
   Send($device_name&"{ENTER}")
   Sleep(1000)
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   While FileExists($data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".txt")
	  $N_measurement += 1
   WEnd
   ControlSetText("Save Result File", "", "[CLASS:Edit; INSTANCE:1]", $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".txt")
   ControlClick("Save Result File", "", "[CLASS:Button; INSTANCE:2]")

; -- Phase 2 -- light impedance bias scan at each intensity
; Activate the WinDETA window
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   ; Load the light impedance presets
   Send("!fl") ; Opens load preset window
   WinWaitActive("Load Preset File")
   ControlCommand("Load Preset File", "", 1184, "SendCommandID", 41062)
   ControlClick("Load Preset File", "", "[CLASS:SysListView32; INSTANCE:1]")
   Send("MikeHeiber{ENTER}")
   ControlSetText("Load Preset File", "", "[CLASS:Edit; INSTANCE:1]", $impedance_light_preset)
   ControlClick("Load Preset File", "", "[CLASS:Button; INSTANCE:2]")
   ; Wait for presets to load
   WinWait("Online")
   Sleep(2000)
   ; Set the measurement averaging
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   Send("!ma")
   WinWaitActive("Averaging Options")
   Send($N_averaging)
   ControlClick("Averaging Options", "", "[CLASS:Button; INSTANCE:3]")
   ; Set the measurement frequency
   WinActivate("WinDETA")
   WinWaitActive("WinDETA")
   Send("!mc")
   WinWaitActive("Start Conditions")
   Send($measurement_frequency)
   ControlClick("Start Conditions", "", "[CLASS:Button; INSTANCE:9]")

; Loop through the LED_array and measure at each light intensity
For $i = 0 To UBound($LED_array)-1
   ; Activate the LED driver control panel window
	  WinActivate("Led Driver Control Panel")
	  WinWaitActive("Led Driver Control Panel")
	  ; Adjust the LED light intensity by setting the "Set Current" to $LED_array[$i]
	  ControlSetText("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]", $LED_array[$i])
	  ControlClick("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]")
	  Send("{ENTER}")
	  ; Click the LED Set Parameters button
	  ControlClick("Led Driver Control Panel", "", "[CLASS:TRzBitBtn; INSTANCE:6]")

   ; Activate the WINDETA window
	  WinActivate("WinDETA")
	  WinWaitActive("WinDETA")
	  ; Run the light impedance
	  WinWait("Multi Graphics")
	  WinActivate("WinDETA")
	  WinWaitActive("WinDETA")
	  Send("!mt")
	  WinWaitActive("Warning")
	  ControlClick("Warning", "", "[CLASS:Button; INSTANCE:1]")
	  ; Detect when the measurement is finished
	  Sleep(5000) ; wait 5 sec
	  AutoItSetOption("WinTitleMatchMode", 3) ; Set option that window name must match exactly
	  WinWait("Online")
	  AutoItSetOption("WinTitleMatchMode", 1) ; Reset to partial match
	  ; Save the light impedance data
	  WinActivate("WinDETA")
	  WinWaitActive("WinDETA")
	  Send("!fa")
	  WinWaitActive("Save Result File")
	  ; Click OK on the save settings
	  ControlClick("Save Result File as ASCII", "", "[CLASS:Button; INSTANCE:9]")
	  WinWaitActive("Save Result File")
	  If $i == 0 Then
		 ; Navigate to device data folder
		 Sleep(1000)
		 ControlSetText("Save Result File", "", "[CLASS:Edit; INSTANCE:1]", "Data")
		 ControlClick("Save Result File", "", "[CLASS:Button; INSTANCE:2]")
		 Sleep(1000)
		 ControlSetText("Save Result File", "", "[CLASS:Edit; INSTANCE:1]", $device_name)
		 ControlClick("Save Result File", "", "[CLASS:Button; INSTANCE:2]")
		 Sleep(1000)
	  EndIf
	  ; Check for previous measurements and avoid name conflict
	  Local $N_measurement = 1
	  While FileExists($data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_LED"&$LED_array[$i]&"mA_"&$N_measurement&".txt")
		 $N_measurement += 1
	  WEnd
	  If $LED_array[$i] == 0 Then
		 ControlSetText("Save Result File", "", "[CLASS:Edit; INSTANCE:1]", $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_LEDdark_"&$N_measurement&".txt")
	  Else
		 ControlSetText("Save Result File", "", "[CLASS:Edit; INSTANCE:1]", $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_LED"&$LED_array[$i]&"mA_"&$N_measurement&".txt")
	  EndIf
	  ControlClick("Save Result File", "", "[CLASS:Button; INSTANCE:2]")
Next

; Turn off the LED
; Activate the LED driver control panel window
If WinActivate("Led Driver Control Panel") Then
   WinWaitActive("Led Driver Control Panel")
   ; Adjust the LED light intensity by setting the "Set Current" to $LED_array[$i]
   ControlSetText("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]", "0")
   ControlClick("Led Driver Control Panel", "", "[Class:TRzEdit; INSTANCE:2]")
   Send("{ENTER}")
   ; Click the LED Set Parameters button
   ControlClick("Led Driver Control Panel", "", "[CLASS:TRzBitBtn; INSTANCE:6]")
EndIf
