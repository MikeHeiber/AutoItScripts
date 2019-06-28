; Copyright (c) 2019 Michael C. Heiber
; This script file is part of the AutoItScripts repository, which is subject to the MIT License.
; For more information, see the LICENSE file that accompanies this repository.
; The AutoItScripts repository can be found on Github at https://github.com/MikeHeiber/AutoItScripts
;
; Measurement script by Michael C. Heiber
; with contributions from Vinod Sangwan and Sam Amsterdam
;
; Who is doing the measurement?
Local $measurement_persons = "Sam Amsterdam"
; Define the sample name
Local $device_name = "WY6_3D"
; Define the device area
Local $device_area = 0.06 ; cm^2
; Define the solar simulator mismatch factor
Local $mismatch_factor = 0.87
; Define the data directory (Where to save the data?)
Local $data_dir = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"
; Define the filters to be used in the test
Local $filter_positions[] = [12, 8, 7, 6, 5, 4, 3, 2, 1] ; OD
; Filter table
; pos OD
; 1   0.0
; 2   0.1
; 3   0.2
; 4   0.3
; 5   0.4
; 6   0.5
; 7   0.6
; 8   1.0
; 9   1.3
; 10  2.0
; 11  3.0
; 12  dark

; JV measurment settings
Local $start_voltage = -2.0 ; V
Local $end_voltage = 1.0 ; V
Local $step_size = 0.01 ; V

; ===================================================================================================

; AutoIt Settings
; Partial window title matching
AutoItSetOption("WinTitleMatchMode",2)

; Setup data folder
If Not FileExists($data_dir&$device_name&" JV Data") Then
   DirCreate($data_dir&$device_name&" JV Data")
EndIf

; Run the filter wheel control software if it's not already running
If Not WinActivate("PuTTY") Then
	; Open Putty
   run("C:\Program Files\PuTTY\putty.exe")
   WinWaitActive("PuTTY Configuration")
   ; Connect to the filter wheel
   Send("{TAB 4}{DOWN 2}!l{Enter}")
   WinWaitActive("PuTTY")
   Send("{ENTER}")
   ; Set filter to dark condition
   Send("pos=12{ENTER}")
   ; Wait for 4 seconds for the filter wheel to move
   Sleep(4000)
EndIf

; Run Labview JV measurement Tool
If Not WinActivate("KE2400") Then
   run("C:\Program Files\National Instruments\LabVIEW 2019\LabVIEW.exe")
   Local $window_handle = WinWaitActive("LabVIEW")
   Send("^o")
   WinWaitActive("Select a File to Open")
   Send("C:\Users\Impedance Users\Desktop\IPDA Measurements\IV\KE2400 IV MEASUREMENT.vi{ENTER}")
EndIf

; Set JV Measurement settings
WinWaitActive("KE2400")
Send("{TAB}")
; Set start voltage
Send($start_voltage&"{TAB}")
; Set end voltage
Send($end_voltage&"{TAB}")
; Set compliance
Send("1e-2{TAB}")
; Set measurement speed
Send("{DOWN}{TAB 6}")
; Set number of steps
Local $N_steps = ($end_voltage-$start_voltage)/$step_size
Send($N_steps&"{TAB 3}")
; Turn off bi-polar measurement
Send("{ENTER}")

; Perform intensity dependent JV measurements
; Loop through the filter positions array and measure at each light intensity
Local $filters[] = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 1.0, 1.3, 2.0, 3.0, 10]

For $i = 0 To UBound($filter_positions)-1

   ; Activate the putty filter wheel control
   WinActivate("PuTTY")
   WinWaitActive("PuTTY")
   ; Set the filter position
   Send("pos="&$filter_positions[$i])
   Send("{ENTER}")
   ; Wait for 4 seconds for the filter wheel to move
   Sleep(4000)

   ; Activate the Labview window
   WinActivate("KE2400")
   Local $window_handle = WinWaitActive("KE2400")
   ; Run the JV test
   Send("{TAB 5}{ENTER}")
   ; Check for measurement completion
   Local $color = ""
   While $color <> "FF1612"
	  Sleep(2000)
	  $color = Hex(PixelGetColor(420, 320, $window_handle),6)
   WEnd

   ; Save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   Local $partial_path = ""
   If $filter_positions[$i] = 1 Then
	  $partial_path = $data_dir&$device_name&" JV Data\"&$device_name&"_JV_1sun"
   ElseIf $filter_positions[$i] = 12 Then
	  $partial_path = $data_dir&$device_name&" JV Data\"&$device_name&"_JV_dark"
   Else
	  $partial_path = $data_dir&$device_name&" JV Data\"&$device_name&"_JV_"&$filters[$filter_positions[$i]-1]&"OD"
   EndIf
   While FileExists($partial_path&"_"&$N_measurement&".txt")
	  $N_measurement += 1
   WEnd
   ; Save new data file
   Send("{TAB 2}{ENTER}")
   WinWaitActive("Save file...")
   Send($partial_path&"_"&$N_measurement&".txt{ENTER}")
Next

; Close the aperture
; Activate the filter wheel putty control
WinActivate("PuTTY")
WinWaitActive("PuTTY")
; Set the filter position
Send("pos=12{ENTER}")
; Close Putty
WinClose("PuTTY")
WinWaitActive("PuTTY Exit")
Send("{ENTER}")

; Close the Labview IV program
WinActivate("KE2400")
WinWaitActive("KE2400")
WinClose("KE2400")
WinWaitActive("Save changes")
Send("{TAB}{ENTER}")

; Plot the results
; Open Igor Pro
If Not WinActivate("Igor Pro") Then
	run("C:\Program Files (x86)\WaveMetrics\Igor Pro Folder\Igor.exe")
EndIf
WinWaitActive("Igor Pro")
; Load the JV data
; Activate the command window
Send("^j")
; Wait a half second for the window to activate
Sleep(500)
; Execute the load command to import the JV data into Igor
Local $JV_dir = $data_dir&$device_name&" JV Data\"
$JV_dir = StringReplace($JV_dir,"\",":")
$JV_dir = StringReplace($JV_dir,"::",":")
Send('FEDMS_LoadJVFolder("'&$JV_dir&'","'&$measurement_persons&'","","Hersam Lab",'&$device_area&','&$mismatch_factor&'){ENTER}')
; Wait 2 sec before executing the next command
Sleep(2000)
; Analyze the JV intensity data
Send('FEDMS_AnalyzeJVIntensity("'&$device_name&'"){ENTER}'
; Wait 2 sec before executing the next command
Sleep(2000)
; Plot the photocurrent data for all intensities
Send('FEDMS_PlotPhotocurrentData("'&$device_name&'"){ENTER}'
