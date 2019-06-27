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
; Define the active layer thickness
Local $active_thickness = 100 ; nm
; Define the data directory (Where to save the data?)
Local $data_dir = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"
; Define the filters to be used in the test
Local $filter_positions[] = [12, 8, 7, 6, 5, 4, 3, 2, 1] ; OD

; Filter table
; pos OD
; 1   0
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
; 12  4.0

; Define the impedance bias scan measurement frequency
Local $measurement_freq = 1000 ; Hz
; AC amplitude
Local $ac_amplitude = 100 ; mV
; Define number of cycles for averaging
Local $N_cycles = 50 ; s

; Dark C-f scan settings
Local $dc_bias = -2.0 ; V
; Start frequency (100 recommended)
Local $start_freq = 100 ; Hz
; End frequency (1e6 recommended)
Local $end_freq = 5e6 ; Hz
; Number of points per decade
Local $N_ppd = 10

; C-V scan settings
; Start voltage
Local $start_voltage = -1.0 ; V
; End voltage
Local $end_voltage = 0.9 ; V
; Define voltage step size
Local $step_size = 25 ; mV

; ===================================================================================================

; AutoIt Settings
; Partial window title matching
AutoItSetOption("WinTitleMatchMode",2)

; Setup data folder
If Not FileExists($data_dir&$device_name&" Impedance Data") Then
   DirCreate($data_dir&$device_name&" Impedance Data")
EndIf
; Run the filter wheel control software if it's not already running
If Not WinActivate("PuTTY") Then
	; Open putty
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

; Run Zplot if it's not already running
If Not WinActivate("ZPlot)") Then
	; Open Zplot
   run("C:\SAI\Programs\ZPlot.exe")
   WinWaitActive("ZPlot")
   ; Open measurement settings
   Send("!si")
   WinWaitActive("Setup Instruments")
   ; Change to Analyzer tab
   Send("{RIGHT}")
   ; Click cycles radio button
   ControlClick("Setup Instruments", "", "[CLASS:TRadioButton; INSTANCE:3]")
   ; Set N_cycles
   Send("{TAB}"&$N_cycles&"{ENTER}")
EndIf

; -- Phase 1 -- dark impedance frequency scan
; Activate the Zplot window
   WinActivate("ZPlot")
   WinWaitActive("ZPlot")
   ; Set the dark C-f measurement conditions
   Local $window_handle = WinWaitActive("ZPlot")
   ; Set dc bias
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit5]", $dc_bias)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit5]")
   ; Set ac amplitude
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit4]", $ac_amplitude)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit4]")
   ; Set frequency range
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit3]", $start_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit3]")
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit2]", $end_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit2]")
   ; Set points per decade
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit1]", $N_ppd)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit1]")
   ; Run the dark impedance measurement
   Send("^s")
   ; Check for measurement completion
   WinWaitActive("Save Data As")
   ; Save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   While FileExists($data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".z")
	  $N_measurement += 1
   WEnd
   ; Save new data file
   Send($data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".z")
   Send("{ENTER}")

; -- Phase 2 -- light impedance bias scan at each intensity
; Activate the ZPlot window
   WinActivate("ZPlot")
   Local $window_handle = WinWaitActive("ZPlot")
   ; Switch tab to Control E: Sweep DC
   Send("{RIGHT}")
   Sleep(500)
   ; Set the light C-V settings
   ; Set frequency
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit10]", $measurement_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit10]")
   ; Set start voltage
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit8]", $start_voltage)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit8]")
   ; Set end voltage
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit7]", $end_voltage)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit7]")
   ; Set sweep segments to 1
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit4]", "1")
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit4]")
   ; Set sweep step size
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit6]", $step_size)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit6]")
   ; Set data rate to 1
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit5]", "1")
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit5]")

; Loop through the filter positions array and measure at each light intensity
Local $filters[] = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 1.0, 1.3, 2.0, 3.0, 10]

For $i = 0 To UBound($filter_positions)-1

   ; Activate the filter wheel putty control
   WinActivate("PuTTY")
   WinWaitActive("PuTTY")
   ; Set the filter position
   Send("pos="&$filter_positions[$i]&"{ENTER}")
   ; Wait for 4 seconds for filter wheel to move
   Sleep(4000)

   ; Activate the Zplot window
   WinActivate("ZPlot")
   WinWaitActive("ZPlot")
   ; Run the light impedance CV sweep
   Send("^s")
   ; Check for measurement completion
   WinWaitActive("Save Data As")
   ; Save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   Local $partial_path = ""
   If $filter_positions[$i] = 12 Then
	  $partial_path = $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_dark"
   ElseIf $filter_positions[$i] = 1 Then
	  $partial_path = $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_1sun"
   Else
	  $partial_path = $data_dir&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_"&$filters[$filter_positions[$i]-1]&"OD"
   EndIf
   While FileExists($partial_path&"_"&$N_measurement&".z")
	  $N_measurement += 1
   WEnd
   ; Save new data file
   Send($partial_path&"_"&$N_measurement&".z")
   Send("{ENTER}")
Next

; Turn off light
; Activate the filter wheel putty control
WinActivate("PuTTY")
WinWaitActive("PuTTY")
; Set the filter position
Send("pos=12{ENTER}")
; Close Putty
WinClose("PuTTY")
WinWaitActive("PuTTY Exit")
Send("{ENTER}")

; Close ZPlot
WinActivate("ZPlot")
WinWaitActive("ZPlot")
WinClose("ZPlot")
WinWaitActive("Information")
Send("{TAB}{ENTER}")

; Plot the results
