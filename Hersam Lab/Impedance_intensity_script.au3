; Measurement script by Michael C. Heiber, Vinod Sangwan, and Sam Amsterdam
; Who is doing the measurement?
Local $measurement_persons = "Sam Amsterdam"
; Sample Info
Local $device_name = "WY6_1A"
; Define the filters to be tested
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
Local $measurement_freq = 100 ; Hz
; AC amplitude
Local $ac_amplitude = 20 ; mV
; Define number of cycles for averaging
Local $N_cycles = 10 ; s

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
Local $start_voltage = -2.0 ; V
; End voltage
Local $end_voltage = 1.0 ; V
; Define voltage step size
Local $step_size = 50 ; mV


; Setup data folder
If Not FileExists("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data") Then
   DirCreate("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data")
EndIf
; Run the filter wheel control software if it's not already running
If Not WinActivate("COM3 - PuTTY") Then
   run("C:\Program Files\PuTTY\putty.exe")
   WinWaitActive("PuTTY Configuration")
   Send("{TAB}{TAB}{TAB}{TAB}{DOWN}{DOWN}!l{Enter}")
   WinWaitActive("COM3 - PuTTY")
   Send("{ENTER}")
   ; set filter to dark condition
   Send("pos=12{ENTER}")
   Sleep(2000)
EndIf

; Run Zplot if it's not already running
If Not WinActivate("ZPlot)") Then
   run("C:\SAI\Programs\ZPlot.exe")
   WinWaitActive("ZPlot")
   ; open measurement settings
   Send("!si")
   WinWaitActive("Setup Instruments  (Solartron 1260)")
   ; change to Analyzer tab
   Send("{RIGHT}")
   ; click cycles radio button
   ControlClick("Setup Instruments  (Solartron 1260)", "", "[CLASS:TRadioButton; INSTANCE:3]")
   ; set N_cycles
   Send("{TAB}"&$N_cycles)
   Send("{ENTER}")
EndIf

; -- Phase 1 -- dark impedance frequency scan
; Activate the Zplot window
   WinActivate("ZPlot")
   WinWaitActive("ZPlot")
   ; Set the dark C-f measurement conditions
   Local $window_handle = WinWaitActive("ZPlot")
   ; set dc bias
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit5]", $dc_bias)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit5]")
   ; set ac amplitude
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit4]", $ac_amplitude)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit4]")
   ; set frequency range
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit3]", $start_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit3]")
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit2]", $end_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit2]")
   ; set points per decade
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit1]", $N_ppd)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit1]")
   ; Run the dark impedance measurement
   Send("^s")
   ; check for measurement completion
   WinWaitActive("Save Data As...")
   ; save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   While FileExists("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".z")
	  $N_measurement += 1
   WEnd
   ; save new data file
   Send("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCf_dark_"&$N_measurement&".z")
   Send("{ENTER}")

; -- Phase 2 -- light impedance bias scan at each intensity
; Activate the ZPlot window
   WinActivate("ZPlot")
   Local $window_handle = WinWaitActive("ZPlot")
   ; switch tab to Control E: Sweep DC
   Send("{RIGHT}")
   Sleep(500)
   ; Set the light C-V settings
   ; set frequency
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit10]", $measurement_freq)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit10]")
   ; set start voltage
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit8]", $start_voltage)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit8]")
   ; set end voltage
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit7]", $end_voltage)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit7]")
   ; set sweep segments to 1
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit4]", "1")
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit4]")
   ; set sweep step size
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit6]", $step_size)
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit6]")
   ; set data rate to 1
   ControlSetText($window_handle, "", "[CLASSNN:ValidateEdit5]", "1")
   ControlClick($window_handle, "", "[CLASSNN:ValidateEdit5]")

; Loop through the filter positions array and measure at each light intensity
Local $filters[] = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 1, 1.3, 2, 3, 4]

For $i = 0 To UBound($filter_positions)-1

   ; Activate the filter wheel putty control
   WinActivate("COM3 - PuTTY")
   WinWaitActive("COM3 - PuTTY")
   ; set the filter position
   Send("pos="&$filter_positions[$i])
   Send("{ENTER}")
   Sleep(5000)

   ; Activate the Zplot window
   WinActivate("ZPlot")
   WinWaitActive("ZPlot")
   ; Run the light impedance
   Send("^s")
   ; check for measurement completion
   WinWaitActive("Save Data As...")
   ; save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   Local $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_"&$filters[$filter_positions[$i]-1]&"OD"
   If $filter_positions[$i] = 12 Then
	  $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_dark"
   EndIf
   If $filter_positions[$i] = 1 Then
	  $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" Impedance Data\"&$device_name&"_ImpedanceCV_1sun"
   EndIf
   While FileExists($partial_path&"_"&$N_measurement&".z")
	  $N_measurement += 1
   WEnd
   ; save new data file
   Send($partial_path&"_"&$N_measurement&".z")
   Send("{ENTER}")
Next


; Turn off light
; Activate the filter wheel putty control
WinActivate("COM3 - PuTTY")
WinWaitActive("COM3 - PuTTY")
; set the filter position
Send("pos=12{ENTER}")

; Plot the results
