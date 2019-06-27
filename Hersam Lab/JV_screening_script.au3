; Measurement script by Michael C. Heiber, Vinod Sangwan, and Sam Amsterdam
;
;
; Who is doing the measurement?
Local $measurement_persons = "Sam Amsterdam"
; Sample Info
Local $device_name = "WY6_1C"
; Define the filters to be tested
Local $filter_positions[] = [12, 1] ; OD
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

; JV measurment settings
; start voltage
Local $start_voltage = -2.0 ; V
Local $end_voltage = 1.0 ; V
Local $step_size = 0.05 ; V

; Setup data folder
If Not FileExists("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" JV Data") Then
   DirCreate("C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" JV Data")
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

; Run Labview JV measurement Tool
If Not WinActivate("KE2400") Then
   run("C:\Program Files\National Instruments\LabVIEW 2019\LabVIEW.exe")
   Local $window_handle = WinWaitActive("LabVIEW")
   Send("^o")
   WinWaitActive("Select a File to Open")
   Send("C:\Users\Impedance Users\Desktop\IPDA Measurements\IV\KE2400 IV MEASUREMENT.vi{ENTER}")
EndIf

; Set JV Measurement settings
WinActivate("KE2400 IV MEASUREMENT.vi")
WinWaitActive("KE2400 IV MEASUREMENT.vi")
Send("{TAB}")
; set start voltage
Send($start_voltage&"{TAB}")
; set end voltage
Send($end_voltage&"{TAB}")
; set compliance
Send("1e-2{TAB}")
; set measurement speed
Send("{DOWN}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
; set number of steps
Local $N_steps = ($end_voltage-$start_voltage)/$step_size
Send($N_steps&"{TAB}{TAB}{TAB}")
; turn off bi-polar measurement
Send("{ENTER}")

; Perform intensity dependent JV measurements
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

   ; Activate the Labview window
   WinActivate("KE2400")
   Local $window_handle = WinWaitActive("KE2400")
   ; Run the JV test
   Send("{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}")
   ; check for measurement completion
   Sleep(2000)
   Local $color = Hex(PixelGetColor(420, 320, $window_handle),6)
   While $color <> "FF1612"
	  Sleep(2000)
	  $color = Hex(PixelGetColor(420, 320, $window_handle),6)
   WEnd

   ; save the data
   ; Check for previous measurements and avoid name conflict
   Local $N_measurement = 1
   Local $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" JV Data\"&$device_name&"_JVscreen_"&$filters[$filter_positions[$i]-1]&"OD"
   If $filter_positions[$i] = 1 Then
	  $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" JV Data\"&$device_name&"_JVscreen_1sun"
   EndIf
   If $filter_positions[$i] = 12 Then
	  $partial_path = "C:\Users\Impedance Users\Desktop\IPDA Measurements\Data\"&$device_name&" JV Data\"&$device_name&"_JVscreen_dark"
   EndIf
   While FileExists($partial_path&"_"&$N_measurement&".txt")
	  $N_measurement += 1
   WEnd
   ; save new data file
   Send("{TAB}{TAB}{ENTER}")
   WinWaitActive("Save file...")
   Send($partial_path&"_"&$N_measurement&".txt{ENTER}")
Next

; Turn off light
; Activate the filter wheel putty control
WinActivate("COM3 - PuTTY")
WinWaitActive("COM3 - PuTTY")
; set the filter position
Send("pos=12{ENTER}")
Sleep(2000)

; Plot the results
; Open Igor Pro
;WinActivate("Igor Pro")
;WinWaitActive("Igor Pro")
; Activate the command window
;Send("!J")
;Send("")