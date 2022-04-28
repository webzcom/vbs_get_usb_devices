'http://blogs.technet.com/b/heyscriptingguy/archive/2005/03/15/how-can-i-determine-which-usb-devices-are-connected-to-a-computer.aspx
'Win32_PnPEntity info here: http://msdn.microsoft.com/en-us/library/aa394353(v=vs.85).aspx 
strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDevices = objWMIService.ExecQuery _
    ("Select * From Win32_USBControllerDevice")

For Each objDevice in colDevices
    strDeviceName = objDevice.Dependent
    strQuotes = Chr(34)
    strDeviceName = Replace(strDeviceName, strQuotes, "")
    arrDeviceNames = Split(strDeviceName, "=")
    strDeviceName = arrDeviceNames(1)
    Set colUSBDevices = objWMIService.ExecQuery _
        ("Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'")
    For Each objUSBDevice in colUSBDevices
        Wscript.Echo objUSBDevice.Description & ": " & objUSBDevice.Manufacturer
		
    Next    
Next
