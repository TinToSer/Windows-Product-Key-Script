strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM SoftwareLicensingService",,48) 
For Each objItem in colItems 

    Wscript.Echo "OA3xOriginalProductKey: " & objItem.OA3xOriginalProductKey
Next