On Error Resume Next

intStartingAddress = 1
intEndingAddress = 254
strSubnet = "192.168.1."

For i = intStartingAddress to intEndingAddress
    strComputer = strSubnet & i
    Wscript.Echo strComputer
Next
  


