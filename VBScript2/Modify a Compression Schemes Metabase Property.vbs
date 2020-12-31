' Description: Demonstration script that modifies a compression schemes property (HcSendCacheHeaders) in the IIS metabase.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & _
    "/W3SVC/Filters/Compression/Parameters")

objIIS.HcSendCacheHeaders = TRUE
objIIS.SetInfo

