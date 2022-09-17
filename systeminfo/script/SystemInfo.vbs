' Startup Variables
computerName = ""
processorName = ""
processorSpeed = 0
memorySize = 0
hardDisk=0
tapeName = ""

'Code Start
strComputer = "."
SetNameAndMemory (strComputer)
processorName = GetProcessor (strComputer)
processorSpeed = GetAverageValue(Cint(GetProcessorSpeed (strComputer)))
hardDisk = GetHardDisk(strComputer)
strResult =             "##[  SISTEM BILGISI  ]#####################################" & vbcrlf
strResult = strResult &vbcrlf
strResult = strResult & "Bilgisayar Adi		#  " & computerName & vbcrlf 
strResult = strResult & "Islemci			#  " & processorName & vbcrlf 
strResult = strResult & "Hiz			#  " & processorSpeed & vbcrlf 
strResult = strResult & "Ram Boyutu		#  " & memorySize & " MB" & vbcrlf 
strResult = strResult & "Hard Disk Boyutu		#  " & hardDisk & " GB" & vbcrlf 
strResult = strResult & vbcrlf
strResult = strResult & "###################################################" & vbcrlf
msgbox(strResult)
'------------------------------------------
sub SetNameAndMemory(strComputer)
Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWmi.ExecQuery("Select * from Win32_ComputerSystem")
for each objComputer in colSettings
computerName = objComputer.Name
memorySize = Clng(objComputer.TotalPhysicalMemory/(1024*1024)) + 1
next
end sub
'------------------------------------------
function GetProcessor(strComputer)
Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWmi.ExecQuery("Select * from Win32_Processor")
strReturn = ""
for each objProcessor in colSettings
strReturn = objProcessor.Name 
next
GetProcessor = strReturn
end function
'------------------------------------------
function GetProcessorSpeed(strComputer)
Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWmi.ExecQuery("Select * from Win32_Processor")
strReturn = ""
for each objProcessor in colSettings
strReturn = objProcessor.MaxClockSpeed
next
GetProcessorSpeed = strReturn
end function
'------------------------------------------
function GetHardDisk(strComputer)
Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWmi.ExecQuery("Select * from Win32_LogicalDisk")
fullSize = 0
for each objLogicalDisk in colSettings
if (CInt(objLogicalDisk.DriveType)=3) then
	fullSize = fullSize + Clng(objLogicalDisk.Size/(1024*1024*1024))
end if
next
GetHardDisk = fullSize
end function
'------------------------------------------
function GetAverageValue(unCalculatedNumber)
returnValue=0
i=0
if (unCalculatedNumber<1000) then
	returnValue = unCalculatedNumber
else
i = unCalculatedNumber mod 100
unCalculatedNumber = unCalculatedNumber - i + 100
returnValue = unCalculatedNumber
end if
GetAverageValue = returnValue
end function
'------------------------------------------


