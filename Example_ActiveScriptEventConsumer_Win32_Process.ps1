##########################################################################
###  Using WMIClass accelerator to bind WMI Filter and Consumer Events  ##
##########################################################################
$CIMClass = 'Win32_Process'
$ProcessName = 'Calculator.exe'

# Creating a new event filter
$instanceFilter = ([wmiclass]"\\.\root\subscription:__EventFilter").CreateInstance()
$instanceFilter.QueryLanguage = 'WQL'
$instanceFilter.Query = "SELECT * FROM __InstanceCreationEvent WITHIN .1 WHERE targetInstance ISA '$CIMClass' AND TargetInstance.Name='$ProcessName'"
$instanceFilter.Name = "Test_$CIMClass`_Process_Creation_Filter"
$instanceFilter.EventNamespace = 'root\cimv2'
$result = $instanceFilter.Put()
$newFilter = $result.Path

# Creating a new event consumer
$instanceConsumer = ([wmiclass]"\\.\root\subscription:ActiveScriptEventConsumer").CreateInstance()
$instanceConsumer.Name = "Test_$CIMClass`_Process_Creation_Consumer"
$instanceConsumer.ScriptingEngine = 'VBScript'
$scriptText = @"
Set objShell = CreateObject("Wscript.shell")
sTargetInstance = TargetEvent.TargetInstance.Name
objShell.run("powershell.exe -ExecutionPolicy Bypass -Command `$null = Get-WmiObject -Q \""SELECT * FROM $CIMClass WHERE name='" & sTargetInstance & "'\"" | Invoke-WmiMethod -Name Terminate -verbose")
"@
$instanceConsumer.ScriptText = $scriptText
$result = $instanceConsumer.Put()
$newConsumer = $result.Path

# Bind filter and consumer
$instanceBinding = ([wmiclass]"\\.\root\subscription:__FilterToConsumerBinding").CreateInstance()
$instanceBinding.Filter = $newFilter
$instanceBinding.Consumer = $newConsumer
$result = $instanceBinding.Put()
$newBinding = $result.Path

## Removing WMI Subscriptions using [wmi] and Delete() Method
# ([wmi]$newFilter).Delete()
# ([wmi]$newConsumer).Delete()
# ([wmi]$newBinding).Delete()
