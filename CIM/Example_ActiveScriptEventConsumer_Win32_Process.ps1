<#
# Based on code from 
# https://adamtheautomator.com/your-goto-guide-for-working-with-windows-wmi-events-and-powershell/#Binding_the_Event_Filter_and_Consumer_Together
#>
# Turn into Parameters
$eventClass = '__InstanceCreationEvent'
$timespan = '.1'
$targetClass = 'Win32_Process'
$propertyName = 'Name'
$value = 'Calculator.exe'
$PathToApplication = "$PSHOME\powershell.exe"
$PathToScript = 'C:\ExampleScript.ps1'
$parameter = "-ProcessName"

## Create Event filter ##
# Create unique name
$filterName = "My_$targetClass`_Filter"

# Construct the WQL Query
$FilterQuery = "SELECT * FROM $EventClass WITHIN $timespan WHERE targetInstance ISA '$targetClass' AND TargetInstance.$propertyName='$value'"

# Hashtable for the property argument
$property = @{
  Name = $filterName
  EventNameSpace = "Root/CIMV2"
  QueryLanguage = "WQL"
  Query = $FilterQuery
}

# Hashtable for splatting
$splat = @{
    ClassName = '__EventFilter'
    Namespace = "Root/SubScription"
    Property = $property
}
$CIMFilterInstance = New-CimInstance @splat

# Confirm Event Filter
$splat.Remove('Property')
Get-CimInstance @splat

## Create CommandLineEventConsumer ##
# Create unique name
$consumerName = "My_$targetClass`_AppConsumer"

# Construct command to invoke on trigger
$commandLineTemplate = "$($PathToApplication -replace '\\','\\') $($PathToScript -replace '\\','\\') $Parameter %TargetInstance.$propertyName%"

# Hashtable for property argument
$property = @{
  Name = $consumerName
  CommandLineTemplate = $commandLineTemplate
}

# Hashtable for splatting
$splat = @{
ClassName = 'CommandLineEventConsumer'
NameSpace = 'ROOT/subscription'
Property  = $property
}
$CIMEventConsumer = New-CimInstance @splat

# Confirm Event Consumer
$splat.Remove('Property')
Get-CimInstance @splat

# Bind filter and consumer
# Hashtable for Property argument
$property = @{
  Filter   = [ref]$CIMFilterInstance
  Consumer = [ref]$CIMEventConsumer
}

# Hashtable for splatting
$splat = @{
ClassName = '__FilterToConsumerBinding'
Namespace = "root/subscription"
Property  = $property
}
$CIMBinding = New-CimInstance @splat

# Function to easily remove the event listeners
function Remove-EventListener {
# Remove EventFilter
Get-CimInstance -Namespace Root/Subscription -ClassName __EventFilter | Where-Object {$_.name -eq $filterName} | Remove-CimInstance

# Remove EventConsumer
Get-CimInstance -Namespace Root/Subscription -ClassName CommandLineEventConsumer | Where-Object {$_.Name -eq $consumerName} | Remove-CimInstance

# Remove Binding
Get-CimInstance -Namespace Root/Subscription -ClassName __FilterToConsumerBinding | Where-Object {$_.Filter.Name -eq $filterName} | Remove-CimInstance
}
