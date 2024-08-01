# Define the commands to be run at logon
$actionCommand = 'powershell.exe -Command "net use z: /delete; net use z: \\airdrie-dc\company\labreports /user:dtsttk\laba labairdrie /persist:yes"'

# Create a new action for the scheduled task
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-Command `"$actionCommand`""

# Define the trigger to run the task at logon of any user
$trigger = New-ScheduledTaskTrigger -AtLogOn

# Define the principal for the task to run with highest privileges
$principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Users" -RunLevel Highest

# Define the settings for the scheduled task
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register the scheduled task
Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -TaskName "MapNetworkDrive" -Description "Maps the Z: drive to the specified network share at user logon"
