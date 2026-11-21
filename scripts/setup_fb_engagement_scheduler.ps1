$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument '/c call "D:\AI AGENT THUY\AI branding\chay_tu_dong_fb_tuong_tac.bat"' -WorkingDirectory "D:\AI AGENT THUY\AI branding"
$trigger1 = New-ScheduledTaskTrigger -Daily -At "9:00AM"
$trigger2 = New-ScheduledTaskTrigger -Daily -At "5:00PM"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask -TaskName "Eureka_Facebook_Personal_Engagement_Daily" -Action $action -Trigger @($trigger1, $trigger2) -Settings $settings -Force
Write-Host "Task Eureka_Facebook_Personal_Engagement_Daily created successfully!"
