function Get-TogglDetailedReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)][Alias('Since')]
        [datetime]$From = (Get-Date),
        [Parameter(Position=1)][Alias('Until')]
        [datetime]$To = (Get-Date),
        [string]$Client,
        [string]$Project,
        [string]$UserAgent = $(if (Test-Path $PSScriptRoot\user_agent) { Get-Content $PSScriptRoot\user_agent } else { Set-Content -Value (Read-Host -Prompt 'Email Address') -Path $PSScriptRoot\user_agent -PassThru | Out-String }),
        [string]$User = $(if (Test-Path $PSScriptRoot\api_token) { Get-Content $PSScriptRoot\api_token } else { $User = Set-Content -Value (Read-Host -Prompt 'API Token') -Path $PSScriptRoot\api_token -PassThru | Out-String }),
        [string]$WorkspaceID = $(if (Test-Path $PSScriptRoot\workspace_id) { Get-Content $PSScriptRoot\workspace_id } else { Set-Content -Value (Read-Host -Prompt 'Workspace ID') -Path $PSScriptRoot\workspace_id -PassThru | Out-String })
    )

    $pass = "api_token"
    $pair = "$($User):$($pass)"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $basic_auth_value = "Basic $base64"
    $headers = @{Authorization=$basic_auth_value}
    $content_type = "application/json"
    $since = $From.ToString('yyyy-MM-dd')
    $until = $To.ToString('yyyy-MM-dd')
    
    # Query Toggl API for report details
    $uri_report = "https://toggl.com/reports/api/v2/details?since=$since&until=$until&display_hours=decimal&rounding=on&user_agent=$UserAgent&workspace_id=$WorkspaceID"
    $toggl_response = Invoke-RestMethod -Uri $uri_report -Headers $headers -ContentType $content_type
    $response_total = $toggl_response.total_count
    $page_num = 1
    $report = @()

    while ($response_total -gt 0) {
        $toggl_response = Invoke-RestMethod -Uri ($uri_report + '&page=' + $page_num) -Headers $headers -ContentType $content_type
        $TogglResponseData = $toggl_response.data
        $report += $TogglResponseData
        $response_total = $response_total - $toggl_response.per_page

        $page_num++
    }
    
    $report = $report | Select-Object @{n='Date';e={Get-Date $_.start -Format MM/dd/yyyy}},
                                      @{n='Client';e={$_.client}},
                                      @{n='Project';e={$_.project -replace '.*?(\b\d+\b).*','$1'}},
                                      @{n='Description';e={$_.description}},
                                      @{n='Hours';e={'{0:n2}' -f ($_.dur/1000/60/60)}},
                                      @{n='WorkType';e={$_.tags -as [string]}} | Sort-Object Date

    if ($Client) {
        $report = $report.Where{$_.Client -match $Client}
    }
    if ($Project) {
        $report = $report.Where{$_.Project -match $Project}
    }
    
    Write-Output $report
}


function Get-TogglUtilizationReport {
    # TODO Add the ability to run a report for multiple pay periods
    
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)]
        [Alias('Day')]
        [datetime]$Date = $(if ((Get-Date).Hour -lt 17) { (Get-Date).AddDays(-1) } else { Get-Date }),
        [string]$UserAgent = $(if (Test-Path $PSScriptRoot\user_agent) { Get-Content $PSScriptRoot\user_agent } else { Set-Content -Value (Read-Host -Prompt 'Email Address') -Path $PSScriptRoot\user_agent -PassThru | Out-String }),
        [string]$User = $(if (Test-Path $PSScriptRoot\api_token) { Get-Content $PSScriptRoot\api_token } else { $User = Set-Content -Value (Read-Host -Prompt 'API Token') -Path $PSScriptRoot\api_token -PassThru | Out-String }),
        [string]$WorkspaceID = $(if (Test-Path $PSScriptRoot\workspace_id) { Get-Content $PSScriptRoot\workspace_id } else { Set-Content -Value (Read-Host -Prompt 'Workspace ID') -Path $PSScriptRoot\workspace_id -PassThru | Out-String })
    )

    $CurrentDay = $Date.Day
    $CurrentMonth = $Date.Month
    $CurrentYear = $Date.Year

    # Calculate normal working hours for the specified pay period
    if ($CurrentDay -ge 1 -and $CurrentDay -le 15) {
        $PayPeriodStartDay = 1
        $Since = "$CurrentYear-$CurrentMonth-01"
        $Until = "$CurrentYear-$CurrentMonth-$CurrentDay"
        $NumberOfWorkDays = 0

        for ($x = $PayPeriodStartDay; $x -le $CurrentDay; $x += 1) {
            switch ((get-date $CurrentYear-$CurrentMonth-$x).DayOfWeek) {
                'Monday' { $NumberOfWorkDays += 1 }
                'Tuesday' { $NumberOfWorkDays += 1 }
                'Wednesday' { $NumberOfWorkDays += 1 }
                'Thursday' { $NumberOfWorkDays += 1 }
                'Friday' { $NumberOfWorkDays += 1 }
            }
        }
        
        $WorkingHours = $NumberOfWorkDays * 8
    } elseif ($CurrentDay -ge 16 -and $CurrentDay -le 31) {
        $PayPeriodStartDay = 16
        $Since = "$CurrentYear-$CurrentMonth-16"
        $Until = "$CurrentYear-$CurrentMonth-$CurrentDay"
        $NumberOfWorkDays = 0

        for ($x = $PayPeriodStartDay; $x -le $CurrentDay; $x += 1) {
            switch ((get-date $CurrentYear-$CurrentMonth-$x).DayOfWeek) {
                'Monday' { $NumberOfWorkDays += 1 }
                'Tuesday' { $NumberOfWorkDays += 1 }
                'Wednesday' { $NumberOfWorkDays += 1 }
                'Thursday' { $NumberOfWorkDays += 1 }
                'Friday' { $NumberOfWorkDays += 1 }
            }
        }
        
        $WorkingHours = $NumberOfWorkDays * 8
    }

    $detailed_report = Get-TogglDetailedReport -From $Since -To $Until

    $TotalHours = ($detailed_report | Measure-Object -Property Hours -Sum).Sum
    $BillableHours = ($detailed_report.Where{$_.WorkType -eq 'Billable'} | Measure-Object -Property Hours -Sum).Sum
    if ($BillableHours -eq $null) {
        $BillableHours = 0
    }
    $UtilizedHours = ($detailed_report.Where{$_.WorkType -eq 'Utilized'} | Measure-Object -Property Hours -Sum).Sum
    if ($UtilizedHours -eq $null) {
        $UtilizedHours = 0
    }
    $PtoHours = ($detailed_report.Where{$_.WorkType -eq 'PTO'} | Measure-Object -Property Hours -Sum).Sum
    if ($PtoHours -eq $null) {
        $PtoHours = 0
    }
    $HolidayHours = ($detailed_report.Where{$_.WorkType -eq 'Holiday'} | Measure-Object -Property Hours -Sum).Sum
    if ($HolidayHours -eq $null) {
        $HolidayHours = 0
    }
    $NonBillableHours = ($detailed_report.Where{$_.WorkType -eq 'Non-Billable'} | Measure-Object -Property Hours -Sum).Sum
    if ($NonBillableHours -eq $null) {
        $NonBillableHours = 0
    }
    $OvertimeHours = $TotalHours-$WorkingHours
    if ($OvertimeHours -lt 0) {
        $OvertimeHours = 0
    }
    $BillablePercent = ($BillableHours/($WorkingHours-$PtoHours-$HolidayHours))*100
    $UtilizedPercent = (($BillableHours+$UtilizedHours)/($WorkingHours-$PtoHours-$HolidayHours))*100
    
    # Output summary totals
    $obj = New-Object -TypeName PSObject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Normal -Value ('{0:N2}' -f $WorkingHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Overtime -Value ('{0:N2}' -f $OvertimeHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name PTO -Value ('{0:N2}' -f $PTOHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Holiday -Value ('{0:N2}' -f $HolidayHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Non-Billable -Value ('{0:N2}' -f $NonBillableHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Utilized -Value ('{0:N2}' -f $UtilizedHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Billable -Value ('{0:N2}' -f $BillableHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name TotalHours -Value ('{0:N2}' -f $TotalHours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name BillablePercent -Value ('{0:N0}' -f $BillablePercent)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name UtilizedPercent -Value ('{0:N0}' -f $UtilizedPercent)
    
    Write-Output $obj
}
