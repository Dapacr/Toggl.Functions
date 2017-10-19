function Get-TogglDetailedReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)][Alias('Since')]
        [datetime]$From = (Get-Date),
        [Parameter(Position=1)][Alias('Until')]
        [datetime]$To = (Get-Date),
        [string]$Client,
        [string]$Project,
        [string]$Description,
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
                                      @{n='Ticket';e={$_.project -replace '.*?(\b\d+\b).*','$1'}},
                                      @{n='Description';e={$_.description}},
                                      @{n='Project';e={$_.project -replace '^Ticket\s#\s\d+\s\((.+)\)','$1'}},
                                      @{n='Hours';e={'{0:n2}' -f ($_.dur/1000/60/60)}},
                                      @{n='WorkType';e={$_.tags -as [string]}} | Sort-Object Date

    if ($Client) {
        $report = $report.Where{$_.Client -match $Client}
    }
    if ($Project) {
        $report = $report.Where{$_.Project -match $Project}
    }
    if ($Description) {
        $report = $report.Where{$_.Description -match $Description}
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

    
    # Calculate normal working hours for the specified pay period
    $normal_hours = 0
    $work_days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
    if ($Date.Day -ge 1 -and $Date.Day -le 15) {
        $period_start = 1
        $from = '{0:yyyy-MM}-01' -f $Date
        $to = '{0:yyyy-MM-dd}' -f $Date

        for ($x=$period_start; $x -le $Date.Day; $x++) {
            $day = '{0:yyyy-MM}-{1}' -f $Date, $x
            $day_of_week = (Get-Date $day).DayOfWeek
            
            if ($day_of_week -in $work_days) {
                $normal_hours += 8
            }
        }
    } elseif ($Date.Day -ge 16 -and $Date.Day -le 31) {
        $period_start = 16
        $from = '{0:yyyy-MM}-16' -f $Date
        $to = '{0:yyyy-MM-dd}' -f $Date

        for ($x=$period_start; $x -le $Date.Day; $x++) {
            $day = '{0:yyyy-MM}-{1}' -f $Date, $x
            $day_of_week = (Get-Date $day).DayOfWeek
            
            if ($day_of_week -in $work_days) {
                $normal_hours += 8
            }
        }
    }
    
    $detailed_report = Get-TogglDetailedReport -From $from -To $to

    $total_hours = ($detailed_report | Measure-Object -Property Hours -Sum).Sum
    $billable_hours = ($detailed_report.Where{$_.WorkType -eq 'Billable'} | Measure-Object -Property Hours -Sum).Sum
    if ($billable_hours -eq $null) {
        $billable_hours = 0
    }
    $utilized_hours = ($detailed_report.Where{$_.WorkType -eq 'Utilized'} | Measure-Object -Property Hours -Sum).Sum
    if ($utilized_hours -eq $null) {
        $utilized_hours = 0
    }
    $pto_hours = ($detailed_report.Where{$_.WorkType -eq 'PTO'} | Measure-Object -Property Hours -Sum).Sum
    if ($pto_hours -eq $null) {
        $pto_hours = 0
    }
    $holiday_hours = ($detailed_report.Where{$_.WorkType -eq 'Holiday'} | Measure-Object -Property Hours -Sum).Sum
    if ($holiday_hours -eq $null) {
        $holiday_hours = 0
    }
    $non_billable_hours = ($detailed_report.Where{$_.WorkType -eq 'Non-Billable'} | Measure-Object -Property Hours -Sum).Sum
    if ($non_billable_hours -eq $null) {
        $non_billable_hours = 0
    }
    $overtime_hours = $total_hours-$normal_hours
    if ($overtime_hours -lt 0) {
        $overtime_hours = 0
    }
    $percent_billable = ($billable_hours/($normal_hours-$pto_hours-$holiday_hours))*100
    $percent_utilized = (($billable_hours+$utilized_hours)/($normal_hours-$pto_hours-$holiday_hours))*100
    
    # Output summary totals
    $obj = New-Object -TypeName PSObject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name TotalHours -Value ('{0:N2}' -f $total_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Normal -Value ('{0:N2}' -f $normal_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Overtime -Value ('{0:N2}' -f $overtime_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name PTO -Value ('{0:N2}' -f $pto_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Holiday -Value ('{0:N2}' -f $holiday_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Non-Billable -Value ('{0:N2}' -f $non_billable_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Utilized -Value ('{0:N2}' -f $utilized_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Billable -Value ('{0:N2}' -f $billable_hours)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name BillablePercent -Value ('{0:N0}' -f $percent_billable)
    Add-Member -InputObject $obj -MemberType NoteProperty -Name UtilizedPercent -Value ('{0:N0}' -f $percent_utilized)
    
    Write-Output $obj
}
