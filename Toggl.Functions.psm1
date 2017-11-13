<#
        .Synopsis
        Returns report data for the specified date range.
        .Description
        This cmdlet uses Toggl's tag functionality to classify time entries.
        
        The following tags are used:
        
            Billable
            Holiday
            Non-Billable
            PTO
            Training
            Utilized

        .Link
        https://github.com/Dapacruz/VMware.VimAutomation.Custom
#>
function Get-TogglDetailedReport {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory, Position=0)][Alias('Since')]
        [datetime]$From,
        [Parameter(Position=1)][Alias('Until')]
        [datetime]$To = (Get-Date),
        [string]$Client,
        [string]$Ticket,
        [string]$Project,
        [string]$Description,
        [string]$WorkType,
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
    
    $report = $report | Select-Object @{n='Date';e={Get-Date $_.start -Format yyyy-MM-dd}},
                                      @{n='Client';e={$_.client}},
                                      @{n='Ticket';e={($_.project -replace '(.*?\b(\d+)\b.*|.*)','$2') -replace '$^','n/a'}},
                                      @{n='Description';e={$_.description}},
                                      @{n='Project';e={$_.project -replace '^Ticket\s#\s\d+\s\((.+)\)','$1'}},
                                      @{n='Duration(hrs)';e={'{0:n2}' -f ($_.dur/1000/60/60)}},
                                      @{n='WorkType';e={$_.tags -as [string]}} | Sort-Object Date
    
    if ($Client) {
        $report = $report.Where{$_.Client -match $Client}
    }
    if ($Ticket) {
        $report = $report.Where{$_.Ticket -match $Ticket}
    }
    if ($Project) {
        $report = $report.Where{$_.Project -match $Project}
    }
    if ($Description) {
        $report = $report.Where{$_.Description -match $Description}
    }
    if ($WorkType) {
        $report = $report.Where{$_.WorkType -match $WorkType}
    }

    # Insert custom TypeName (defined in $PSScriptRoot\format.ps1xml) to control default display
    foreach ($obj in $report) {
        $obj.PSTypeNames.Insert(0,'Toggl.Report.Detailed')
    }
        
    Write-Output $report
}


<#
        .Synopsis
        Calculate a utilization report for all pay periods in the specified date range.
        .Description
        This cmdlet uses Toggl's tag functionality to classify time entries.
        
        The following tags are used:
        
            Billable
            Holiday
            Non-Billable
            PTO
            Training
            Utilized

        .Link
        https://github.com/Dapacruz/VMware.VimAutomation.Custom
#>
function Get-TogglUtilizationReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)]
        [datetime]$From = $(if ((Get-Date).Hour -lt 17) { (Get-Date).AddDays(-1) } else { Get-Date }),
        [Parameter(Position=1)]
        [datetime]$To = (Get-Date),
        [switch]$ExcludeCurrentPeriod,
        [string]$UserAgent = $(if (Test-Path $PSScriptRoot\user_agent) { Get-Content $PSScriptRoot\user_agent } else { Set-Content -Value (Read-Host -Prompt 'Email Address') -Path $PSScriptRoot\user_agent -PassThru | Out-String }),
        [string]$User = $(if (Test-Path $PSScriptRoot\api_token) { Get-Content $PSScriptRoot\api_token } else { $User = Set-Content -Value (Read-Host -Prompt 'API Token') -Path $PSScriptRoot\api_token -PassThru | Out-String }),
        [string]$WorkspaceID = $(if (Test-Path $PSScriptRoot\workspace_id) { Get-Content $PSScriptRoot\workspace_id } else { Set-Content -Value (Read-Host -Prompt 'Workspace ID') -Path $PSScriptRoot\workspace_id -PassThru | Out-String })
    )

    $work_days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
    $report = @()
    
    # Normalize $From
    if ($From.Day -gt 1 -and $From.Day -le 15) {
        $From = Get-Date $From -Day 1
    } elseif ($From.Day -gt 16) {
        $From = Get-Date $From -Day 16
    }
    
    if ($ExcludeCurrentPeriod) {
        if ($To -ge (Get-Date -Day 16)) {
            $To = Get-Date -Day 15
        } elseif ($To -ge (Get-Date -Day 1)) {
            $To = (Get-Date -Day 1).AddDays(-1)
        }
    }
    
    while ($From -le $To) {
        $month_first_day = Get-Date $From -Day 1
        $month_last_day = $month_first_day.AddMonths(1).AddDays(-1)
        
        # Determine period start and end
        switch ($From.Day) {
            1 {
                $period_start = $From.Day
                if ($From.AddDays(14) -gt $To) {
                    $period_end = $To.Day
                } else {
                    $period_end = $From.AddDays(14).Day
                }
            }
            16 {
                $period_start = $From.Day
                if ($month_last_day -gt $To) {
                    $period_end = $To.Day
                } else {
                    $period_end = $month_last_day.Day
                }
            }
            default {throw "Current Date out of range: $From"}
        }
    
        # Calculate normal working hours for the specified pay period
        $normal_hours = 0
        for ($x=$period_start; $x -le $period_end; $x++) {
            $day = '{0:yyyy-MM}-{1}' -f $From, $x
            $day_of_week = (Get-Date $day).DayOfWeek
            
            if ($day_of_week -in $work_days) {
                $normal_hours += 8
            }
        }
        
        $detailed_report = Get-TogglDetailedReport -From ('{0:yyyy-MM}-{1}' -f $From, $period_start) -To ('{0:yyyy-MM}-{1}' -f $From, $period_end)

        $total_hours = ($detailed_report | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($total_hours -eq $null) {
            $total_hours = 0
        }
        $billable_hours = ($detailed_report.Where{$_.WorkType -eq 'Billable'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($billable_hours -eq $null) {
            $billable_hours = 0
        }
        $utilized_hours = ($detailed_report.Where{$_.WorkType -eq 'Utilized'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($utilized_hours -eq $null) {
            $utilized_hours = 0
        }
        $training_hours = ($detailed_report.Where{$_.WorkType -eq 'Training'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($training_hours -eq $null) {
            $training_hours = 0
        }
        $pto_hours = ($detailed_report.Where{$_.WorkType -eq 'PTO'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($pto_hours -eq $null) {
            $pto_hours = 0
        }
        $holiday_hours = ($detailed_report.Where{$_.WorkType -eq 'Holiday'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($holiday_hours -eq $null) {
            $holiday_hours = 0
        }
        $non_billable_hours = ($detailed_report.Where{$_.WorkType -eq 'Non-Billable'} | Measure-Object -Property 'Duration(Hrs)' -Sum).Sum
        if ($non_billable_hours -eq $null) {
            $non_billable_hours = 0
        }
        $overtime_hours = $total_hours-$normal_hours
        if ($overtime_hours -lt 0) {
            $overtime_hours = 0
        }
        if ($normal_hours -gt 0) {
            $percent_billable = ($billable_hours/($normal_hours-$pto_hours-$holiday_hours-$training_hours))*100
            $percent_utilized = (($billable_hours+$utilized_hours)/($normal_hours-$pto_hours-$holiday_hours-$training_hours))*100
        } else {
            $percent_billable = 0
            $percent_utilized = 0
        }
        # Output summary totals
        $obj = New-Object -TypeName PSObject
        
        # Insert custom TypeName (defined in $PSScriptRoot\format.ps1xml) to control default display
        $obj.PSTypeNames.Insert(0,'Toggl.Report.Utilization')
        
        Add-Member -InputObject $obj -MemberType NoteProperty -Name PeriodStart -Value ('{0:yyyy-MM-dd}' -f (Get-Date $From -Day $period_start))
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Total(hrs)' -Value ('{0:N2}' -f $total_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Normal(hrs)' -Value ('{0:N2}' -f $normal_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Overtime(hrs)' -Value ('{0:N2}' -f $overtime_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Training(hrs)' -Value ('{0:N2}' -f $training_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'PTO(hrs)' -Value ('{0:N2}' -f $pto_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Holiday(hrs)' -Value ('{0:N2}' -f $holiday_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Non-Billable(hrs)' -Value ('{0:N2}' -f $non_billable_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Utilized(hrs)' -Value ('{0:N2}' -f $utilized_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Billable(hrs)' -Value ('{0:N2}' -f $billable_hours)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Billable(%)' -Value ('{0:N0}' -f $percent_billable)
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Utilized(%)' -Value ('{0:N0}' -f $percent_utilized)

        $report += $obj
        
        switch ($From.Day) {
            1 {$From = $From.AddDays(15)}
            16 {$From = Get-Date $From.AddMonths(1) -Day 1}
            default {throw "Current date out of range: $From"}
        }
    }
    
    Write-Output $report
}


Update-FormatData -PrependPath $PSScriptRoot\format.ps1xml