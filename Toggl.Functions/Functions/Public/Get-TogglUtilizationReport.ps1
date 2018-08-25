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
        .Example
        Get-TogglUtilizationReport
        Generate a utilization report for the current pay period
        .Example
        Get-TogglUtilizationReport -From (Get-Date).AddMonths(-12) -ExcludeCurrentPeriod | Export-Csv -NoTypeInformation 'c:\temp\TogglUtilizationReport.csv'
        Generate a utilization report for the past year, excluding the current pay period, and outputs to a CSV file
        .Example
        Get-TogglUtilizationReport -From (Get-Date).AddMonths(-12) -ExcludeCurrentPeriod | Out-GridView
        Generate a utilization report for the past year, excluding the current pay period, and outputs to grid view
        .Link
        https://github.com/Dapacruz/Toggl.Functions
#>
function Get-TogglUtilizationReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)]
        [datetime]$From,
        [Parameter(Position=1)]
        [datetime]$To,
        [switch]$ExcludeCurrentPeriod
    )

    $work_days = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
    $report = @()
    
    # Normalize $From
    if (-not $From) {
        $From = Get-Date -Day 1
    }
    
    if ($ExcludeCurrentPeriod) {
        $To = (Get-Date -Day 1).AddDays(-1)
    }
    elseif (-not $To) {
        $To = Get-Date
    }
    
    while ($From -le $To) {
        $month_first_day = Get-Date $From -Day 1
        $month_last_day = $month_first_day.AddMonths(1).AddDays(-1)
        $normal_hours = 0
        $total_hours = 0
        $billable_hours = 0
        $utilized_hours = 0
        $training_hours = 0
        $pto_hours = 0
        $holiday_hours = 0
        $non_billable_hours = 0
        $overtime_hours = 0
        
        # Determine period start and end
        $period_start = $From.Day
        if ($To -lt $month_last_day) {
            $period_end = $To.Day
        } else {
            $period_end = $month_last_day.Day
        }
    
        # Calculate normal working hours for the specified pay period
        for ($x=$period_start; $x -le $period_end; $x++) {
            $day = '{0:yyyy-MM}-{1}' -f $From, $x
            $day_of_week = (Get-Date $day).DayOfWeek
            
            if ($day_of_week -in $work_days) {
                $normal_hours += 8
            }
        }
        
        [system.array]$detailed_report = Get-TogglDetailedReport -From ('{0:yyyy-MM}-{1}' -f $From, $period_start) -To ('{0:yyyy-MM}-{1}' -f $From, $period_end)

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
            $percent_billable = ($billable_hours/$normal_hours)*100
            $percent_utilized = (($billable_hours+$utilized_hours)/$normal_hours)*100
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
        
        # Advance to next pay period
        $From = Get-Date $From.AddMonths(1) -Day 1
    }
    
    Write-Output $report
}
